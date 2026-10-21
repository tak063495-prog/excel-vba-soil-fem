param([Parameter(Mandatory=$true)][string]$SourceBook,[Parameter(Mandatory=$true)][string]$DestinationBook,[Parameter(Mandatory=$true)][string]$Out,[switch]$RunRoad)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running. No test started.'}
$repo=(Get-Location).Path;$source=(Resolve-Path -LiteralPath $SourceBook).Path;$destination=[IO.Path]::GetFullPath($DestinationBook)
if(Test-Path -LiteralPath $destination){throw 'Destination exists.'}
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
Copy-Item -LiteralPath $source -Destination $destination
Copy-Item -LiteralPath (Join-Path (Split-Path $source) 'soilfem_native_x64.dll') -Destination (Split-Path $destination) -Force
$checks=[Collections.Generic.List[object]]::new();$app=$null;$wb=$null
function Check($condition,$label){if(-not $condition){throw $label};$checks.Add($label)}
function Snapshot {
 $data=@{};foreach($entry in @(@('材料データ',13),@('節点データ',9),@('要素データ',10),@('載荷',10),@('ステージ',6),@('結果節点',12),@('結果要素',72),@('ステージ結果',20),@('結果増分',20))){
  try{$ws=$wb.Worksheets.Item($entry[0])}catch{continue};$last=$ws.Cells($ws.Rows.Count,1).End(-4162).Row
  $data[$entry[0]]=$ws.Range($ws.Cells(1,1),$ws.Cells([Math]::Max(2,$last),$entry[1])).Value2|ConvertTo-Json -Depth 6 -Compress
 };return $data
}
function Inspect($phase){
 $stage=$wb.Worksheets.Item('ステージ');$load=$wb.Worksheets.Item('載荷')
 Check ([string]$stage.Range('I4').Value2 -match 'DISP=0はゼロ変位拘束' -and [string]$stage.Range('I4').Value2 -match 'FORCE=0はゼロ節点力') "$phase stage legend correct"
 Check ([string]$load.Range('M12').Value2 -match 'DISP=0はゼロ変位' -and [string]$load.Range('M12').Value2 -notmatch 'DISP/FORCE.*無効') "$phase load zero legend correct"
 Check ([string]$load.Range('M2').Value2 -match 'キーワードまたは節点番号' -and [string]$load.Range('M2').Value2 -notmatch 'プルダウン') "$phase target legend correct"
 $target=$load.Range('B2');$saved=$target.Formula
 try{
  $v=$target.Validation;Check ($v.Type -eq 7 -and $v.ShowError -and $v.AlertStyle -eq 1) "$phase custom STOP validation"
  $id=$wb.Worksheets.Item('節点データ').Range('A2').Value2
  foreach($value in @('TOP','BOTTOM','LEFT','RIGHT','BOX',$id)){$target.Formula=[string]$value;Check ([bool]$v.Value) "$phase accepts $value"}
  foreach($value in @(-1.0,0.0,1.5,999999.0)){$target.Formula=[string]$value;Check (-not [bool]$v.Value) "$phase rejects $value"}
 }finally{$target.Formula=$saved}
}
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $wb=$app.Workbooks.Open($destination,0,$false);$prefix="'"+$wb.Name+"'!"
 $moduleBefore=@{};foreach($m in $wb.VBProject.VBComponents){if($m.CodeModule.CountOfLines){$moduleBefore[$m.Name]=$m.CodeModule.Lines(1,$m.CodeModule.CountOfLines)}}
 $baseline=Snapshot
 $cm=$wb.VBProject.VBComponents.Item('FEMIo').CodeModule;$cm.DeleteLines(1,$cm.CountOfLines);$cm.AddFromString([IO.File]::ReadAllText((Join-Path $repo 'src/vba/FEMIo.bas')))
 [void]$app.Run($prefix+'FEMEnsureStageSheets')
 $ctl=$app.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()};Check (-not $app.VBE.CommandBars.FindControl(1,578).Enabled) 'VBE compile'
 Inspect 'open updated copy'
 [void]$app.Run($prefix+'SetP0SilentMode',$true)
 $times=@()
 if($RunRoad){
  $baseline|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path (Split-Path $destination) 'road_baseline_tables.json') -Encoding UTF8
  foreach($run in 1,2){
   [void]$app.Run($prefix+'FEMInvalidateInputCache');[void]$app.Run($prefix+'FEMInvalidateSettingCache')
   $timer=[Diagnostics.Stopwatch]::StartNew();[void]$app.Run($prefix+'P0_RunAnalysis');$timer.Stop();$times+=,$timer.Elapsed.TotalSeconds
   Check ([string]$wb.Worksheets.Item('診断').Range('B2').Value2 -ceq 'PASS') "analysis $run PASS"
   Inspect "analysis $run"
   $actual=Snapshot;foreach($key in $baseline.Keys){Check ($baseline[$key] -ceq $actual[$key]) "analysis $run unchanged $key"}
   $actual|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path (Split-Path $destination) "road_run${run}_tables.json") -Encoding UTF8
   $logs=Join-Path (Split-Path $destination) "run${run}_logs";New-Item -ItemType Directory -Path $logs|Out-Null
   $outdir=Join-Path (Split-Path $destination) (([IO.Path]::GetFileNameWithoutExtension($destination))+'_out')
   Get-ChildItem -LiteralPath $outdir -File|Where-Object {$_.Extension -in '.csv','.log','.txt'}|Copy-Item -Destination $logs
  }
 }
 $wb.Save();$wb.Close($false);$wb=$null
 $wb=$app.Workbooks.Open($destination,0,$false);$prefix="'"+$wb.Name+"'!";Inspect 'save reopen before macro'
 [void]$app.Run($prefix+'FEMEnsureStageSheets');Inspect 'save reopen after regeneration'
 $actual=Snapshot;foreach($key in $baseline.Keys){Check ($baseline[$key] -ceq $actual[$key]) "reopen unchanged $key"}
 foreach($m in $wb.VBProject.VBComponents){
  if($m.Name -eq 'FEMIo'){$expected=[IO.File]::ReadAllText((Join-Path $repo 'src/vba/FEMIo.bas'))}
  elseif($moduleBefore.ContainsKey($m.Name)){$expected=$moduleBefore[$m.Name]}else{continue}
  Check (($m.CodeModule.Lines(1,$m.CodeModule.CountOfLines).Trim() -replace "`r",'') -eq ($expected.Trim() -replace "`r",'')) ('embedded module matches '+$m.Name)
 }
 # Empty-stage defaults are tested only in memory, then discarded.
 $wb.Worksheets.Item('ステージ').Range('A2:F2000').ClearContents();[void]$app.Run($prefix+'FEMEnsureStageSheets')
 Check ([string]$wb.Worksheets.Item('ステージ').Range('F3').Value2 -match 'DISP=0はゼロ変位拘束') 'empty stage default has correct zero DISP meaning'
 $wb.Close($false);$wb=$null
 Check ((Get-FileHash -LiteralPath $source).Hash -ceq $sourceHash) 'source workbook unchanged'
 [pscustomobject]@{status='PASS';checks=$checks;analysis_seconds=$times;analyses=if($RunRoad){2}else{0};matched_tables=@($baseline.Keys);numeric_comparison='all available input/result tables exactly unchanged';destination_sha256=(Get-FileHash -LiteralPath $destination).Hash.ToLower();scope='programmatic validation; no physical keyboard user test; only FEMIo wording changed; embedded source comparison ignores VBE identifier casing'}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $Out -Encoding UTF8
 Write-Output "PASS $($checks.Count) checks; same-session analyses=$($times.Count)"
}finally{if($wb){$wb.Close($false)};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
