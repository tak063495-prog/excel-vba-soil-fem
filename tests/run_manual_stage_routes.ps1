param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$InputJson,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop';$xl=$null;$wb=$null
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
if(Test-Path -LiteralPath $Out){throw 'Preserve existing evidence.'}
$source=(Resolve-Path -LiteralPath $Book).Path;$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$data=Get-Content -LiteralPath $InputJson -Raw -Encoding UTF8 | ConvertFrom-Json -AsHashtable
New-Item -ItemType Directory -Path $Out | Out-Null;$Out=(Resolve-Path -LiteralPath $Out).Path
$copy=Join-Path $Out 'stage_routes.xlsm';Copy-Item -LiteralPath $source -Destination $copy
Copy-Item -LiteralPath (Join-Path (Split-Path $source) 'soilfem_native_x64.dll') -Destination $Out
try{
 $xl=New-Object -ComObject Excel.Application;$xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
 [void]$xl.Run($prefix+'SetP0SilentMode',$true)
 $code=@{};foreach($m in $wb.VBProject.VBComponents){if($m.CodeModule.CountOfLines){$code[$m.Name]=$m.CodeModule.Lines(1,$m.CodeModule.CountOfLines)}}
 foreach($name in @('材料データ','節点データ','要素データ','ステージ','載荷','接合','結果節点','結果要素','結果塑性','ステージ結果','結果増分')){
  try{$s=$wb.Worksheets.Item($name)}catch{if($name -in @('結果塑性','結果増分')){continue};throw}
  $s.Unprotect();$s.Range($s.Cells(2,1),$s.Cells([Math]::Max(2000,$s.UsedRange.Rows.Count),[Math]::Max(30,$s.UsedRange.Columns.Count))).ClearContents()
 }
 foreach($name in $data.Keys){
  $rows=$data[$name];$s=$wb.Worksheets.Item($name);$s.Unprotect();$cols=$rows[0].Count
  $v=[object[,]]::new($rows.Count,$cols)
  for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $cols;$c++){$v[$r,$c]=$rows[$r][$c]}}
  $s.Range($s.Cells(1,1),$s.Cells($rows.Count,$cols)).Value2=$v
 }
 $wb.Worksheets.Item('形状入力').Range('B8:D2000').ClearContents();$wb.Worksheets.Item('形状入力').Range('G8:J2000').ClearContents()
 $wb.Worksheets.Item('水条件').Range('B3').Value2='NONE'
 [void]$xl.Run($prefix+'FEMInvalidateInputCache');[void]$xl.Run($prefix+'FEMInvalidateSettingCache')
 # Follow the public seven-screen route, including the Worksheet_Change handlers.
 [void]$xl.Run($prefix+'FEM24Analysis');$s=$wb.Worksheets.Item('解析・保存');$xl.EnableEvents=$true
 $s.Range('C4').Value2='保存・再開'
 if($s.Range('A8').Text -notmatch '結果の残し方'){throw ('Unexpected save field: '+$s.Range('A8').Text)}
 $s.Range('D8').Value2='ALL';$xl.EnableEvents=$false
 $settings=$wb.Worksheets.Item('設定');$found=$settings.Columns.Item(5).Find('OUTPUT_STAGE_MODE')
 if($null -eq $found){throw 'Missing stage output setting.'}
 if($settings.Cells($found.Row,3).Value2 -ne 'ALL'){throw 'ALL did not reach canonical settings.'}
 $cmd=$xl.VBE.CommandBars.FindControl(1,578);if($cmd.Enabled){$cmd.Execute()};if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Compile failed.'}
 $wb.Save();$timer=[Diagnostics.Stopwatch]::StartNew();[void]$xl.Run($prefix+'P0_RunAnalysis');$timer.Stop()
 $status=[string]$wb.Worksheets.Item('診断').Cells(2,2).Value2;if($status -ne 'PASS'){throw ('Analysis: '+$status)}
 function Rows($matrix){$list=[Collections.Generic.List[object]]::new();for($r=1;$r -le $matrix.GetLength(0);$r++){$row=@();for($c=1;$c -le $matrix.GetLength(1);$c++){$row+=,$matrix.GetValue($r,$c)};$list.Add($row)};return ,$list.ToArray()}
 $tables=@{};$routes=@()
 foreach($pair in @(@('節点結果','結果節点',12),@('要素結果','結果要素',72))){
  [void]$xl.Run($prefix+'FEM24Results');$s=$wb.Worksheets.Item('結果');$xl.EnableEvents=$true;$s.Range('C4').Value2=[string]$pair[0];$xl.EnableEvents=$false
  [void]$xl.Run($prefix+'FEM24Detail');if($xl.ActiveSheet.Name -ne $pair[1]){throw 'Wrong original result table.'}
  $s=$xl.ActiveSheet;$last=$s.Cells($s.Rows.Count,1).End(-4162).Row
  $tables[$pair[1]]=Rows $s.Range($s.Cells(1,1),$s.Cells($last,[int]$pair[2])).Value2
  $ids=@($tables[$pair[1]] | Select-Object -Skip 1 | ForEach-Object {$_[0]} | Sort-Object -Unique)
  $routes+=@{section=$pair[0];original=$s.Name;stage_ids=$ids;rows=$last-1;header=$tables[$pair[1]][0]}
 }
 $s=$wb.Worksheets.Item('ステージ結果');$last=$s.Cells($s.Rows.Count,1).End(-4162).Row;$tables['ステージ結果']=Rows $s.Range('A1:T'+$last).Value2
 $tables|ConvertTo-Json -Depth 14|Set-Content -LiteralPath (Join-Path $Out 'tables.json') -Encoding UTF8
 # Verify actual A/B column filtering in Excel; do not infer it from the in-memory tables.
 $s=$wb.Worksheets.Item('結果節点');$s.Unprotect();$last=$s.Cells($s.Rows.Count,1).End(-4162).Row;$range=$s.Range('A1:L'+$last);$filters=@()
 if($s.FilterMode){$s.ShowAllData()};$s.AutoFilterMode=$false;$range.EntireRow.Hidden=$false
 $s.Range('A:B').EntireColumn.AutoFit()
 $stages=if($data['節点データ'].Count -eq 329){@(1,4,6,11)}else{@(1..9)}
 $point=if($data['節点データ'].Count -eq 329){19}else{351}
 foreach($id in $stages){
  # Excel's checkbox filter uses the displayed number, including its format.
  $match=0;for($i=1;$i -lt $tables['結果節点'].Count;$i++){if($tables['結果節点'][$i][0] -eq $id -and $tables['結果節点'][$i][1] -eq $point){$match=$i+1;break}}
  if($match -eq 0){throw 'Requested stage/point missing.'}
  $criteriaA='='+$s.Cells($match,1).Text.Trim();$criteriaB='='+$s.Cells($match,2).Text.Trim()
  [void]$range.AutoFilter(1,$criteriaA);[void]$range.AutoFilter(2,$criteriaB)
  $visible=$s.Range('A2:A'+$last).SpecialCells(12)
  if($visible.CountLarge -ne 1){throw ('A/B filter ambiguous '+$id)}
  $r=$visible.Row;if($r -ne $match){throw 'Wrong filtered result.'}
  $filters+=@{stage=$id;node=$point;display_criteria=@($criteriaA,$criteriaB);values_A_B_F_J=@($s.Cells($r,1).Value2,$s.Cells($r,2).Value2,$s.Cells($r,6).Value2,$s.Cells($r,10).Value2)}
 }
 $s.ShowAllData();$s.AutoFilterMode=$false
 foreach($m in $wb.VBProject.VBComponents){if($code.ContainsKey($m.Name) -and $code[$m.Name] -cne $m.CodeModule.Lines(1,$m.CodeModule.CountOfLines)){throw ('Code changed '+$m.Name)}}
 $wb.Save();$wb.Close($false);$wb=$null
 if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash){throw 'Distribution changed.'}
 @{status='PASS';seconds=$timer.Elapsed.TotalSeconds;source_sha256=$sourceHash;input=$InputJson;compile='PASS';application_code_unchanged=$true;save_route='解析・保存 C4=保存・再開 D8=ALL';routes=$routes;filters=$filters;tables=$tables}|ConvertTo-Json -Depth 15|Set-Content -LiteralPath (Join-Path $Out 'result.json') -Encoding UTF8
 Write-Output ('PASS '+$timer.Elapsed.TotalSeconds)
}catch{($_ | Out-String)+$_.ScriptStackTrace | Set-Content -LiteralPath (Join-Path $Out 'failure.log');throw}finally{if($wb){$wb.Close($false)};if($xl){$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
