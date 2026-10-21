param(
 [string]$SourceBook='workbook/ui_20261020/2DSoilFEM_ui_20261021.xlsm',
 [Parameter(Mandatory=$true)][string]$DestinationBook,
 [Parameter(Mandatory=$true)][string]$Record,
 [string]$ViewerSource='src/vba/FEMViewer.bas',
 [switch]$SavedResults,
 [switch]$ButtonActions,
 [switch]$HoldForVisual
)
$ErrorActionPreference='Stop';$app=$null;$wb=$null
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no test started.'}
$source=(Resolve-Path -LiteralPath $SourceBook).Path;$dest=[IO.Path]::GetFullPath($DestinationBook)
if(Test-Path -LiteralPath $dest){throw 'Destination exists; preserve it.'}
New-Item -ItemType Directory -Force (Split-Path $dest)|Out-Null
New-Item -ItemType Directory -Force (Split-Path ([IO.Path]::GetFullPath($Record)))|Out-Null
Copy-Item -LiteralPath $source -Destination $dest
$checks=[Collections.Generic.List[string]]::new();$views=[Collections.Generic.List[object]]::new()
function Check([bool]$ok,[string]$label){if(-not $ok){throw "FAIL: $label"};$checks.Add($label)}
function Snap([string]$name){$ws=$wb.Worksheets.Item($name);$last=[Math]::Max(2,$ws.Cells($ws.Rows.Count,1).End(-4162).Row);return ($ws.Range($ws.Cells(1,1),$ws.Cells($last,80)).Value2|ConvertTo-Json -Depth 6 -Compress)}
function Modules(){ $d=@{};foreach($m in $wb.VBProject.VBComponents){$d[$m.Name]=if($m.CodeModule.CountOfLines){$m.CodeModule.Lines(1,$m.CodeModule.CountOfLines)}else{''}};return $d }
function Compile(){ $cmd=$app.VBE.CommandBars.FindControl(1,578);if($cmd.Enabled){$cmd.Execute()};Check (-not $app.VBE.CommandBars.FindControl(1,578).Enabled) 'Excel full VBA compile' }
function Draw([string]$mode){
 $h=$null;try{$h=$wb.VBProject.VBComponents.Item('FigureProbe')}catch{}
 if(-not $h){$h=$wb.VBProject.VBComponents.Add(1);$h.Name='FigureProbe';$h.CodeModule.AddFromString("Public Function FigureErrors() As String`r`nFigureErrors = FEMViewerLastError & chr(124) & FEMUiMeshPreviewError`r`nEnd Function")}

 $timer=[Diagnostics.Stopwatch]::StartNew();$p="'"+$wb.Name+"'!"
 [void]$app.Run($p+'SetP0SilentMode',$true)
 if($ButtonActions){
  $sheet=if($mode -eq 'MESH'){'操作パネル'}else{'出力・表示'}
  $button=if($mode -eq 'MESH'){'UI_Run'}else{'UI_ResultFigure'}
  $action=[string]$wb.Worksheets.Item($sheet).Shapes.Item($button).OnAction
  Check ($action.StartsWith("'"+$wb.Name.Replace("'","''")+"'!")) "$mode button bound to renamed workbook"
  [void]$app.Run($action)
 }else{[void]$app.Run($p+$(if($mode -eq 'MESH'){'FEMUiPreviewMesh'}else{'FEMViewResult'}))}
 $timer.Stop();$errorText=[string]$app.Run($p+'FigureErrors');if($errorText -cne '|'){throw $errorText};$s=$wb.Worksheets.Item('図').Shapes.Item('P1Agg_ViewSVG')
 Check ($s.Type -eq 6) "$mode native group, no SVG picture decoder"
 Check ($s.GroupItems.Count -gt 10) "$mode has geometry/text primitives"
 Check ($s.AlternativeText -like '*要素*') "$mode element caption"
 Check ($wb.ActiveSheet.Name -ceq '図') "$mode activates figure sheet"
 $views.Add(@{mode=$mode;seconds=$timer.Elapsed.TotalSeconds;type=[int]$s.Type;children=$s.GroupItems.Count;width=[double]$s.Width;height=[double]$s.Height})
}
try{
 $app=New-Object -ComObject Excel.Application;$app.EnableEvents=$true;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $wb=$app.Workbooks.Open($dest,0,$false);$before=Modules
 $tables=@{};foreach($n in '節点データ','要素データ','材料データ','ステージ','載荷','接合','結果節点','結果要素','ステージ結果','結果増分'){if(@($wb.Worksheets|ForEach-Object{$_.Name}) -contains $n){$tables[$n]=Snap $n}}
 $dirty=$wb.Names.Item('_FEMViewerResultsDirty').RefersTo
 foreach($n in 'FEMViewer','FEMNativeFigure'){
  $m=$null;try{$m=$wb.VBProject.VBComponents.Item($n)}catch{}
  if(-not $m){$m=$wb.VBProject.VBComponents.Add(1);$m.Name=$n}
  $cm=$m.CodeModule;if($cm.CountOfLines){$cm.DeleteLines(1,$cm.CountOfLines)}
  $path=if($n -eq 'FEMViewer'){$ViewerSource}else{'src/vba/FEMNativeFigure.bas'}
  $cm.AddFromString([IO.File]::ReadAllText((Resolve-Path -LiteralPath $path).Path))
 }
 Compile
 $after=Modules;foreach($n in $before.Keys){if($n -ne 'FEMViewer'){Check ($after[$n] -ceq $before[$n]) "unchanged VBA $n"}}
 $app.EnableEvents=$true
 Draw 'MESH'
 if($SavedResults){Draw 'RESULT'}
 foreach($n in $tables.Keys){Check ($tables[$n] -ceq (Snap $n)) "unchanged input/result $n"}
 Check ($dirty -ceq $wb.Names.Item('_FEMViewerResultsDirty').RefersTo) 'result freshness marker retained'
 $wb.VBProject.VBComponents.Remove($wb.VBProject.VBComponents.Item('FigureProbe'));Compile
 $wb.Save();$wb.Close($false);$wb=$null
 $wb=$app.Workbooks.Open($dest,0,$false)
 Compile
 foreach($n in $tables.Keys){Check ($tables[$n] -ceq (Snap $n)) "reopen unchanged input/result $n"}
 Draw 'MESH';if($SavedResults){Draw 'RESULT'}
 $wb.VBProject.VBComponents.Remove($wb.VBProject.VBComponents.Item('FigureProbe'));Compile
 $wb.Save()
 $recordData=@{status='PASS';source_sha256=(Get-FileHash $source).Hash;excel_version=[string]$app.Version;excel_build=[string]$app.Build;saved_results=[bool]$SavedResults;checks=$checks.ToArray();views=$views.ToArray()}
 $recordData|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $Record -Encoding UTF8
 if($HoldForVisual){$app.Visible=$true;$app.ScreenUpdating=$true;'ready'|Set-Content ($Record+'.ready');while(-not(Test-Path ($Record+'.close'))){Start-Sleep -Milliseconds 500}}
 Write-Output ('PASS '+$checks.Count+' checks')
}catch{Write-Output $_.Exception.Message;Write-Output $_.ScriptStackTrace;if($HoldForVisual){$app.Visible=$true;$app.VBE.MainWindow.Visible=$true;while(-not(Test-Path ($Record+'.close'))){Start-Sleep -Milliseconds 500}};throw}finally{if($wb){$wb.Close($false)};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)};[GC]::Collect();[GC]::WaitForPendingFinalizers()}

$recordData.book_sha256=(Get-FileHash $dest).Hash
$recordData.viewer_source_sha256=(Get-FileHash $ViewerSource).Hash.ToLowerInvariant()
$recordData.renderer_source_sha256=(Get-FileHash 'src/vba/FEMNativeFigure.bas').Hash.ToLowerInvariant()
$recordData|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $Record -Encoding UTF8
