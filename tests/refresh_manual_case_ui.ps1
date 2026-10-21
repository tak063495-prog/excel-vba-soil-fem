param([Parameter(Mandatory=$true)][string]$SourceBook,[Parameter(Mandatory=$true)][string]$DestinationBook,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop';$app=$null;$wb=$null
$src=(Resolve-Path -LiteralPath $SourceBook).Path;$dst=[IO.Path]::GetFullPath($DestinationBook)
if(Test-Path -LiteralPath $dst){throw 'DestinationBook already exists.'};if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
Copy-Item -LiteralPath $src -Destination $dst
try {
  $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
  $wb=$app.Workbooks.Open($dst,0,$false);$prefix="'"+$wb.Name+"'!"
  $mods=@('FEMUiLayout','FEMUiInputs','FEMIo','FEMViewIncrement','FEMIncrementSvg','FEMStrainMath')
  foreach($n in $mods){$cm=$wb.VBProject.VBComponents.Item($n).CodeModule;$cm.DeleteLines(1,$cm.CountOfLines);$cm.AddFromString([IO.File]::ReadAllText((Join-Path (Join-Path (Get-Location) 'src/vba') ($n+'.bas'))))}
  function Snap($name,$cols){$s=$wb.Worksheets.Item($name);$lr=$s.Cells($s.Rows.Count,1).End(-4162).Row;return (($s.Range($s.Cells(1,1),$s.Cells([Math]::Max(2,$lr),$cols)).Value2|ConvertTo-Json -Depth 5 -Compress))}
  $before=@{};foreach($x in @(@('結果節点',12),@('結果要素',72),@('ステージ結果',20))){$before[$x[0]]=Snap $x[0] $x[1]}
  $ret=[string]$app.Run($prefix+'FEMUiRefreshInputs');if($ret -notlike 'PASS*'){throw "refresh failed: $ret"}
  $cmd=$app.VBE.CommandBars.FindControl(1,578);if($cmd.Enabled){$cmd.Execute()};if($app.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Compile failed.'}
  [void]$app.Run($prefix+'SetP0SilentMode',$true);$wb.Worksheets.Item('図').Shapes.Item('P1Agg_ViewSVG').Delete();[void]$app.Run($prefix+'FEMViewResult');$shape=$wb.Worksheets.Item('図').Shapes.Item('P1Agg_ViewSVG');$wb.Save();$wb.Close($false);$wb=$null
  $wb=$app.Workbooks.Open($dst,0,$false);[void]$app.Run("'"+$wb.Name+"'!SetP0SilentMode",$true);$wb.Worksheets.Item('図').Shapes.Item('P1Agg_ViewSVG').Delete();[void]$app.Run("'"+$wb.Name+"'!FEMViewResult");$shape=$wb.Worksheets.Item('図').Shapes.Item('P1Agg_ViewSVG')
  foreach($x in @(@('結果節点',12),@('結果要素',72),@('ステージ結果',20))){if($before[$x[0]] -ne (Snap $x[0] $x[1])){throw "Table changed: $($x[0])"}}
  [pscustomobject]@{status='PASS';destination=$dst;shape=$shape.Name;tables_unchanged=$true}|ConvertTo-Json|Set-Content -LiteralPath $Out -Encoding UTF8
} finally {if($wb){$wb.Close($false)};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
