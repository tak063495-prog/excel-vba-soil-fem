param([Parameter(Mandatory=$true)][string]$Book,[string]$Out='tests/results/native_figure/primitives.json')
$ErrorActionPreference='Stop';$app=$null;$wb=$null
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$src=(Resolve-Path $Book).Path;$dir=Join-Path (Get-Location) ('tests/tmp/figure_primitives_'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $dir|Out-Null;$copy=Join-Path $dir 'test.xlsm';Copy-Item $src $copy
try{
 $app=New-Object -ComObject Excel.Application;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
 $wb=$app.Workbooks.Open($copy,0,$false)
 foreach($n in 'IncrementDisplayHarness','NativeFigureHarness'){$m=$wb.VBProject.VBComponents.Add(1);$m.Name=$n;$m.CodeModule.AddFromString([IO.File]::ReadAllText((Resolve-Path ('tests/native/'+$n+'.bas')).Path))}
 $cmd=$app.VBE.CommandBars.FindControl(1,578);if($cmd.Enabled){$cmd.Execute()};if($app.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Compile failed.'}
 $r=[string]$app.Run("'"+$wb.Name+"'!TNativeFigure");if($r -notlike 'PASS|*'){throw $r}
 @{status='PASS';compile='PASS';result=$r;source_sha256=(Get-FileHash $src).Hash;excel_version=[string]$app.Version;excel_build=[string]$app.Build}|ConvertTo-Json|Set-Content $Out -Encoding UTF8
 Write-Output $r
}finally{if($wb){$wb.Close($false)};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
