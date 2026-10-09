param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputRoot|Out-Null
$pathTask=Join-Path $OutputRoot 'srm_roundoff_scratch.xlsm'
Copy-Item -LiteralPath $SourceWorkbook -Destination $pathTask -Force
$codeTask=@'
Public Function LiteratureBracketRoundoff() As Variant
  Dim lower As Double, upper As Double, tol As Double, mid As Double
  V3CostSearchEnabled = False
  tol = 0.0125
  upper = 0.9: lower = upper - 0.1
  upper = NextFsByBracket(lower, upper, tol)
  lower = NextFsByBracket(lower, upper, tol)
  upper = NextFsByBracket(lower, upper, tol)
  mid = NextFsByBracket(lower, upper, tol)
  LiteratureBracketRoundoff = Array(lower, upper, upper - lower, tol, mid > lower, mid, (upper - lower) > tol)
End Function
'@
$xlTask=New-Object -ComObject Excel.Application
$xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1
$wbTask=$null
try{
  $wbTask=$xlTask.Workbooks.Open([IO.Path]::GetFullPath($pathTask),0,$false)
  # Add a wrapper in the same module to call the actual private bracket helper.
  # The workbook is closed without saving; no production code is changed.
  $engineTask=$wbTask.VBProject.VBComponents.Item('FEMEngine').CodeModule
  $engineTask.AddFromString($codeTask)
  $valuesTask=@($xlTask.Run("'"+$wbTask.Name+"'!LiteratureBracketRoundoff"))
  if([bool]$valuesTask[4]){throw 'Roundoff caused an unnecessary extra trial'}
  @{value_names=@('pass','fail','width','tolerance','extra_trial','next_fs','strict_width_exceeds_tolerance');values=$valuesTask;source_sha256=(Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $OutputRoot 'srm_roundoff.json') -Encoding UTF8
}finally{
  if($null -ne $wbTask){$wbTask.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbTask)}
  $xlTask.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
if((Get-FileHash -LiteralPath $pathTask -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash){throw 'Scratch workbook was unexpectedly saved'}
Get-Content -LiteralPath (Join-Path $OutputRoot 'srm_roundoff.json') -Encoding UTF8
