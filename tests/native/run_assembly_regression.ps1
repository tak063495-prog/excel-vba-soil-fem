param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$Book=(Resolve-Path -LiteralPath $Book).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve prior evidence.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
function WaitOwnedExcelExit {
 for($i=0;$i -lt 20;$i++){
  if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){return}
  Start-Sleep -Seconds 1
 }
 throw 'Excel remains running; no next test dispatched.'
}
foreach($flow in 'INCONSISTENT','DAVIS'){
 foreach($enabled in 0,1){
  WaitOwnedExcelExit
  $out=Join-Path $OutputDirectory ($flow+'_'+$enabled+'.json')
  & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'run_native_compare.ps1') -Book $Book -Out $out -Suite smoke -Modes AUTO -Flow $flow -Precision AUTO -StateCopy 1 -GmresSimd 1 -IoFast 1 -EventFast 1 -AssemblyFast $enabled
  if($LASTEXITCODE -ne 0){throw "Regression failed: $out"}
 }
 & python (Join-Path $PSScriptRoot 'summarize_io_regression.py') (Join-Path $OutputDirectory ($flow+'_0.json')) (Join-Path $OutputDirectory ($flow+'_1.json')) --scope assembly --out (Join-Path $OutputDirectory ($flow+'_summary.json'))
 if($LASTEXITCODE -ne 0){throw 'Numerical/log mismatch.'}
}
foreach($enabled in 0,1){
 WaitOwnedExcelExit
 $out=Join-Path $OutputDirectory ('builtin_'+$enabled+'.json')
 & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'run_native_compare.ps1') -Book $Book -Out $out -Suite builtin -Modes AUTO -Flow INCONSISTENT -Precision AUTO -FixedFs 1 -StateCopy 1 -GmresSimd 1 -IoFast 1 -EventFast 1 -AssemblyFast $enabled
 if($LASTEXITCODE -ne 0){throw "Regression failed: $out"}
}
& python (Join-Path $PSScriptRoot 'summarize_io_regression.py') (Join-Path $OutputDirectory 'builtin_0.json') (Join-Path $OutputDirectory 'builtin_1.json') --scope assembly --out (Join-Path $OutputDirectory 'builtin_summary.json')
if($LASTEXITCODE -ne 0){throw 'Builtin numerical/log mismatch.'}
