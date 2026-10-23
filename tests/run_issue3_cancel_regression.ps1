param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Out){throw 'Preserve previous evidence.'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no dispatch.'}
New-Item -ItemType Directory -Path $Out|Out-Null
$fixture=Join-Path $PSScriptRoot 'fixtures/fixed_stress/column_smoke.json'
function WaitExcel {
 for($i=0;$i -lt 60;$i++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){return};Start-Sleep -Milliseconds 500}
 throw 'Excel remains; no further dispatch.'
}
foreach($native in 'OFF','AUTO'){
 WaitExcel
 $first=Join-Path $Out ($native+'_first_FS')
 & pwsh -NoProfile -File (Join-Path $PSScriptRoot run_groundwater_fixture.ps1) -Book $Book -Fixture $fixture -Out $first -CancelAt FS -CheckLedger -CancelAfterAccepted 0 -Native $native *> ($first+'.log')
 if($LASTEXITCODE -ne 0){throw "First-step cancellation failed: $first"}
 Write-Output "PASS $native first FS"
 foreach($point in 'FS','FLAG','CAPTURE','SAVE','PLASTIC','WRITE_TIME','WRITE_STAGE'){
  WaitExcel
  $case=Join-Path $Out ($native+'_'+$point)
  & pwsh -NoProfile -File (Join-Path $PSScriptRoot run_groundwater_fixture.ps1) -Book $Book -Fixture $fixture -Out $case -CancelAt $point -CheckLedger -Native $native *> ($case+'.log')
  if($LASTEXITCODE -ne 0){throw "Cancellation regression failed: $case"}
  Write-Output "PASS $native $point"
 }
}
foreach($point in 'FS','CAPTURE','SAVE','PLASTIC'){
 WaitExcel
 $case=Join-Path $Out ('ordinary_'+$point)
 & pwsh -NoProfile -File (Join-Path $PSScriptRoot run_groundwater_fixture.ps1) -Book $Book -Fixture $fixture -Out $case -CancelAt $point -CheckLedger -FaultNumber 5 *> ($case+'.log')
 if($LASTEXITCODE -ne 0){throw "Ordinary error regression failed: $case"}
 Write-Output "PASS ordinary $point"
}
