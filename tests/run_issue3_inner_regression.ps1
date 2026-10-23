param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Out){throw 'Preserve previous evidence.'}
New-Item -ItemType Directory -Path $Out|Out-Null
function WaitExcel {
 for($i=0;$i -lt 60;$i++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){return};Start-Sleep -Milliseconds 500}
 throw 'Excel remains; no further dispatch.'
}
$fixture=Join-Path $PSScriptRoot 'fixtures/fixed_stress/column_inner_io.json'
foreach($native in 'OFF','AUTO'){
 foreach($writer in 'PLCSV','FSCSV','LEDGERCSV','GWCSV'){
  foreach($place in 'TIME','STAGE'){
   WaitExcel
   $point=$writer+'_'+$place;$case=Join-Path $Out ($native+'_'+$point)
   & pwsh -NoProfile -File (Join-Path $PSScriptRoot run_groundwater_fixture.ps1) -Book $Book -Fixture $fixture -Out $case -CancelAt $point -CheckLedger -Native $native *> ($case+'.log')
   if($LASTEXITCODE -ne 0){throw "Nested cancellation failed: $case"}
   Write-Output "PASS $native $point Err18"
  }
 }
}
foreach($writer in 'PLCSV','FSCSV','LEDGERCSV','GWCSV'){
 foreach($place in 'TIME','STAGE'){
  WaitExcel
  $point=$writer+'_'+$place;$case=Join-Path $Out ('ordinary_'+$point)
  $options=@();if($place -eq 'TIME'){$options=@('-ExpectedOrdinaryFailure')}
  & pwsh -NoProfile -File (Join-Path $PSScriptRoot run_groundwater_fixture.ps1) -Book $Book -Fixture $fixture -Out $case -CancelAt $point -CheckLedger -FaultNumber 5 -Native AUTO @options *> ($case+'.log')
  if($LASTEXITCODE -ne 0){throw "Ordinary-error control failed: $case"}
  Write-Output "PASS ordinary $point Err5"
 }
}
