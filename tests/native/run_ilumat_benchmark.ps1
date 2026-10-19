param([Parameter(Mandatory=$true)][string]$ReferenceBook,[Parameter(Mandatory=$true)][string]$CandidateBook,[Parameter(Mandatory=$true)][string]$OutputDirectory,[int]$Pairs=3)
$ErrorActionPreference='Stop'
$ReferenceBook=(Resolve-Path -LiteralPath $ReferenceBook).Path
$CandidateBook=(Resolve-Path -LiteralPath $CandidateBook).Path
if((Get-FileHash -LiteralPath $ReferenceBook).Hash -cne (Get-FileHash -LiteralPath $CandidateBook).Hash){throw 'Compare identical compiled VBA books.'}
if($Pairs -lt 1){throw 'Pairs must be positive.'}
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve existing evidence.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
for($pair=1;$pair -le $Pairs;$pair++){
 $order=if($pair%2 -eq 1){@(0,1)}else{@(1,0)}
 foreach($enabled in $order){
  for($wait=0;$wait -lt 20;$wait++){
   if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){break}
   Start-Sleep -Seconds 1
  }
  if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel remains running; no next test dispatched.'}
  $book=if($enabled -eq 0){$ReferenceBook}else{$CandidateBook}
  $out=Join-Path $OutputDirectory ("pair${pair}_enabled${enabled}.json")
  & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'run_native_compare.ps1') -Book $book -Out $out -Suite smokesrm -CaseName Case1 -Variants off -Modes AUTO -Flow INCONSISTENT -Precision AUTO -StateCopy 1 -GmresSimd 1 -IoFast 1 -EventFast 1 -AssemblyFast 1 -GmresTuned 1 -GmresPrepared $enabled
  if($LASTEXITCODE -ne 0){throw "Benchmark failed: $out"}
 }
}
