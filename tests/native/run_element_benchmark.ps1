param([Parameter(Mandatory=$true)][string]$ReferenceBook,[Parameter(Mandatory=$true)][string]$CandidateBook,[Parameter(Mandatory=$true)][string]$OutputDirectory,[int]$Pairs=3,[switch]$Regression,[int]$StartPair=1)
$ErrorActionPreference='Stop'
$ReferenceBook=(Resolve-Path -LiteralPath $ReferenceBook).Path
$CandidateBook=(Resolve-Path -LiteralPath $CandidateBook).Path
if((Get-FileHash -LiteralPath $ReferenceBook).Hash -cne (Get-FileHash -LiteralPath $CandidateBook).Hash){throw 'Compare identical compiled VBA books.'}
if($Pairs -lt 1 -or $StartPair -lt 1){throw 'Pairs and StartPair must be positive.'}
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve existing evidence.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
function RunPair([string]$Name,[string]$Suite,[string]$Flow,[int[]]$Order){
 foreach($enabled in $Order){
  for($wait=0;$wait -lt 20;$wait++){
   if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){break}
   Start-Sleep -Seconds 1
  }
  if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel remains running.'}
  $book=if($enabled -eq 0){$ReferenceBook}else{$CandidateBook}
  $out=Join-Path $OutputDirectory ($Name+"_enabled${enabled}.json")
  $extra=@{}
  if($Suite -eq 'smokesrm'){$extra=@{CaseName='Case1';Variants=@('off')}}
  & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'run_native_compare.ps1') -Book $book -Out $out -Suite $Suite -Modes AUTO -Flow $Flow -Precision AUTO -StateCopy 1 -GmresSimd 1 -IoFast 1 -EventFast 1 -AssemblyFast 1 -GmresTuned 1 -GmresPrepared 1 @extra
  if($LASTEXITCODE -ne 0){throw "Benchmark failed: $out"}
 }
}
if($Regression){
 foreach($flow in 'INCONSISTENT','DAVIS'){RunPair $flow 'smoke' $flow @(0,1)}
 RunPair 'builtin' 'builtin' 'INCONSISTENT' @(0,1)
}else{
 for($pair=$StartPair;$pair -lt $StartPair+$Pairs;$pair++){
  $order=if($pair%2 -eq 1){@(0,1)}else{@(1,0)}
  RunPair "pair$pair" 'smokesrm' 'INCONSISTENT' $order
 }
}
