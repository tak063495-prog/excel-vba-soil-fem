param(
 [Parameter(Mandatory=$true)][string]$Book,
 [Parameter(Mandatory=$true)][string]$OutputDirectory,
 [ValidateRange(1,10)][int]$Pairs=3,
 [switch]$Resume
)
$ErrorActionPreference='Stop'
$Book=(Resolve-Path -LiteralPath $Book).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){
 if(-not $Resume){throw 'Preserve existing evidence; use -Resume to skip completed runs.'}
}else{New-Item -ItemType Directory -Path $OutputDirectory | Out-Null}
for($pair=1;$pair -le $Pairs;$pair++){
 $order=if($pair%2 -eq 1){@(0,1)}else{@(1,0)}
 foreach($enabled in $order){
  $out=Join-Path $OutputDirectory ("pair${pair}_enabled${enabled}.json")
  if($Resume -and (Test-Path -LiteralPath $out)){Write-Output "Preserved $out";continue}
  & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'run_native_compare.ps1') -Book $Book -Out $out -Suite smokesrm -CaseName Case1 -Variants off -Modes AUTO -Flow INCONSISTENT -Precision AUTO -StateCopy $enabled -GmresSimd $enabled
  if($LASTEXITCODE -ne 0){throw "Benchmark failed: $out"}
  # The COM server can exit just after the child PowerShell releases its RCWs.
  for($wait=0;$wait -lt 20;$wait++){
   if(-not (Get-Process EXCEL -ErrorAction SilentlyContinue)){break}
   Start-Sleep -Seconds 1
  }
 }
}
