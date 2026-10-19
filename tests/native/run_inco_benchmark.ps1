param(
 [Parameter(Mandatory=$true)][string]$Book,
 [Parameter(Mandatory=$true)][string]$OutputDirectory,
 [ValidateRange(1,10)][int]$Pairs=3,
 [switch]$Resume,
 [switch]$IoOnly,
 [switch]$EventOnly,
 [switch]$AssemblyOnly,
 [switch]$GmresOnly
)
$ErrorActionPreference='Stop'
if(@($IoOnly,$EventOnly,$AssemblyOnly,$GmresOnly | Where-Object {$_}).Count -gt 1){throw 'Select only one comparison scope.'}
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
  $native=if($IoOnly -or $EventOnly -or $AssemblyOnly -or $GmresOnly){1}else{$enabled}
  $io=if($EventOnly -or $AssemblyOnly -or $GmresOnly){1}elseif($IoOnly){$enabled}else{0}
  $event=if($AssemblyOnly -or $GmresOnly){1}elseif($EventOnly){$enabled}else{0}
  $assembly=if($GmresOnly){1}elseif($AssemblyOnly){$enabled}else{0}
  $tuned=if($GmresOnly){$enabled}else{0}
  & pwsh -NoProfile -File (Join-Path $PSScriptRoot 'run_native_compare.ps1') -Book $Book -Out $out -Suite smokesrm -CaseName Case1 -Variants off -Modes AUTO -Flow INCONSISTENT -Precision AUTO -StateCopy $native -GmresSimd $native -IoFast $io -EventFast $event -AssemblyFast $assembly -GmresTuned $tuned
  if($LASTEXITCODE -ne 0){throw "Benchmark failed: $out"}
  # The COM server can exit just after the child PowerShell releases its RCWs.
  for($wait=0;$wait -lt 20;$wait++){
   if(-not (Get-Process EXCEL -ErrorAction SilentlyContinue)){break}
   Start-Sleep -Seconds 1
  }
 }
}
