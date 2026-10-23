param(
  [Parameter(Mandatory=$true)][string]$Book,
  [string]$FixtureRoot = "$PSScriptRoot/fixtures/issue50",
  [string]$Out = "$PSScriptRoot/tmp/issue50_resume",
  [ValidateSet('AUTO','OFF')][string]$Native='AUTO',
  [string[]]$Cases=@('default_start_no_detail','explicit_hold','explicit_start','explicit_linear','explicit_blank_mode_hold','water_series')
)
$ErrorActionPreference = 'Stop'
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Another Excel is running; do not dispatch.' }
$root = (Resolve-Path -LiteralPath $FixtureRoot).Path
$destination = [IO.Path]::GetFullPath($Out)
if (Test-Path -LiteralPath $destination) { throw "Preserve previous evidence: $destination" }
New-Item -ItemType Directory -Path $destination | Out-Null
$runner = Join-Path $PSScriptRoot 'run_groundwater_fixture.ps1'
foreach ($name in $Cases) {
  $fixture = Join-Path $root ("issue50_$name.json")
  # C000003 is a mid-WATER checkpoint; stg03_WATER is the completed WATER
  # checkpoint.  Each invocation performs a fresh run before reopening.
  foreach ($resumeName in 'mid','completed') {
    $caseOut = Join-Path $destination ("issue50_${name}_${resumeName}")
    $resumeState = if ($resumeName -eq 'mid') { 'C000003' } else { 'stg03_WATER' }
    & pwsh -NoProfile -File $runner -Book $Book -Fixture $fixture -Out $caseOut -Resume -ResumeState $resumeState -ReopenResume -Flow INCONSISTENT -Native $Native -WaterObservables *> ($caseOut+'.log')
    if ($LASTEXITCODE -ne 0) { throw "Issue #50 case failed: ${name}_${resumeName}" }
    for($i=0;$i -lt 60;$i++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){break};Start-Sleep -Milliseconds 500}
    if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel remains; no further dispatch.'}
  }
}
python -X utf8 (Join-Path $PSScriptRoot 'check_issue50_resume.py') --root $destination --out (Join-Path $destination 'comparison.json') --cases $Cases
if ($LASTEXITCODE -ne 0) { throw 'Issue #50 comparison failed.' }
Write-Output "PASS: $destination"
