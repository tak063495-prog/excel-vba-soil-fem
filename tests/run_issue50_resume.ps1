param(
  [Parameter(Mandatory=$true)][string]$Book,
  [string]$FixtureRoot = "$PSScriptRoot/fixtures/issue50",
  [string]$Out = "$PSScriptRoot/tmp/issue50_resume"
)
$ErrorActionPreference = 'Stop'
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Another Excel is running; do not dispatch.' }
$root = (Resolve-Path -LiteralPath $FixtureRoot).Path
$destination = [IO.Path]::GetFullPath($Out)
if (Test-Path -LiteralPath $destination) { throw "Preserve previous evidence: $destination" }
New-Item -ItemType Directory -Path $destination | Out-Null
$runner = Join-Path $PSScriptRoot 'run_groundwater_fixture.ps1'
foreach ($name in 'default_start_no_detail','explicit_hold','explicit_start','explicit_linear') {
  $fixture = Join-Path $root ("issue50_$name.json")
  $caseOut = Join-Path $destination ("issue50_$name")
  & pwsh -NoProfile -File $runner -Book $Book -Fixture $fixture -Out $caseOut -Resume -ReopenResume -Flow INCONSISTENT -Native AUTO
  if ($LASTEXITCODE -ne 0) { throw "Issue #50 case failed: $name" }
}
python -X utf8 (Join-Path $PSScriptRoot 'check_issue50_resume.py') --root $destination --out (Join-Path $destination 'comparison.json')
if ($LASTEXITCODE -ne 0) { throw 'Issue #50 comparison failed.' }
Write-Output "PASS: $destination"
