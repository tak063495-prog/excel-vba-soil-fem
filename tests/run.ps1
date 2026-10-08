$ErrorActionPreference = 'Stop'
$rootTask = Split-Path -Parent $PSScriptRoot
Push-Location $rootTask
try {
  New-Item -ItemType Directory -Force tests/tmp,tests/tmp/modules,tests/results | Out-Null
  $suiteTask = @(
    'test_best03_defaults.ps1',
    'test_best03_unit.ps1',
    'test_best03_faults.ps1',
    'test_practical_smoke.ps1',
    'test_practical_results.ps1',
    'test_practical_ui.ps1',
    'test_practical_final.ps1',
    'test_practical_workflows.ps1'
  )
  foreach ($scriptTask in $suiteTask) {
    Write-Host "Running $scriptTask"
    & (Join-Path $PSScriptRoot $scriptTask)
    if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) { throw "$scriptTask failed with exit code $LASTEXITCODE" }
  }
  Write-Host 'Running test_issue1_regression.ps1'
  & (Join-Path $PSScriptRoot 'test_issue1_regression.ps1') -ExpectFixed
  if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) { throw "Issue #1 regression failed with exit code $LASTEXITCODE" }
  Write-Host 'Practical regression suite completed.'
}
finally {
  Pop-Location
}
