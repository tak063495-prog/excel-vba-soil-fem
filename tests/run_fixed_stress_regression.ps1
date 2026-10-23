param(
    [Parameter(Mandatory = $true)] [string] $Book,
    [Parameter(Mandatory = $true)] [string] $Out
)

$ErrorActionPreference = 'Stop'
$runner = Join-Path $PSScriptRoot 'run_groundwater_fixture.ps1'
$fixtureRoot = Join-Path $PSScriptRoot 'fixtures\fixed_stress'

if (-not (Test-Path -LiteralPath $Book -PathType Leaf)) { throw "Book not found: $Book" }
if (Test-Path -LiteralPath $Out) { throw "Preserve previous evidence: $Out" }
New-Item -ItemType Directory -Path $Out | Out-Null

$fixtures = @(
    'column_smoke.json',
    'column_biot_anisotropic.json',
    'column_nonlinear_srm.json',
    'nonlinear_control.json',
    'fill_30day.json',
    'watercurve.json',
    'water_external_linear.json',
    'closed.json',
    't6.json',
    'joint_transient.json',
    'wet_birth_ramp.json',
    'named_unload.json',
    'saturation_rejection.json',
    'srm_time_rejection.json',
    'pressure_gauge.json',
    'rollback_instant.json',
    'saturated_slope_srm.json'
)

foreach ($fixture in $fixtures) {
    if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) { throw 'Another Excel is running; do not dispatch.' }
    $fixturePath = Join-Path $fixtureRoot $fixture
    if (-not (Test-Path -LiteralPath $fixturePath -PathType Leaf)) { throw "Fixture not found: $fixturePath" }
    $caseName = [IO.Path]::GetFileNameWithoutExtension($fixture)
    $caseOut = Join-Path $Out $caseName
    Write-Output "START $caseName $(Get-Date -Format o)"
    pwsh -NoProfile -File $runner -Book $Book -Fixture $fixturePath -Out $caseOut -Resume *> ($caseOut+'.log')
    if ($LASTEXITCODE -ne 0) { throw "Fixture runner failed ($LASTEXITCODE): $fixture" }
    for($i=0;$i -lt 40;$i++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){break};Start-Sleep -Milliseconds 500}
    if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw "Excel remains after $fixture; no further dispatch"}
    Write-Output "DONE $caseName $(Get-Date -Format o)"
}
