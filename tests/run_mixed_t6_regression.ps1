param(
 [Parameter(Mandatory=$true)][string]$Book,
 [Parameter(Mandatory=$true)][string]$Out,
 [string]$ReferenceBook=''
)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no test started.'}
if(Test-Path -LiteralPath $Out){throw 'Preserve previous evidence.'}
$Book=(Resolve-Path -LiteralPath $Book).Path
$Out=[IO.Path]::GetFullPath($Out)
if($ReferenceBook){$ReferenceBook=(Resolve-Path -LiteralPath $ReferenceBook).Path}
New-Item -ItemType Directory -Path $Out | Out-Null
function WaitExcel {
 for($attempt=0;$attempt -lt 60;$attempt++){
  if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){return}
  Start-Sleep -Milliseconds 500
 }
 throw 'Excel still running; no duplicate launched.'
}
function RunCase([string]$name,[string]$fixture,[string]$selected,[string]$native='AUTO',[switch]$Resume){
 WaitExcel
 $parameters=@('-NoProfile','-File',(Join-Path $PSScriptRoot 'run_groundwater_fixture.ps1'),'-Book',$selected,'-Fixture',$fixture,'-Out',(Join-Path $Out $name),'-Native',$native)
 if($Resume){$parameters+='-Resume'}
 & pwsh @parameters *> (Join-Path $Out ($name+'.log'))
 if($LASTEXITCODE -ne 0){throw "Failed: $name; see its log."}
 "PASS $name"
}
foreach($name in 'q8_gravity','mixed_gravity','t6_gravity','q8_traction','mixed_traction','t6_traction','q8_seepage','mixed_seepage','t6_seepage'){
 RunCase $name (Join-Path $PSScriptRoot "fixtures/mixed_t6/$name.json") $Book
}
RunCase 'mixed_gravity_off' (Join-Path $PSScriptRoot 'fixtures/mixed_t6/mixed_gravity.json') $Book 'OFF'
RunCase 'mixed_plastic_auto' (Join-Path $PSScriptRoot 'fixtures/mixed_t6/mixed_plastic.json') $Book 'AUTO'
RunCase 'mixed_plastic_off' (Join-Path $PSScriptRoot 'fixtures/mixed_t6/mixed_plastic.json') $Book 'OFF'
foreach($name in 'mixed_joint_water_wet','mixed_joint_water_stages'){
 RunCase $name (Join-Path $PSScriptRoot "fixtures/mixed_t6/$name.json") $Book -Resume
}
if($ReferenceBook){
 RunCase 'q8_reference_srm' (Join-Path $PSScriptRoot 'fixtures/groundwater_slope_low_srm.json') $ReferenceBook
 RunCase 'q8_candidate_srm' (Join-Path $PSScriptRoot 'fixtures/groundwater_slope_low_srm.json') $Book
}
RunCase 'mixed_slope_srm' (Join-Path $PSScriptRoot 'fixtures/mixed_t6/mixed_slope_srm.json') $Book -Resume
python -X utf8 (Join-Path $PSScriptRoot 'check_mixed_t6.py') $Out --out (Join-Path $Out 'independent.json')
if($LASTEXITCODE -ne 0){throw 'Independent check failed.'}
'ALL PASS'
