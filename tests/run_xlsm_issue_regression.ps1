param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
$Book=(Resolve-Path $Book).Path
$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path $Out){throw 'Preserve previous evidence; choose a new output directory.'}
New-Item -ItemType Directory $Out|Out-Null
function Checked([string]$case){
 if($LASTEXITCODE -ne 0){throw "Failed: $case"}
 for($i=0;$i -lt 40;$i++){
  if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){return}
  Start-Sleep -Milliseconds 500
 }
 throw 'Excel is still running; no further test dispatched.'
}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no dispatch.'}
$fixtures=Join-Path $PSScriptRoot 'fixtures/issues_xlsm'
$cases=@('issue45_start_only','issue45_hidden_knot','issue45_equivalent','issue45_numeric_head',
 'issue44_internal_head','issue44_external_linear','issue44_external_layers',
 'issue46_reordered','issue46_multiple','issue46_final_reset','issue46_no_reset',
 'issue43_water_ramp_cache','issue43_steady_stages','issue43_water_srm','issue46_final_output','issue26_hold','issue38_elastic_control')
foreach($case in $cases){
 Write-Output "START $case $(Get-Date -Format o)"
 $args=@('-NoProfile','-File',"$PSScriptRoot/run_groundwater_fixture.ps1",'-Book',$Book,
  '-Fixture',"$fixtures/$case.json",'-Out',"$Out/cases/$case")
 if($case -notin 'issue45_start_only','issue45_hidden_knot','issue45_numeric_head','issue44_internal_head'){$args+='-Resume'}
 pwsh @args *> "$Out/$case.log"
 Checked $case
 Write-Output "DONE $case $(Get-Date -Format o)"
}
foreach($case in 'issue46_reordered','issue46_multiple','issue46_final_reset','issue46_no_reset','issue46_final_output'){
 pwsh -NoProfile -File "$PSScriptRoot/run_issue46_viewer.ps1" -Book "$Out/cases/$case/fresh_results.xlsm" -Out "$Out/view/$case" -CheckCoordinateGuard:($case -eq 'issue46_no_reset') *> "$Out/view_$case.log"
 Checked "viewer $case"
}
pwsh -NoProfile -File "$PSScriptRoot/run_issue27_guard.ps1" -Book $Book -Out "$Out/guard" *> "$Out/guard.log"
Checked guard
pwsh -NoProfile -File "$PSScriptRoot/run_issue38_rerun.ps1" -Book "$Out/cases/issue26_hold/fresh_results.xlsm" -Out "$Out/rerun" *> "$Out/rerun.log"
Checked rerun
pwsh -NoProfile -File "$PSScriptRoot/run_issue43_snapshot.ps1" -Book "$Out/cases/issue43_steady_stages/fresh_results.xlsm" -Out "$Out/snapshot_steady" -DirtyProfile W -DeactivateMaterial 2 *> "$Out/snapshot_steady.log"
Checked snapshotSteady
pwsh -NoProfile -File "$PSScriptRoot/run_issue43_snapshot.ps1" -Book "$Out/cases/issue43_water_srm/fresh_results.xlsm" -Out "$Out/snapshot_hydro" -DirtyProfile W *> "$Out/snapshot_hydro.log"
Checked snapshotHydro
python "$PSScriptRoot/check_issue_water_profiles.py" --root "$Out/cases" --out "$Out/water_analytic.json"
if($LASTEXITCODE -ne 0){throw 'Water analytic comparison failed.'}
python "$PSScriptRoot/check_xlsm_issue_states.py" --root $Out --out "$Out/states.json"
if($LASTEXITCODE -ne 0){throw 'Accepted-state/output comparison failed.'}
Write-Output 'PASS: XLSM issue regression completed.'
