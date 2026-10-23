param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path $Out){throw 'Preserve prior evidence'}
New-Item -ItemType Directory $Out|Out-Null
function Checked([string]$step){if($LASTEXITCODE -ne 0){throw "Failed $step"};for($i=0;$i -lt 120;$i++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){return};Start-Sleep -Milliseconds 500};throw 'Excel remains; no dispatch'}
$cases=@('issues_xlsm/issue44_joint_soil','issues_xlsm/issue44_joint_prescribed','issues_xlsm/issue44_joint_continuous','issues_xlsm/issue44_joint_conflict',
'issues_xlsm/issue44_internal_head','issues_xlsm/issue44_external_linear','issues_xlsm/issue44_external_layers',
'issues_xlsm/issue45_start_only','issues_xlsm/issue45_hidden_knot','issues_xlsm/issue45_numeric_head','issues_xlsm/issue45_equivalent',
'joint_water_continuous','joint_water_soil_steady','joint_water_prescribed_steady','joint_water_conflicting_heads')
foreach($case in $cases){
 $name=Split-Path $case -Leaf
 Write-Output "START $name $(Get-Date -Format o)"
 $params=@('-NoProfile','-File',"$PSScriptRoot/run_groundwater_fixture.ps1",'-Book',$Book,'-Fixture',"$PSScriptRoot/fixtures/$case.json",'-Out',"$Out/$name")
 if($name -notin 'issue44_joint_continuous','issue44_joint_conflict','issue44_internal_head','issue45_start_only','issue45_hidden_knot','issue45_numeric_head','joint_water_conflicting_heads'){$params+='-Resume'}
 pwsh @params *> "$Out/$name.log"
 Checked $name
}
python -X utf8 "$PSScriptRoot/check_issue44_joint.py" --root $Out --out "$Out/joint_analytic.json"
Checked analytic
python -X utf8 "$PSScriptRoot/check_issue_water_profiles.py" --root $Out --out "$Out/water_analytic.json"
Checked profiles
Set-Content "$Out/PASS.txt" PASS
