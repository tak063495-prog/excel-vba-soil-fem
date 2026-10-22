param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$ReferenceBook,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
$Book=(Resolve-Path $Book).Path;$ReferenceBook=(Resolve-Path $ReferenceBook).Path
$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path $Out){throw 'Preserve previous evidence.'}
New-Item -ItemType Directory -Path $Out | Out-Null
function WaitExcel {for($i=0;$i -lt 40;$i++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){return};Start-Sleep -Milliseconds 500};throw 'Excel is running; no further test dispatched.'}
function Checked {if($LASTEXITCODE -ne 0){throw "Driver exit $LASTEXITCODE"};WaitExcel}
WaitExcel
foreach($case in 'dry','wet','prescribed','continuous','prescribed_steady','soil_steady','partial','slip','opening','stages','srm','struct_pair'){
 $args=@('-NoProfile','-File',"$PSScriptRoot/run_groundwater_fixture.ps1",'-Book',$Book,'-Fixture',"$PSScriptRoot/fixtures/joint_water_$case.json",'-Out',"$Out/$case")
 if($case -ne 'dry'){$args+='-Resume'}
 pwsh @args *> "$Out/$case.log";Checked
 python "$PSScriptRoot/check_joint_water.py" --fixture "$PSScriptRoot/fixtures/joint_water_$case.json" --run "$Out/$case" --report "$Out/$case/check.json"
 if($LASTEXITCODE -ne 0){throw "Independent check failed: $case"}
}
pwsh -NoProfile -File "$PSScriptRoot/run_groundwater_fixture.ps1" -Book $ReferenceBook -Fixture "$PSScriptRoot/fixtures/joint_water_dry.json" -Out "$Out/dry_reference" *> "$Out/dry_reference.log";Checked
python "$PSScriptRoot/check_joint_water.py" --compare "$Out/dry_reference" "$Out/dry" --report "$Out/dry_comparison.json"
if($LASTEXITCODE -ne 0){throw 'Dry regression failed.'}
foreach($case in 'missing_mode','inconsistent_pressure','conflicting_heads','wrong_soil_mode'){
 pwsh -NoProfile -File "$PSScriptRoot/run_groundwater_fixture.ps1" -Book $Book -Fixture "$PSScriptRoot/fixtures/joint_water_$case.json" -Out "$Out/$case" *> "$Out/$case.log";Checked
}

foreach($case in 'wet','slip'){
 pwsh -NoProfile -File "$PSScriptRoot/run_groundwater_fixture.ps1" -Book $Book -Fixture "$PSScriptRoot/fixtures/joint_water_$case.json" -Out "$Out/${case}_davis" -Flow DAVIS -Resume *> "$Out/${case}_davis.log";Checked
 python "$PSScriptRoot/check_joint_water.py" --fixture "$PSScriptRoot/fixtures/joint_water_$case.json" --run "$Out/${case}_davis" --report "$Out/${case}_davis/check.json"
 if($LASTEXITCODE -ne 0){throw "Wet DAVIS check failed: $case"}
}
