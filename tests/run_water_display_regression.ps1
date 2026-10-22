param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel is running; no test started.'}
if(Test-Path $Out){throw 'Preserve previous evidence; choose a new output directory.'}
$Book=(Resolve-Path $Book).Path
$Out=[IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory $Out | Out-Null
function WaitExcel { for($i=0;$i -lt 120;$i++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){return};Start-Sleep -Milliseconds 500};throw 'Excel remains open; no next test started.' }
function Run($script,$params,$name){WaitExcel;& pwsh -NoProfile -File (Join-Path $PSScriptRoot $script) @params *> (Join-Path $Out ($name+'.log'));if($LASTEXITCODE -ne 0){throw "Failed $name"};Add-Content (Join-Path $Out 'queue.log') "PASS $name"}
Run 'run_groundwater.ps1' @('-Book',$Book,'-Out',"$Out/basic") 'basic'
foreach($name in 'cutoff','box_inside','stages'){
 Run 'run_groundwater_fixture.ps1' @('-Book',$Book,'-Fixture',"$PSScriptRoot/fixtures/groundwater_$name.json",'-Out',"$Out/$name",'-Resume') $name
}
Run 'run_groundwater_fixture.ps1' @('-Book',$Book,'-Fixture',"$PSScriptRoot/fixtures/water_stage_one_way.json",'-Out',"$Out/one_way",'-Resume') 'one_way'
foreach($name in 'fixed_stress','time_reserved','missing_stage','conflicting_profile','orphan','duplicate','unknown_profile'){
 Run 'run_groundwater_fixture.ps1' @('-Book',$Book,'-Fixture',"$PSScriptRoot/fixtures/water_stage_guards/$name.json",'-Out',"$Out/reject_$name") "reject_$name"
}
Run 'run_groundwater_fixture.ps1' @('-Book',$Book,'-Fixture',"$PSScriptRoot/fixtures/joint_water_wet.json",'-Out',"$Out/joint",'-Resume') 'joint'
foreach($flow in 'INCONSISTENT','DAVIS'){
 Run 'run_groundwater_fixture.ps1' @('-Book',$Book,'-Fixture',"$PSScriptRoot/fixtures/groundwater_slope_low_srm.json",'-Out',"$Out/wet_srm_$flow",'-Flow',$flow) "wet_srm_$flow"
}
WaitExcel;Add-Content (Join-Path $Out 'queue.log') 'ALL PASS'
