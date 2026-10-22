param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$ReferenceBook,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
$Book=(Resolve-Path -LiteralPath $Book).Path;$ReferenceBook=(Resolve-Path -LiteralPath $ReferenceBook).Path
$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path -LiteralPath $Out){throw 'Preserve prior evidence.'}
New-Item -ItemType Directory -Path $Out | Out-Null
function WaitExcel {for($i=0;$i -lt 40;$i++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){return};Start-Sleep -Milliseconds 500};throw 'Excel is running; no further test dispatched.'}
function Checked {if($LASTEXITCODE -ne 0){throw "Driver exit $LASTEXITCODE"};WaitExcel}
WaitExcel
pwsh -NoProfile -File "$PSScriptRoot/run_groundwater.ps1" -Book $Book -Out "$Out/basic" *> "$Out/basic.log";Checked
foreach($case in 'box','cutoff','stages','box_inside','separated','initial_birth','all_birth'){
 $args=@('-NoProfile','-File',"$PSScriptRoot/run_groundwater_fixture.ps1",'-Book',$Book,'-Fixture',"$PSScriptRoot/fixtures/groundwater_$case.json",'-Out',"$Out/$case")
 if($case -in 'box','stages','initial_birth','all_birth'){$args+='-Resume'}
 pwsh @args *> "$Out/$case.log";Checked
}
foreach($flow in 'INCONSISTENT','DAVIS'){
 foreach($role in 'reference','candidate'){
  $selected=if($role -eq 'reference'){$ReferenceBook}else{$Book}
  foreach($suite in 'smoke','smokesrm'){
   $args=@('-NoProfile','-File',"$PSScriptRoot/native/run_native_compare.ps1",'-Book',$selected,'-Out',"$Out/dry_${role}_${flow}_${suite}.json",'-Suite',$suite,'-Modes','AUTO','-Flow',$flow,'-Variants','off','-Precision','AUTO','-StateCopy',1,'-GmresSimd',1,'-IoFast',1,'-EventFast',1,'-AssemblyFast',1,'-GmresTuned',1,'-GmresPrepared',1)
   if($suite -eq 'smokesrm'){$args+=@('-CaseName','Case2')}
   pwsh @args *> "$Out/dry_${role}_${flow}_${suite}.log";Checked
  }
 }
}
foreach($native in 'OFF','AUTO'){
 pwsh -NoProfile -File "$PSScriptRoot/run_groundwater_fixture.ps1" -Book $Book -Fixture "$PSScriptRoot/fixtures/groundwater_slope_fixed.json" -Out "$Out/wet_fixed_$native" -Native $native *> "$Out/wet_fixed_$native.log";Checked
}
pwsh -NoProfile -File "$PSScriptRoot/run_groundwater_fixture.ps1" -Book $Book -Fixture "$PSScriptRoot/fixtures/groundwater_slope_low_srm.json" -Out "$Out/wet_srm" -Resume *> "$Out/wet_srm.log";Checked
pwsh -NoProfile -File "$PSScriptRoot/test_native_figure.ps1" -Book $Book -Out "$Out/figure.json" *> "$Out/figure.log";Checked
python "$PSScriptRoot/check_groundwater.py" --basic "$Out/basic" --box "$Out/box" --cutoff "$Out/cutoff" --out "$Out/independent.json"
if($LASTEXITCODE -ne 0){throw 'Independent comparison failed.'}
python "$PSScriptRoot/check_groundwater.py" --cutoff "$Out/separated" --fixture "$PSScriptRoot/fixtures/groundwater_separated.json" --out "$Out/separated_independent.json"
if($LASTEXITCODE -ne 0){throw 'Separated domains comparison failed.'}
python "$PSScriptRoot/compare_groundwater_regression.py" $Out
if($LASTEXITCODE -ne 0){throw 'Complete vector/Fs/residual comparison failed.'}
