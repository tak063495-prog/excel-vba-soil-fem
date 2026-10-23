param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no dispatch.'}
if(Test-Path $Out){throw 'Preserve previous evidence.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory -Path $Out|Out-Null
function WaitExcel {
 for($i=0;$i -lt 60;$i++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){return};Start-Sleep -Milliseconds 500}
 throw 'Excel remains; no further dispatch.'
}
& pwsh -NoProfile -File "$PSScriptRoot/run_issue59_geometry.ps1" -Book $Book -Out (Join-Path $Out probe) *> (Join-Path $Out probe.log)
if($LASTEXITCODE -ne 0){throw 'Geometry probe failed'}
foreach($fixture in (Get-ChildItem "$PSScriptRoot/fixtures/issue59/excel/*.json")){
 WaitExcel
 $case=Join-Path $Out $fixture.BaseName
 & pwsh -NoProfile -File "$PSScriptRoot/run_groundwater_fixture.ps1" -Book $Book -Fixture $fixture.FullName -Out $case *> ($case+'.log')
 if($LASTEXITCODE -ne 0){throw "Input geometry failed: $fixture"}
 Write-Output "PASS $($fixture.BaseName)"
}
WaitExcel
