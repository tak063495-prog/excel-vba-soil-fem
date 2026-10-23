param([string]$Book='workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm',[string]$Out='tests/tmp/issue49_57/targeted',[string]$StartAt='')
$ErrorActionPreference='Stop'
if(Test-Path $Out){throw 'Preserve evidence'}
New-Item -ItemType Directory $Out|Out-Null
foreach($fixture in (Get-ChildItem tests/fixtures/issue49_57/*.json | Where-Object {$_.BaseName -ge $StartAt} | Sort-Object Name)){
 if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel running'}
 $n=$fixture.BaseName;Write-Output "START $n $(Get-Date -Format o)"
 $extra=@()
 if($n -like 'water_*'){$extra=@('-Resume','-ResumeState','stg03_WATER','-ReopenResume')}
 if($n -like 'named_*'){$extra=@('-Resume','-ResumeState','stg03_APPLY','-ReopenResume')}
 & pwsh -NoProfile -File tests/run_groundwater_fixture.ps1 -Book $Book -Fixture $fixture.FullName -Out (Join-Path $Out $n) @extra *> (Join-Path $Out ($n+'.log'))
 if($LASTEXITCODE -ne 0){throw "Failed $n"}
 for($i=0;$i -lt 40;$i++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){break};Start-Sleep -Milliseconds 500}
 Write-Output "DONE $n $(Get-Date -Format o)"
}
$marker=if($StartAt){'subset_done.txt'}else{'all_done.txt'}
Set-Content (Join-Path $Out $marker) PASS
