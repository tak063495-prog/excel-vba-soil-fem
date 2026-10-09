param(
 [string]$SourceWorkbook=(Join-Path $PSScriptRoot '../../workbook/2DSoilFEM_20261008_practical.xlsm'),
 [string]$OutputRoot=(Join-Path $PSScriptRoot '../tmp/inconsistent'),
 [string]$PythonPath='python',
 [switch]$SkipSpectral
)
$ErrorActionPreference='Stop'
$checksTask=@('check_nonsymmetric_solver','check_regularized_solver','check_failure_classification','check_finalization','check_watchdog','check_fd_guard')
if(-not $SkipSpectral){$checksTask=@('check_spectral_points','check_material_substeps')+$checksTask}
foreach($checkTask in $checksTask){
 Write-Host ('Running '+$checkTask)
 $argsTask=@{SourceWorkbook=$SourceWorkbook;OutputRoot=(Join-Path $OutputRoot $checkTask)}
 if($checkTask -in @('check_spectral_points','check_material_substeps')){$argsTask.PythonPath=$PythonPath}
 & (Join-Path $PSScriptRoot ($checkTask+'.ps1')) @argsTask
 if($LASTEXITCODE -and $LASTEXITCODE -ne 0){throw ($checkTask+' failed with exit code '+$LASTEXITCODE)}
}
Write-Host 'INCONSISTENT material, linear, recovery and failure checks completed.'
