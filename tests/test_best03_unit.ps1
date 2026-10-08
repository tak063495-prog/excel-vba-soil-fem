$ErrorActionPreference='Stop'
$rootTask=Split-Path -Parent $PSScriptRoot
Push-Location $rootTask
try {
New-Item -ItemType Directory -Path tests/tmp,tests/results -Force|Out-Null
$taskPath=Join-Path $pwd 'tests/tmp/best03_unit.xlsm'
Copy-Item -LiteralPath 'workbook/2DSoilFEM_20261008_practical.xlsm' -Destination $taskPath -Force
$xlTask=New-Object -ComObject Excel.Application
try {
  $xlTask.EnableEvents=$false; $xlTask.DisplayAlerts=$false; $xlTask.AutomationSecurity=1
  $wbTask=$xlTask.Workbooks.Open($taskPath,0,$false)
  $wbTask.VBProject.VBComponents.Item('FEMAccel').CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $pwd 'tests/fixtures/unit/split_accel_test.txt'),[Text.Encoding]::UTF8))
  $wbTask.VBProject.VBComponents.Item('FEMEngine').CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $pwd 'tests/fixtures/unit/split_engine_test.txt'),[Text.Encoding]::UTF8))
  $wbTask.VBProject.VBComponents.Item('FEMAdaptive').CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $pwd 'tests/fixtures/unit/policy_control_test.txt'),[Text.Encoding]::UTF8))
  $logTestTask=[IO.File]::ReadAllText((Join-Path $pwd 'tests/fixtures/unit/policy_log_test.txt'),[Text.Encoding]::UTF8) -replace '(?m)^Option Explicit\r?\n',''
  $wbTask.VBProject.VBComponents.Item('FEMPolicyLog').CodeModule.AddFromString($logTestTask)
  $resultsTask=@()
  foreach($nameTask in @('PolicyLogSelfTest','AdaptivePolicyV3SelfTest','SplitEligibilitySelfTest','SplitPredictorSelfTest','P6AccelLinearSelfTest')){
    $resultTask=$xlTask.Run("'"+$wbTask.Name+"'!"+$nameTask)
    Write-Output $resultTask
    $resultsTask+=$resultTask
    if(-not $resultTask.StartsWith('PASS')){throw $resultTask}
  }
  $resultsTask | Set-Content 'tests/results/practical_best03_unit_checks_20261008.txt' -Encoding UTF8
  $wbTask.Close($false)
} finally {
  if($null -ne $wbTask){try{$wbTask.Close($false)}catch{}}
  $xlTask.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)|Out-Null
}

} finally { Pop-Location }
