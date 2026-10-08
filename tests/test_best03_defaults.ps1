$ErrorActionPreference='Stop'
$rootTask=Split-Path -Parent $PSScriptRoot
Push-Location $rootTask
try {
New-Item -ItemType Directory -Path tests/tmp,tests/results -Force|Out-Null
$targetTask=Join-Path $pwd 'tests/tmp/best03_defaults.xlsm'
Copy-Item -LiteralPath workbook/2DSoilFEM_20261008_practical.xlsm -Destination $targetTask -Force
$mappingTask=Get-Content tests/fixtures/ui_mapping.json -Raw -Encoding UTF8|ConvertFrom-Json
$xlTask=New-Object -ComObject Excel.Application
try{
  $xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1
  $wbTask=$xlTask.Workbooks.Open($targetTask,0,$false)
  $prefixTask="'"+$wbTask.Name+"'!"
  $testTask=$wbTask.VBProject.VBComponents.Add(1);$testTask.Name='BestDefaultHarness'
  $testTask.CodeModule.AddFromString(@'
Public Function BestDefaultState() As String
  FEMInvalidateSettingCache
  AdaptiveResetRun
  BestDefaultState = CStr(AdaptiveEnabled) & "|" & CStr(StepRecoveryEnabled)
End Function
'@)
  $wbTask.VBProject.VBComponents.Item('FEMIo').CodeModule.AddFromString(@'
Public Function BestTestSeedRecovery() As Double
  Dim ws As Worksheet, r As Long
  Set ws = ThisWorkbook.Worksheets("設定")
  FEMRebuildSettingsLayout ws
  For r = 5 To ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    If CStr(ws.Cells(r, 5).Value2) = "ACCEL_STEP_RECOVERY" Then
      BestTestSeedRecovery = CDbl(ws.Cells(r, 3).Value2)
      Exit Function
    End If
  Next r
  Err.Raise vbObjectError + 3975, , "Recovery setting not seeded"
End Function
'@)
  $checksTask=[Collections.Generic.List[string]]::new()
  if($xlTask.Run($prefixTask+'BestDefaultState') -ne 'True|False'){throw 'delivered defaults invalid'}
  $checksTask.Add('PASS delivered preset: adaptive ON, recovery OFF')
  if($wbTask.Names.Item('_FEMViewerResultsDirty').RefersTo -ne '=TRUE'){throw 'inherited results not marked stale'}
  $checksTask.Add('PASS inherited results require re-analysis under BEST_03')
  foreach($expectedTask in @(@('ACCEL_V1_PREDICTOR',1),@('ACCEL_V2A_REUSE',1),@('ACCEL_V4_ANDERSON',1),@('ACCEL_V2B_COST',0),@('ACCEL_V3_COST_SEARCH',0),@('ACCEL_V5_GMRES_LU',0),@('ACCEL_STEP_FRESH_LU',0),@('ACCEL_TRACE',0))){
    $mTask=$mappingTask|Where-Object {$_.key -eq $expectedTask[0]}
    if($wbTask.Worksheets.Item($mTask.sheet).Range($mTask.cell).Value2 -ne $expectedTask[1]){throw ('preset '+$expectedTask[0])}
    $checksTask.Add('PASS preset '+$expectedTask[0]+'='+$expectedTask[1])
  }
  $xlTask.Run($prefixTask+'FEMWriteNumericSetting','ACCEL_STEP_RECOVERY',[double]1)
  if($xlTask.Run($prefixTask+'BestDefaultState') -ne 'True|True'){throw 'explicit experimental opt-in failed'}
  $checksTask.Add('PASS explicit recovery opt-in remains available')
  $mTask=$mappingTask|Where-Object {$_.key -eq 'ACCEL_STEP_RECOVERY'}
  $wbTask.Worksheets.Item($mTask.sheet).Unprotect()
  $wbTask.Worksheets.Item($mTask.sheet).Cells($mTask.row,6).ClearContents()
  $wbTask.Worksheets.Item('設定').Cells($mTask.legacyRow,5).ClearContents()
  if($xlTask.Run($prefixTask+'BestDefaultState') -ne 'True|False'){throw 'missing recovery key fallback enabled experiment'}
  $checksTask.Add('PASS missing key fallback keeps recovery OFF')
  if($xlTask.Run($prefixTask+'BestTestSeedRecovery') -ne 0){throw 'newly seeded recovery setting is not OFF'}
  $checksTask.Add('PASS regenerated settings seed recovery OFF')
  $checksTask|Set-Content tests/results/practical_best03_defaults_20261008.txt -Encoding UTF8
  Write-Output ('PASS '+$checksTask.Count+' best preset, missing-key, seeding and opt-in checks')
  $wbTask.Close($false)
}finally{
  if($null -ne $wbTask){try{$wbTask.Close($false)}catch{}}
  try{$xlTask.Quit()}catch{}
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)|Out-Null
}

} finally { Pop-Location }
