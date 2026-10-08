$ErrorActionPreference='Stop'
$rootTask=Split-Path -Parent $PSScriptRoot
Push-Location $rootTask
try {
New-Item -ItemType Directory -Path tests/tmp,tests/results -Force|Out-Null
$meshTask=Get-Content tests/fixtures/smoke_mesh.json -Raw | ConvertFrom-Json
$resultsTask=[Collections.Generic.List[object]]::new()
New-Item -ItemType Directory -Path work/tests -Force | Out-Null
function Set-TestSetting($wb,$key,$value) {
  $method=$(if($value -is [string]){'FEMWriteTextSetting'}else{'FEMWriteNumericSetting'})
  $xlTask.Run("'"+$wb.Name+"'!"+$method,$key,$value)
}
function Set-TestRows($wb,$sheet,$rows,$cols) {
  $ws=$wb.Worksheets.Item($sheet)
  $ws.Range($ws.Cells(2,1),$ws.Cells($ws.UsedRange.Rows.Count,$cols)).ClearContents()
  $data=[object[,]]::new($rows.Count,$cols)
  for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $cols;$c++){$data[$r,$c]=$rows[$r][$c]}}
  $ws.Range($ws.Cells(2,1),$ws.Cells($rows.Count+1,$cols)).Value2=$data
}
$harnessTask=@'
Public Function TestMetrics() As String
  On Error GoTo Failed
  TestMetrics = ResultStatus & "|" & CStr(AnalysisOK) & "|" & CStr(RelativeResidualFree) & "|" & CStr(P3SuccessfulIncrementCount) & "|" & CStr(P6FactorizationCount) & "|" & CStr(P6FactorizationReuseCount) & "|" & CStr(P3ActivePlasticPointCount) & "|" & AnalysisMessage
  Exit Function
Failed:
  TestMetrics = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function TestDisp() As Variant
  TestDisp = TDisp
End Function
Public Function TestMaterial() As String
  On Error GoTo Failed
  TestMaterial = "material=" & CStr(P2RunMaterialPointSelfTest()) & ";edge=" & CStr(P2RunEdgeReturnSelfTest())
  Exit Function
Failed:
  TestMaterial = "FAIL " & CStr(Err.Number) & " " & Err.Description
End Function
Public Function TestRun() As String
  On Error GoTo Failed
  P0_RunAnalysis
  TestRun = TestMetrics()
  Exit Function
Failed:
  TestRun = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
'@
$copyTask='tests/tmp/best03_faults.xlsm'
Copy-Item -LiteralPath 'workbook/2DSoilFEM_20261008_practical.xlsm' -Destination $copyTask -Force
$xlTask=New-Object -ComObject Excel.Application
try {
  $xlTask.EnableEvents=$false; $xlTask.DisplayAlerts=$false; $xlTask.AutomationSecurity=1
  $wbTask=$xlTask.Workbooks.Open((Join-Path $pwd $copyTask),0,$false)
  $tmTask=$wbTask.VBProject.VBComponents.Add(1); $tmTask.Name='TestHarness'; $tmTask.CodeModule.AddFromString($harnessTask)
  Set-TestRows $wbTask '節点データ' $meshTask.nodes 9
  Set-TestRows $wbTask '要素データ' $meshTask.elements 10
  $wbTask.Worksheets.Item('載荷').Range('A2:F100').ClearContents()
  Set-TestSetting $wbTask 'EXPORT_LOAD' ''
  Set-TestSetting $wbTask 'EXPORT_MODE' 'OFF'
  Set-TestSetting $wbTask 'DEBUG_MODE' 'STAGE'
  Set-TestSetting $wbTask 'ACCEL_TRACE' 0
  Set-TestSetting $wbTask 'SRM_FIXED_FS' 0.9
  $wbTask.Worksheets.Item('材料データ').Cells(2,6).Value2=[double]5
  $wbTask.Worksheets.Item('材料データ').Cells(2,7).Value2=[double]20
  $cmEngineTask=$wbTask.VBProject.VBComponents.Item('FEMEngine').CodeModule
  $engineTask=$cmEngineTask.Lines(1,$cmEngineTask.CountOfLines)
  foreach($varTask in @('baseline','first_reject_baseline','first_reject_fresh')){
    foreach($keyTask in @('ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V2B_COST','ACCEL_V3_COST_SEARCH','ACCEL_V4_ANDERSON','ACCEL_V5_GMRES_LU')){Set-TestSetting $wbTask $keyTask 0}
    Set-TestSetting $wbTask 'ACCEL_V2A_REUSE' 1
    Set-TestSetting $wbTask 'ACCEL_ADAPTIVE' 1
    $engineVariantTask=$engineTask
    if($varTask -ne 'baseline'){
      $engineVariantTask=$engineVariantTask.Replace('Option Explicit',"Option Explicit`r`nPrivate mPolicyTestInjected As Boolean")
      $retryLiteralTask='False'
      if($varTask -eq 'first_reject_fresh'){$retryLiteralTask='True'}
      $faultTask=@'
  If AccelFirstReusePending And Not mPolicyTestInjected Then
    mPolicyTestInjected = True
    P3LineSearchRetryCorrection = POLICY_RETRY_LITERAL
    SetAnalysisFailure RESULT_NONCONVERGED, "INJECTED_POLICY_TEST", vbObjectError + 3715, -1, -1, CurrentIncrement, CurrentIteration
    Exit Function
  End If
'@
      $faultTask=$faultTask.Replace('POLICY_RETRY_LITERAL',$retryLiteralTask)
      if(-not $engineVariantTask.Contains('  cutsBefore = P3LineSearchCuts')){throw 'Fault hook missing'}
      $engineVariantTask=$engineVariantTask.Replace('  cutsBefore = P3LineSearchCuts',$faultTask+"`r`n  cutsBefore = P3LineSearchCuts")
    }
    $cmEngineTask.DeleteLines(1,$cmEngineTask.CountOfLines); $cmEngineTask.AddFromString($engineVariantTask)
    $xlTask.Run("'"+$wbTask.Name+"'!FEMInvalidateSettingCache")
    $runTask=$xlTask.Run("'"+$wbTask.Name+"'!TestRun")
    $accTask=$xlTask.Run("'"+$wbTask.Name+"'!AccelSummary")
    $dispTask=@($xlTask.Run("'"+$wbTask.Name+"'!TestDisp"))
    $resultsTask.Add([PSCustomObject]@{Variant=$varTask;Metrics=$runTask;Acceleration=$accTask;Disp=$dispTask})
    $keepTask=Join-Path $pwd ('tests/results/best03_fault_logs/'+$varTask)
    New-Item -ItemType Directory -Path $keepTask -Force | Out-Null
    foreach($logTask in @('adaptive_v2a.csv','adaptive_v2a_gates.csv','solver_events.csv','perf_summary.txt','perf_summary.csv','increment_summary.csv')){
      Copy-Item -LiteralPath (Join-Path $pwd ('tests/tmp/best03_faults_out/'+$logTask)) -Destination $keepTask -Force
    }
    Write-Output ($varTask+' '+$runTask)
    if(-not $runTask.StartsWith('PASS|True')){throw $runTask}
  }
  $resultsTask | ConvertTo-Json -Depth 5 | Set-Content 'tests/results/practical_best03_fault_checks_20261008.json' -Encoding UTF8
  $wbTask.Close($false)
} finally {
  if($null -ne $wbTask){try{$wbTask.Close($false)}catch{}}
  $xlTask.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)|Out-Null
}

} finally { Pop-Location }
