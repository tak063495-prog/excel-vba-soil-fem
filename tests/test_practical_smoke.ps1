$ErrorActionPreference='Stop'
$rootTask=Split-Path -Parent $PSScriptRoot
Push-Location $rootTask
try {
$meshTask=Get-Content tests/fixtures/smoke_mesh.json -Raw | ConvertFrom-Json
$expectedTask=Get-Content tests/fixtures/smoke_expected.json -Raw | ConvertFrom-Json
$resultsTask=[Collections.Generic.List[object]]::new()
New-Item -ItemType Directory -Path tests/tmp -Force | Out-Null
function Set-TestSetting($wb, $key, $value) {
  $methodTask=$(if($value -is [string]){'FEMWriteTextSetting'}else{'FEMWriteNumericSetting'})
  $xlTask.Run("'"+$wb.Name+"'!"+$methodTask,$key,$value)
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
$flagsTask=@('ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V2B_COST','ACCEL_V3_COST_SEARCH','ACCEL_V4_ANDERSON','ACCEL_V5_GMRES_LU')
$casesTask=@(@('Case1',10.0,5.0,0.5),@('Case2',10.0,25.0,1.1),@('Case3',20.0,5.0,0.9),@('Case4',20.0,35.0,2.0))
$filesTask= ,@('policy7','workbook/2DSoilFEM_20261008_practical.xlsm')
foreach($fTask in $filesTask){
  $copyTask='tests/tmp/practical_smoke_'+$fTask[0]+'.xlsm'
  Copy-Item -LiteralPath $fTask[1] -Destination $copyTask -Force
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
    Set-TestSetting $wbTask 'SOLVER_MAX_ITERATIONS' 2000
    foreach($caseTask in $casesTask){
      $scaleTask=1.0
      if($caseTask[0] -eq 'Case1'){$scaleTask=0.25}
      $caseNodesTask=[Collections.Generic.List[object]]::new()
      foreach($nodeTask in $meshTask.nodes){$nrTask=@($nodeTask); $nrTask[1]=[double]$nrTask[1]*$scaleTask; $nrTask[2]=[double]$nrTask[2]*$scaleTask; $caseNodesTask.Add($nrTask)}
      Set-TestRows $wbTask '節点データ' $caseNodesTask 9
      $wsMatTask=$wbTask.Worksheets.Item('材料データ')
      $wsMatTask.Cells(2,6).Value2=$caseTask[2]; $wsMatTask.Cells(2,7).Value2=$caseTask[1]
      Set-TestSetting $wbTask 'SRM_FIXED_FS' $caseTask[3]
      foreach($varTask in @('off','V1_only','V2a_only','fixed_pair','adaptive_pool','adaptive_no_recovery')){
        foreach($keyTask in $flagsTask){Set-TestSetting $wbTask $keyTask 0}
        Set-TestSetting $wbTask 'ACCEL_ADAPTIVE' 0
        if($varTask -eq 'V1_only'){Set-TestSetting $wbTask $flagsTask[0] 1}
        if($varTask -eq 'V2a_only'){Set-TestSetting $wbTask $flagsTask[1] 1}
        if($varTask -eq 'fixed_pair'){foreach($idxTask in @(0,1)){Set-TestSetting $wbTask $flagsTask[$idxTask] 1}}
        if($varTask -in @('adaptive_pool','adaptive_no_recovery')){
          Set-TestSetting $wbTask 'ACCEL_ADAPTIVE' 1
          foreach($idxTask in @(0,1,4)){Set-TestSetting $wbTask $flagsTask[$idxTask] 1}
        }
        Set-TestSetting $wbTask 'ACCEL_STEP_RECOVERY' $(if($varTask -eq 'adaptive_no_recovery'){0}else{1})
        $xlTask.Run("'"+$wbTask.Name+"'!FEMInvalidateSettingCache")
        $runTask=$xlTask.Run("'"+$wbTask.Name+"'!TestRun")
        $dispTask=@($xlTask.Run("'"+$wbTask.Name+"'!TestDisp"))
        $expectedRowTask=$expectedTask | Where-Object {$_.Case -eq $caseTask[0] -and $_.Variant -eq $varTask} | Select-Object -First 1
        if($null -eq $expectedRowTask){throw "Missing smoke fixture row: $($caseTask[0])/$varTask"}
        if($runTask -ne $expectedRowTask.Metrics){throw "Smoke Metrics mismatch: $($caseTask[0])/$varTask`nactual=$runTask`nexpected=$($expectedRowTask.Metrics)"}
        $expectedDispTask=@($expectedRowTask.Disp)
        if($dispTask.Count -ne $expectedDispTask.Count){throw "Smoke Disp length mismatch: $($caseTask[0])/$varTask"}
        for($dispIndexTask=0;$dispIndexTask -lt $dispTask.Count;$dispIndexTask++){
          if([math]::Abs([double]$dispTask[$dispIndexTask]-[double]$expectedDispTask[$dispIndexTask]) -gt 1e-12){throw "Smoke Disp mismatch: $($caseTask[0])/$varTask index $dispIndexTask"}
        }
        $accTask=$xlTask.Run("'"+$wbTask.Name+"'!AccelSummary")
        $recTask=[PSCustomObject]@{Workbook=$fTask[0];Case=$caseTask[0];Variant=$varTask;Fs=$caseTask[3];GeometryScale=$scaleTask;Metrics=$runTask;Disp=$dispTask;Acceleration=$accTask}
        if($fTask[0] -eq 'policy7'){
          $outTask=[IO.Path]::ChangeExtension((Join-Path $pwd $copyTask),$null).TrimEnd('.')+'_out'
          $keepTask=Join-Path $pwd ('tests/results/practical_smoke_logs/'+$caseTask[0]+'_'+$varTask)
          New-Item -ItemType Directory -Path $keepTask -Force | Out-Null
          foreach($logTask in @('adaptive_v2a.csv','adaptive_v2a_gates.csv','increment_summary.csv','solver_events.csv','perf_summary.txt','perf_summary.csv')){
            $lpTask=Join-Path $outTask $logTask
            if(Test-Path -LiteralPath $lpTask){Copy-Item -LiteralPath $lpTask -Destination $keepTask -Force}
          }
        }
        $resultsTask.Add($recTask)
        $resultsTask | ConvertTo-Json -Depth 5 | Set-Content 'tests/results/practical_smoke_20261008.json' -Encoding UTF8
        Write-Output ($fTask[0]+' '+$caseTask[0]+' '+$varTask+' '+$runTask)
        if($fTask[0] -eq 'policy7'){Write-Output (($accTask.Split("`n") | Where-Object {$_ -match '^V2aLongIncrement'} ) -join ';')}
        if(-not $runTask.StartsWith('PASS|True')){throw $runTask}
      }
    }
    $wbTask.Close($false)
  } finally {
    if($null -ne $wbTask){try{$wbTask.Close($false)}catch{}}
    $xlTask.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)|Out-Null
  }
}

} finally { Pop-Location }
