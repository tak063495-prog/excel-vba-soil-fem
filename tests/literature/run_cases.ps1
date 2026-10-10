param(
  [ValidateSet('Elastic','SRM','All')][string]$Mode='All',
  [string]$Names='*',
  [switch]$IncludeDavis,
  [ValidateSet('INCONSISTENT','DAVIS')][string]$FlowPolicy='DAVIS',
  [double]$FixedFs=0,
  [switch]$DisableAcceleration,
  [string]$CaseSuffix='',
  [string]$FixturePath=(Join-Path $PSScriptRoot 'fixtures.json'),
  [double]$SrmTolerance=0.0125,
  [double]$SrmMaximum=2.2,
  [double]$Q8HourglassFactor=0.05,
  [string]$SourceWorkbook=(Join-Path $PSScriptRoot '../../workbook/2DSoilFEM_20261008_practical.xlsm'),
  [string]$OutputRoot=(Join-Path $PSScriptRoot '../tmp/literature_validation')
)
$ErrorActionPreference='Stop'
if(-not ('SrmNativeProcess' -as [type])){
 Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SrmNativeProcess {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
 public static uint Id(long hwnd) { uint id; GetWindowThreadProcessId(new IntPtr(hwnd), out id); return id; }
}
'@
}
$sourceTask=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$sourceHashTask=(Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash
$outputTask=[IO.Path]::GetFullPath($OutputRoot)
New-Item -ItemType Directory -Path $outputTask,(Join-Path $outputTask 'cases'),(Join-Path $outputTask 'results') -Force|Out-Null
$fixturesPathTask=(Resolve-Path -LiteralPath $FixturePath).Path
$fixturesHashTask=(Get-FileHash -LiteralPath $fixturesPathTask -Algorithm SHA256).Hash
$fixturesTask=Get-Content $fixturesPathTask -Raw -Encoding UTF8|ConvertFrom-Json
if($SrmTolerance -lt 0.001 -or $SrmMaximum -le 1){throw 'Invalid SRM verification tolerance or maximum'}
if($Q8HourglassFactor -lt 0 -or $Q8HourglassFactor -gt 1){throw 'Invalid Q8 stabilization factor'}
$harnessTask=@'
Public Function LiteratureRun() As String
  On Error GoTo Failed
  SetP0SilentMode True
  P0_RunAnalysis
  LiteratureRun = ResultStatus & "|" & AnalysisMessage
  Exit Function
Failed:
  LiteratureRun = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function LiteratureBuild() As String
  LiteratureBuild = FEM_BUILD_STAMP
End Function
Public Function LiteratureMetrics() As Variant
  LiteratureMetrics = Array(AnalysisOK, RelativeResidualFree, P3FosPass, P3FosFail, P3FosMid, P3FosWidth, P3FosBracket, FSS, P3SrmTrialCount, P3SuccessfulIncrementCount, P6ReactionSumX, P6ReactionSumY, P6FactorizationCount, P6FactorizationReuseCount, P3ActivePlasticPointCount, NumberOfNode, NumberOfElement, BandWidth)
End Function
Public Function LiteratureFailureState() As Variant
  LiteratureFailureState = Array(P3FosInterpretation, P3SrmTrialClassification, P3FailureKind, P3FailureLambda, P3FailureResidual, P3FailureLinearResidual, P3FailureCorrection)
End Function
Public Function LiteratureDisplacements() As Variant
  Dim result() As Double, i As Long, k As Long
  ReDim result(0 To NumberOfNode * 3 - 1)
  For i = 0 To NumberOfNode - 1
    result(i * 3) = i + 1
    k = P6GetInternalFreeNode(FEMFreeNodeMap(i))
    If k >= 0 Then
      result(i * 3 + 1) = TDisp(k * 2)
      result(i * 3 + 2) = TDisp(k * 2 + 1)
    End If
  Next i
  LiteratureDisplacements = result
End Function
Public Function LiteratureStresses() As Variant
  Dim result() As Double, e As Long, g As Long, k As Long
  ReDim result(0 To NumberOfElement * 4 * 5 - 1)
  For e = 0 To NumberOfElement - 1
    For g = 0 To 3
      result(k) = e + 1: result(k + 1) = g + 1
      result(k + 2) = Elem(e).Stmat(0, g)
      result(k + 3) = Elem(e).Stmat(1, g)
      result(k + 4) = Elem(e).Stmat(2, g)
      k = k + 5
    Next g
  Next e
  LiteratureStresses = result
End Function
Public Function LiteratureGaussState() As Variant
  Dim result() As Double, e As Long, g As Long, k As Long
  ReDim result(0 To NumberOfElement * 4 * 8 - 1)
  For e = 0 To NumberOfElement - 1
    For g = 0 To 3
      result(k) = e + 1: result(k + 1) = g + 1
      result(k + 2) = Elem(e).Stmat(16, g)
      result(k + 3) = Elem(e).P2YieldFunction(g)
      result(k + 4) = Abs(Elem(e).P2Yielded(g))
      result(k + 5) = Elem(e).P2PrincipalStress(0, g)
      result(k + 6) = Elem(e).P2PrincipalStress(1, g)
      result(k + 7) = Elem(e).P2PrincipalStress(2, g)
      k = k + 8
    Next g
  Next e
  LiteratureGaussState = result
End Function
'@
function PutRows($wb,$name,$rows,$cols){
  $wsTask=$wb.Worksheets.Item($name)
  $lastTask=[math]::Max(2000,$wsTask.UsedRange.Rows.Count+5)
  $wsTask.Range($wsTask.Cells(2,1),$wsTask.Cells($lastTask,$cols)).ClearContents()
  if($rows.Count -eq 0){return}
  $dataTask=[object[,]]::new($rows.Count,$cols)
  for($rTask=0;$rTask -lt $rows.Count;$rTask++){
    for($cTask=0;$cTask -lt $cols;$cTask++){
      $vTask=$rows[$rTask][$cTask]
      if($null -ne $vTask -and $vTask -isnot [string]){$vTask=[double]$vTask}
      $dataTask[$rTask,$cTask]=$vTask
    }
  }
  $wsTask.Range($wsTask.Cells(2,1),$wsTask.Cells($rows.Count+1,$cols)).Value2=$dataTask
}
function Setting($wb,$key,$value){
  $methodTask=$(if($value -is [string]){'FEMWriteTextSetting'}else{'FEMWriteNumericSetting'})
  $xlTask.Run("'"+$wb.Name+"'!"+$methodTask,$key,$value)
}
function SharedBookHashTask([string]$path){
 $streamTask=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 $hashAlgorithmTask=[Security.Cryptography.SHA256]::Create()
 try{return ([BitConverter]::ToString($hashAlgorithmTask.ComputeHash($streamTask))).Replace('-','')}
 finally{$streamTask.Dispose();$hashAlgorithmTask.Dispose()}
}
foreach($caseTask in $fixturesTask.fixtures){
  $elasticTask=$caseTask.reference.source_id -eq 'elastic_validation'
  if(($Mode -eq 'Elastic' -and -not $elasticTask) -or ($Mode -eq 'SRM' -and $elasticTask) -or $caseTask.name -notlike $Names){continue}
  $policiesTask=@($FlowPolicy)
  if($IncludeDavis -and $FlowPolicy -ne 'DAVIS' -and $caseTask.name -eq 'griffiths_1999_coarse'){$policiesTask+=,'DAVIS'}
  foreach($policyTask in $policiesTask){
    $idTask=$caseTask.name+$CaseSuffix+$(if($policyTask -eq 'DAVIS'){'_davis'}else{''})
    $bookTask=Join-Path $outputTask ('cases/'+$idTask+'.xlsm')
    $resultPathTask=Join-Path $outputTask ('results/'+$idTask+'.json')
    Copy-Item -LiteralPath $sourceTask -Destination $bookTask -Force
    $xlTask=New-Object -ComObject Excel.Application
    try {
      $xlTask.Visible=$false;$xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1
      $wbTask=$xlTask.Workbooks.Open($bookTask,0,$false)
      $tmTask=$wbTask.VBProject.VBComponents.Add(1);$tmTask.Name='LiteratureHarness';$tmTask.CodeModule.AddFromString($harnessTask)
      $prefixTask="'"+$wbTask.Name+"'!"
      $xlTask.Run($prefixTask+'SetP0SilentMode',$true)
      PutRows $wbTask '材料データ' $caseTask.materials 13
      PutRows $wbTask '節点データ' $caseTask.nodes 9
      PutRows $wbTask '要素データ' $caseTask.elements 10
      PutRows $wbTask 'ステージ' $caseTask.stages 6
      PutRows $wbTask '載荷' $caseTask.loads 10
      PutRows $wbTask '接合' @() 5
      Setting $wbTask 'EXPORT_LOAD' '';Setting $wbTask 'EXPORT_MODE' 'OFF';Setting $wbTask 'DEBUG_MODE' 'STAGE';Setting $wbTask 'DEBUG_FLUSH' 1
      Setting $wbTask 'SRM_FIXED_FS' $FixedFs;Setting $wbTask 'SRM_FMAX' $SrmMaximum;Setting $wbTask 'SRM_TOL' $SrmTolerance;Setting $wbTask 'SRM_MODE' 'REAPPLY';Setting $wbTask 'SRM_PSI_POLICY' 'KEEP'
      Setting $wbTask 'Q8_HOURGLASS_FACTOR' $Q8HourglassFactor
      Setting $wbTask 'FLOW_POLICY' $policyTask;Setting $wbTask 'SOLVER' 'BAND_LU';Setting $wbTask 'RCM_POLICY' $(if($elasticTask){'OFF'}else{'ON'})
      foreach($keyTask in @('ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V2B_COST','ACCEL_V3_COST_SEARCH','ACCEL_V4_ANDERSON','ACCEL_V5_GMRES_LU','ACCEL_ADAPTIVE','ACCEL_STEP_RECOVERY','ACCEL_STEP_FRESH_LU','ACCEL_TRACE')){Setting $wbTask $keyTask 0}
      if(-not $elasticTask -and -not $DisableAcceleration){foreach($keyTask in @('ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V4_ANDERSON','ACCEL_ADAPTIVE')){Setting $wbTask $keyTask 1}}
      Setting $wbTask 'OUTPUT_STAGE_MODE' 'ALL';Setting $wbTask 'OUTPUT_INACTIVE_ELEMENTS' 'SKIP'
      $wbTask.Worksheets.Item('材料データ').Range('O9').Value2='検証ケース';$wbTask.Worksheets.Item('材料データ').Range('P9').Value2=$idTask
      $wbTask.Worksheets.Item('材料データ').Range('O10').Value2='条件・出典';$wbTask.Worksheets.Item('材料データ').Range('P10').Value2='同梱の検証報告とfixtures.jsonを参照'
      $xlTask.Run($prefixTask+'FEMViewerMarkDirty');$wbTask.Save()
      $buildTask=$xlTask.Run($prefixTask+'LiteratureBuild')
      $processTask=[SrmNativeProcess]::Id([long]$xlTask.Hwnd)
      @{case=$idTask;state='RUNNING';flow_policy=$policyTask;started_at=(Get-Date).ToString('o');build=$buildTask;source_sha256=$sourceHashTask;fixture_sha256=$fixturesHashTask;input_workbook_sha256=(SharedBookHashTask $bookTask);fixed_fs=$FixedFs;srm_tolerance=$SrmTolerance;srm_maximum=$SrmMaximum;hourglass_factor=$Q8HourglassFactor;excel_process_id=$processTask}|ConvertTo-Json|Set-Content (Join-Path $outputTask ('results/'+$idTask+'_state.json')) -Encoding UTF8
      Write-Host ('START '+$idTask+' elements='+$caseTask.elements.Count+' '+$policyTask)
      $timerTask=[Diagnostics.Stopwatch]::StartNew()
      $statusTask=$xlTask.Run($prefixTask+'LiteratureRun')
      $timerTask.Stop()
      $buildTask=$xlTask.Run($prefixTask+'LiteratureBuild')
      $metricsTask=@($xlTask.Run($prefixTask+'LiteratureMetrics'))
      $dispTask=@($xlTask.Run($prefixTask+'LiteratureDisplacements'))
      $stressTask=@($xlTask.Run($prefixTask+'LiteratureStresses'))
      $failureTask=@($xlTask.Run($prefixTask+'LiteratureFailureState'))
      $gaussTask=@($xlTask.Run($prefixTask+'LiteratureGaussState'))
      if((Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash -ne $sourceHashTask){throw 'Source workbook changed during verification'}
      $resultTask=[ordered]@{fixed_fs=$FixedFs;srm_tolerance=$SrmTolerance;srm_maximum=$SrmMaximum;hourglass_factor=$Q8HourglassFactor;fixture_sha256=$fixturesHashTask;failure_names=@('fos_interpretation','trial_class','failure_kind','failed_lambda','failed_residual','failed_linear_residual','failed_correction');failure_values=$failureTask;case=$idTask;fixture=$caseTask.name;build=$buildTask;source_sha256=$sourceHashTask;status=$statusTask;flow_policy=$policyTask;elapsed_seconds=$timerTask.Elapsed.TotalSeconds;metric_names=@('analysis_ok','relative_residual','fos_pass','fos_fail','fos_mid','fos_width','fos_bracket','fss','srm_trials','accepted_increments','reaction_x','reaction_y','factorizations','reused_factors','plastic_points','nodes','elements','bandwidth');metrics=$metricsTask;displacements_flat=$dispTask;stresses_flat=$stressTask;gauss_state_columns=@('element','gauss','sigma_z','yield_function','yielded','principal_max','principal_mid','principal_min');gauss_state_flat=$gaussTask;reference=$caseTask.reference}
      $resultTask|ConvertTo-Json -Depth 14|Set-Content $resultPathTask -Encoding UTF8
      Write-Host ('DONE '+$idTask+' '+$statusTask+' Fs=['+$metricsTask[2]+','+$metricsTask[3]+'] time='+$timerTask.Elapsed.TotalSeconds)
      $wbTask.VBProject.VBComponents.Remove($tmTask)
      $xlTask.CalculateFull();$wbTask.Save();$wbTask.Close($false)
      $statePathTask=Join-Path $outputTask ('results/'+$idTask+'_state.json')
      $stateRecordTask=Get-Content -LiteralPath $statePathTask -Raw -Encoding UTF8|ConvertFrom-Json
      $stateRecordTask.state='DONE';$stateRecordTask|Add-Member status $statusTask;$stateRecordTask|Add-Member finished_at (Get-Date).ToString('o')
      $stateRecordTask|ConvertTo-Json|Set-Content -LiteralPath $statePathTask -Encoding UTF8
    } finally {
      if($null -ne $wbTask){try{$wbTask.Close($false)}catch{}}
      try{$xlTask.Quit()}catch{}
      if($null -ne $tmTask){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($tmTask)}catch{}}
      if($null -ne $wbTask){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbTask)}catch{}}
      [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)
      [GC]::Collect();[GC]::WaitForPendingFinalizers()
    }
  }
}
