param([switch]$UseRepoSource,[switch]$ExpectFixed,[string]$Label='after')
$ErrorActionPreference='Stop'
$rootTask=Split-Path -Parent $PSScriptRoot
Push-Location $rootTask
try {
New-Item -ItemType Directory -Path tests/tmp,tests/results -Force|Out-Null
$copyTask=Join-Path $pwd ('tests/tmp/issue1_'+$Label+'.xlsm')
Copy-Item -LiteralPath 'workbook/2DSoilFEM_20261008_practical.xlsm' -Destination $copyTask -Force
$resultsTask=[Collections.Generic.List[object]]::new()
function PutRows($wb,$name,$rows,$cols){
  $ws=$wb.Worksheets.Item($name);$ws.Range($ws.Cells(2,1),$ws.Cells(2000,$cols)).ClearContents()
  $data=[object[,]]::new($rows.Count,$cols)
  for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $cols;$c++){$data[$r,$c]=$rows[$r][$c]}}
  $ws.Range($ws.Cells(2,1),$ws.Cells($rows.Count+1,$cols)).Value2=$data
}
function Setting($wb,$key,$value){
  $method=$(if($value -is [string]){'FEMWriteTextSetting'}else{'FEMWriteNumericSetting'})
  $xlTask.Run("'"+$wb.Name+"'!"+$method,$key,$value)
}
function AssertNear($actual,$expected,$name){
  if([math]::Abs($actual-$expected) -gt 1e-7*(1+[math]::Abs($expected))){throw "$name actual=$actual expected=$expected"}
}
$xlTask=New-Object -ComObject Excel.Application
try {
  $xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1
  $wbTask=$xlTask.Workbooks.Open($copyTask,0,$false)
  if($UseRepoSource){
    foreach($moduleTask in @('CONTROL','FEMCore','FEMEngine','FEMSolver','FEMPractical','FEMIo','FEMUiLayout')){
      $cmTask=$wbTask.VBProject.VBComponents.Item($moduleTask).CodeModule
      $cmTask.DeleteLines(1,$cmTask.CountOfLines)
      $cmTask.AddFromString([IO.File]::ReadAllText((Join-Path $pwd ('src/vba/'+$moduleTask+'.bas')),[Text.Encoding]::UTF8))
    }
  }
  $cmTask=$wbTask.VBProject.VBComponents.Item('FEMEngine').CodeModule
  $textTask=$cmTask.Lines(1,$cmTask.CountOfLines)
  # Exercise the supported legacy no-plan branch without changing the delivered input/UI.
  $textTask=$textTask.Replace('Option Explicit',"Option Explicit`r`nPublic IssueForceCombined As Boolean`r`nPublic IssueApplyFailureMode As Long`r`nPublic IssueInjected As Boolean")
  $needleTask="  P3LoadStagePlan`r`n  If ResultStatus"
  if(-not $textTask.Contains($needleTask)){throw 'Stage-plan hook not found'}
  $textTask=$textTask.Replace($needleTask,"  P3LoadStagePlan`r`n  If IssueForceCombined Then P3StageN = 0: P3SrmEnabled = False: P3StagePlanText = `"GRAVITY+APPLY`"`r`n  If ResultStatus")
  $needleTask="      Do`r`n        AccelBeginIteration"
  if(-not $textTask.Contains($needleTask)){throw 'Increment hook not found'}
  $injectTask=@'
      Do
        If stageId = 2 And IssueApplyFailureMode > 0 And (Not IssueInjected Or IssueApplyFailureMode = 2) Then
          IssueInjected = True
          SetAnalysisFailure RESULT_NONCONVERGED, "Issue test injected APPLY failure", vbObjectError + 3202, -1, -1, CurrentIncrement, CurrentIteration
          Exit Do
        End If
        AccelBeginIteration
'@
  $textTask=$textTask.Replace($needleTask,($injectTask -replace "`r?`n","`r`n"))
  $cmTask.DeleteLines(1,$cmTask.CountOfLines);$cmTask.AddFromString($textTask)
  $cmTask.AddFromString(@'
Public Sub IssueMode(ByVal combined As Boolean, ByVal failureMode As Long)
  IssueForceCombined = combined: IssueApplyFailureMode = failureMode: IssueInjected = False
End Sub
Public Function IssueLoadState() As Variant
  Dim i As Long, target() As Double, s(0 To 6) As Double
  ReDim target(lastDof)
  P3BuildTargetForce 0#, 0#, target
  For i = 1 To lastDof Step 2
    s(0) = s(0) + P3SelfWeightForce(i)
    s(1) = s(1) + P3AppliedForce(i)
    s(2) = s(2) + P3LockedSelfWeight(i)
    s(3) = s(3) + P3LockedApplied(i)
    s(4) = s(4) + target(i)
  Next i
  s(5) = P6ReactionSumY: s(6) = P6PerfCutbackCount
  IssueLoadState = s
End Function
Public Function IssueAdditional(ByVal loadScale As Double) As String
  Dim i As Long
  On Error GoTo Failed
  For i = 0 To lastDof: P3AppliedForce(i) = loadScale * P3LockedApplied(i): Next i
  P3StageN = 0: P3PlanHasGravity = False: P3PlanHasApply = True: P3StagePlanText = "APPLY"
  IssueApplyFailureMode = 0
  If Not P3RunLoadStages() Then IssueAdditional = AnalysisMessage: Exit Function
  P3PrepareOutputResidual
  IssueAdditional = "PASS"
  Exit Function
Failed:
  IssueAdditional = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function IssueJointPsi(ByVal psi As Double) As String
  Dim m As Material_Data, msg As String
  m.kind = "JOINT": m.kn = 1000#: m.Young = 1000#: m.ks = 1000#: m.thickness = 1#: m.cohesion = 10#: m.psai = psi
  If P3ValidateJointMaterial(m, msg) Then IssueJointPsi = "ACCEPT" Else IssueJointPsi = "REJECT|" & msg
End Function
Public Function IssueJointPath() As Variant
  Dim i As Long, j As Long, slips As Variant, values(0 To 4) As Double
  NumberOfElement = 1: NumberOfMaterial = 1: NumberOfNode = 6: NumberOfFreeNode = 6: lastDof = 11
  ReDim Elem(0 To 0): ReDim Material(0 To 0): ReDim TDisp(0 To 11)
  Elem(0).IsJoint = True: Elem(0).MatNo = 0
  For j = 0 To 5: Elem(0).node(j) = j: Next j
  Elem(0).node(6) = -1: Elem(0).node(7) = -1
  Elem(0).x(1) = 1#: Elem(0).x(2) = 0.5: Elem(0).x(4) = 1#: Elem(0).x(5) = 0.5
  Material(0).kind = "JOINT": Material(0).kn = 1000#: Material(0).ks = 1000#: Material(0).Young = 1000#: Material(0).thickness = 1#: Material(0).cohesion = 10#
  slips = Array(0#, 0.02, 0.015, 0#, -0.02)
  For i = 0 To 4
    For j = 3 To 5: TDisp(2 * j) = slips(i): TDisp(2 * j + 1) = 0.001: Next j
    P3JointEval 0, False
    values(i) = Elem(0).Stmat(1, 1)
  Next i
  IssueJointPath = values
End Function
'@)
  $tmTask=$wbTask.VBProject.VBComponents.Add(1);$tmTask.Name='IssueHarness'
  $tmTask.CodeModule.AddFromString(@'
Public Function IssueRun() As String
  P0_RunAnalysis
  IssueRun = ResultStatus & "|" & AnalysisMessage
End Function
Public Function IssueDisp() As Variant
  IssueDisp = TDisp
End Function
Public Function IssueStress() As Variant
  Dim v() As Double, e As Long, g As Long, j As Long, k As Long
  ReDim v(12 * NumberOfElement - 1)
  For e = 0 To NumberOfElement - 1: For g = 0 To 3: For j = 0 To 2
    v(k) = Elem(e).Stmat(j, g): k = k + 1
  Next j: Next g: Next e
  IssueStress = v
End Function
Public Function IssueResidual() As Double
  IssueResidual = RelativeResidualFree
End Function
Public Function IssueHydraulicState() As String
  IssueHydraulicState = CStr(P6MixedUP) & "|" & CStr(P6ConsolActive) & "|" & CStr(P6PressureCount) & "|" & CStr(P6ReadSetting("CONSOL_ENABLE", 0#))
End Function
'@)
  $prefixTask="'"+$wbTask.Name+"'!"
  foreach($keyTask in @('ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V2B_COST','ACCEL_V3_COST_SEARCH','ACCEL_V4_ANDERSON','ACCEL_V5_GMRES_LU','ACCEL_STEP_RECOVERY','ACCEL_ADAPTIVE')){Setting $wbTask $keyTask 0}
  Setting $wbTask 'EXPORT_MODE' 'OFF';Setting $wbTask 'EXPORT_LOAD' '';Setting $wbTask 'DEBUG_MODE' 'OFF';Setting $wbTask 'SRM_FIXED_FS' 0;Setting $wbTask 'RCM_POLICY' 'OFF'
  Setting $wbTask 'MESH_BC_BOTTOM' 'PINNED';Setting $wbTask 'MESH_BC_LEFT' 'ROLLER';Setting $wbTask 'MESH_BC_RIGHT' 'NONE';Setting $wbTask 'MESH_BC_TOP' 'NONE';Setting $wbTask 'MESH_BC_PIN_CORNER' 'NONE'
  Setting $wbTask 'MESH_NX' 2;Setting $wbTask 'MESH_NY' 1;Setting $wbTask 'MESH_MATERIAL' 1
  foreach($pTask in @(@(1,0,0),@(2,0,1),@(3,1,1),@(4,1,0))){Setting $wbTask ('MESH_POINT'+$pTask[0]+'_X') ([double]$pTask[1]);Setting $wbTask ('MESH_POINT'+$pTask[0]+'_Y') ([double]$pTask[2])}
  $wbTask.Worksheets.Item('接合').Range('A2:E2000').ClearContents()
  PutRows $wbTask '材料データ' (,@(1.0,14000.0,0.3,1.0,20.0,0.0,1e9,0.0,0.0,'ON','STRUCT',$null,$null)) 13
  $xlTask.Run($prefixTask+'SetP0SilentMode',$true);$xlTask.Run($prefixTask+'ボタン1_Click')
  PutRows $wbTask '載荷' (,@(1.0,'TOP','Y','FORCE',-2.0,$null,$null,$null,$null,'合計-10')) 10
  $stagesTask=@(,@(1.0,'GRAVITY',$null,$null,1.0,'自重20'))
  $stagesTask+=,@(2.0,'LOAD',1.0,$null,1.0,'外力10')
  PutRows $wbTask 'ステージ' $stagesTask 6
  $statesTask=@{};$dispsTask=@{};$stressTask=@{}
  foreach($variantTask in @('explicit','combined','cutback','failed_apply')){
    $modeTask=$(if($variantTask -eq 'cutback'){1}elseif($variantTask -eq 'failed_apply'){2}else{0})
    $xlTask.Run($prefixTask+'IssueMode',($variantTask -ne 'explicit'),$modeTask)
    Write-Host $variantTask
    $statusTask=$xlTask.Run($prefixTask+'IssueRun')
    $stateTask=@($xlTask.Run($prefixTask+'IssueLoadState'))
    $statesTask[$variantTask]=$stateTask;$dispsTask[$variantTask]=@($xlTask.Run($prefixTask+'IssueDisp'));$stressTask[$variantTask]=@($xlTask.Run($prefixTask+'IssueStress'))
    $resultsTask.Add(@{case=$variantTask;status=$statusTask;loads=$stateTask;displacements=$dispsTask[$variantTask];stress=$stressTask[$variantTask];relative_residual=$xlTask.Run($prefixTask+'IssueResidual');hydraulic=$xlTask.Run($prefixTask+'IssueHydraulicState')})
    if($variantTask -ne 'failed_apply' -and -not $statusTask.StartsWith('PASS|')){throw "$variantTask $statusTask"}
    if($variantTask -eq 'failed_apply' -and -not $statusTask.StartsWith('NONCONVERGED|')){throw $statusTask}
    if($ExpectFixed -and $variantTask -ne 'failed_apply' -and $xlTask.Run($prefixTask+'IssueResidual') -gt 1e-7){throw "$variantTask force equilibrium failed"}
  }
  AssertNear $statesTask['explicit'][5] 30 'explicit reaction'
  if($ExpectFixed){
    foreach($vTask in @('combined','cutback')){
      AssertNear $statesTask[$vTask][2] -20 "$vTask gravity lock"
      AssertNear $statesTask[$vTask][4] -30 "$vTask held target"
      AssertNear $statesTask[$vTask][5] 30 "$vTask reaction"
      for($iTask=0;$iTask -lt $dispsTask['explicit'].Count;$iTask++){AssertNear $dispsTask[$vTask][$iTask] $dispsTask['explicit'][$iTask] "$vTask displacement"}
      for($iTask=0;$iTask -lt $stressTask['explicit'].Count;$iTask++){AssertNear $stressTask[$vTask][$iTask] $stressTask['explicit'][$iTask] "$vTask stress"}
    }
    if($statesTask['cutback'][6] -lt 1){throw 'No injected cutback observed'}
    AssertNear $statesTask['failed_apply'][2] -20 'failed APPLY retains accepted gravity'
    AssertNear $statesTask['failed_apply'][3] 0 'failed APPLY does not commit target'
  }
  $xlTask.Run($prefixTask+'IssueMode',$true,0);$xlTask.Run($prefixTask+'IssueRun')|Out-Null
  foreach($sTask in @(2.0,0.0)){
    $statusTask=$xlTask.Run($prefixTask+'IssueAdditional',$sTask)
    if($statusTask -ne 'PASS'){throw $statusTask}
    $stateTask=@($xlTask.Run($prefixTask+'IssueLoadState'))
    $resultsTask.Add(@{case=$(if($sTask -eq 2){'additional_load'}else{'unload'});status=$statusTask;loads=$stateTask})
    if($ExpectFixed){AssertNear $stateTask[5] $(if($sTask -eq 2){40}else{20}) 'additional/unload reaction'}
  }
  $psi0Task=$xlTask.Run($prefixTask+'IssueJointPsi',0.0);$psi10Task=$xlTask.Run($prefixTask+'IssueJointPsi',10.0)
  $pathTask=@($xlTask.Run($prefixTask+'IssueJointPath'))
  $resultsTask.Add(@{case='joint_contract';psi0=$psi0Task;psi10=$psi10Task;path=$pathTask})
  if($psi0Task -ne 'ACCEPT'){throw $psi0Task}
  if($ExpectFixed -and -not $psi10Task.StartsWith('REJECT|')){throw 'Unsupported joint psi accepted'}
  $expectedTask=@(0.0,10.0,10.0,0.0,-10.0)
  for($iTask=0;$iTask -lt 5;$iTask++){AssertNear $pathTask[$iTask] $expectedTask[$iTask] 'documented memoryless joint model'}
  if($ExpectFixed){
    foreach($solverTask in @('BAND_LU','BAND_LDLT_EXPERIMENTAL')){
      Setting $wbTask 'SOLVER' $solverTask;Setting $wbTask 'CONSOL_ENABLE' 1
      Write-Host ('unsupported_consolidation '+$solverTask)
      $statusTask=$xlTask.Run($prefixTask+'IssueRun');$stateTask=$xlTask.Run($prefixTask+'IssueHydraulicState')
      $resultsTask.Add(@{case='unsupported_consolidation';solver=$solverTask;status=$statusTask;hydraulic=$stateTask})
      if(-not $statusTask.StartsWith('INPUT_ERROR|') -or -not $statusTask.Contains('CONSOL_ENABLE')){throw $statusTask}
      if($stateTask -ne 'False|False|0|1'){throw "Hydraulic path activated: $stateTask"}
    }
  }
  if($ExpectFixed){
    Setting $wbTask 'CONSOL_ENABLE' 0;Setting $wbTask 'SOLVER' 'BAND_LU'
    $xlTask.Run($prefixTask+'IssueMode',$false,0)
    foreach($contractTask in @('initial_joint_psi','matset_joint_psi','continuum_psi')){
      $wbTask.Worksheets.Item('接合').Range('A2:E2000').ClearContents()
      $rowsTask=@(,@(1.0,14000.0,0.3,1.0,20.0,30.0,1e9,0.0,0.0,'ON','SOIL',$null,$null))
      $stageTask=@(,@(1.0,'GRAVITY',$null,$null,1.0,'自重'))
      if($contractTask -ne 'continuum_psi'){
        $rowsTask[0][10]='STRUCT'
        $rowsTask+=,@(2.0,14000.0,0.3,1.0,20.0,30.0,1e9,0.0,0.0,'ON','SOIL',$null,$null)
        $rowsTask+=,@(3.0,100000.0,0.3,1.0,0.0,0.0,10000.0,0.0,0.0,'ON','JOINT',100000.0,1.0)
        if($contractTask -eq 'initial_joint_psi'){$rowsTask[2][7]=[double]10}
        else{$stageTask=@(,@(1.0,'MATSET',3.0,'psi=10',1.0,'未対応psi'));$stageTask+=,@(2.0,'GRAVITY',$null,$null,1.0,'自重')}
      } else {$rowsTask[0][7]=[double]5}
      PutRows $wbTask '材料データ' $rowsTask 13
      $xlTask.Run($prefixTask+'ボタン1_Click')
      if($contractTask -ne 'continuum_psi'){
        $wbTask.Worksheets.Item('要素データ').Cells(3,10).Value2=[double]2
        PutRows $wbTask '接合' (,@(1.0,2.0,3.0,1.0,'共有辺')) 5
      }
      PutRows $wbTask 'ステージ' $stageTask 6
      $statusTask=$xlTask.Run($prefixTask+'IssueRun')
      Write-Host ($contractTask+' '+$statusTask)
      $resultsTask.Add(@{case=$contractTask;status=$statusTask})
      if($contractTask -eq 'continuum_psi'){
        if(-not $statusTask.StartsWith('PASS|')){throw $statusTask}
      } elseif(-not $statusTask.StartsWith('INPUT_ERROR|') -or -not $statusTask.Contains('psi')){throw $statusTask}
    }
  }
  $resultsTask|ConvertTo-Json -Depth 8|Set-Content ('tests/results/issue1_'+$Label+'.json') -Encoding UTF8
  Write-Output ("Recorded $($resultsTask.Count) cases; ExpectFixed=$ExpectFixed")
  $wbTask.Close($false)
} finally {
  if($null -ne $wbTask){try{$wbTask.Close($false)}catch{}}
  try{$xlTask.Quit()}catch{}
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)|Out-Null
}
} finally {Pop-Location}
