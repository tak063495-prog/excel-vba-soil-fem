param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force $OutputRoot|Out-Null
$sourceTask=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$hashTask=(Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash
$copyTask=Join-Path $OutputRoot 'watchdog_control.xlsm'
Copy-Item -LiteralPath $sourceTask -Destination $copyTask -Force
$xlTask=New-Object -ComObject Excel.Application
$xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1;$wbTask=$null
try{
 $wbTask=$xlTask.Workbooks.Open([IO.Path]::GetFullPath($copyTask),0,$false)
 $cmTask=$wbTask.VBProject.VBComponents.Item('FEMEngine').CodeModule
 $sTask=$cmTask.Lines(1,$cmTask.CountOfLines)
 $sTask=$sTask.Replace('Option Explicit',"Option Explicit`r`nPublic IncoWatchdogMode As Long")
 $needleTask='Private Function P3EvaluateTrialState(ByRef incrementBase() As Double, ByRef internalForce() As Double) As Boolean'
 $hookTask=@'
  If IncoWatchdogMode > 0 Then
    Elem(0).Stmat(0,0) = 100# + TDisp(0)
    If TDisp(0) <> 0.1 And IncoWatchdogMode = 5 Then
      SetAnalysisFailure RESULT_CAPACITY_ERROR, "Injected material capacity error", 7, 0, 0, 0, 0
      P3EvaluateTrialState = False: Exit Function
    End If
    internalForce(0) = TDisp(0)*TDisp(0)
    If IncoWatchdogMode = 2 Or IncoWatchdogMode >= 6 Then internalForce(0) = 1#+TDisp(0)*TDisp(0)
    P3TrialStateValid = True
    P3EvaluateTrialState = True: Exit Function
  End If
'@
 if(-not $sTask.Contains($needleTask)){throw 'Evaluation hook absent'}
 $sTask=$sTask.Replace($needleTask,($needleTask+"`r`n"+($hookTask -replace "`r?`n","`r`n")))
 $needleTask='Private Function P3RebuildTangentFromSpmat() As Boolean'
 $sTask=$sTask.Replace($needleTask,($needleTask+"`r`n  If IncoWatchdogMode > 0 Then P6CSRValues(0) = 2#*TDisp(0): P3RebuildTangentFromSpmat = True: Exit Function"))
 $cmTask.DeleteLines(1,$cmTask.CountOfLines);$cmTask.AddFromString($sTask)
 $cmTask.AddFromString(@'
Public Function IncoWatchdogRun(ByVal modeId As Long) As Variant
  Dim base(0) As Double, target(0) As Double, internal(0) As Double, corr As Double, ok As Boolean, budget As Long
  On Error GoTo Failed
  SetP0SilentMode True
  IncoWatchdogMode = modeId: nDof = 1: lastDof = 0: BandWidth = 0: NumberOfElement = 1
  ReDim TDisp(0): ReDim UDisp(0): ReDim Disp(0): ReDim Force(0): ReDim NodeCond(0): ReDim Reaction(0)
  ReDim Elem(0): ReDim P3CommittedPlasticStrain(3,3,0): ReDim P3CommittedPlasticMultiplier(3,0)
  ReDim P3CommittedYieldFunction(3,0): ReDim P3CommittedYielded(3,0): ReDim P3CommittedPrincipalStress(2,3,0)
  ReDim P3CommittedSpmatValid(3,0): ReDim P3CommittedSpmat(2,15,3,0): ReDim P3CommittedDpmat(2,2,3,0)
  ReDim P3ActivePlasticPoint(3,0): ReDim P3ElementActive(0): P3ElementActiveReady = True: P3ElementActive(0) = True
  ReDim P6CSRValues(0): P6CSRValues(0) = 0.2
  Elem(0).mStmat(0,0) = 7#: Elem(0).Stmat(0,0) = 7#
  TDisp(0) = 0.1: UDisp(0) = 0.1: Disp(0) = 4.95
  target(0) = 1#: If modeId = 2 Or modeId >= 6 Then target(0) = 0#: Disp(0) = 5#: TDisp(0) = 0#: UDisp(0) = 0#
  P6SolverMemoryLimitBytes = 0#: P6CSRStorageBytes = 0#: P6MixedUP = False
  P3UserCancel = False: P3SearchFatal = False: P3RelativeBoundaryDispError = 0#: P3MaxBoundaryDispError = 0#
  P3GlobalIterationCount = 0: P6IterativeLastResidual = 0.0123: P6IterativeLastIterations = 17
  P3ClearTransientFailure
  mLineSearchFailure = ""
  P3FlowPolicyMode = 0
  budget = 0: If modeId = 6 Or modeId = 8 Then budget = 100
  If modeId = 7 Or modeId = 9 Then budget = 99
  If modeId >= 8 Then
    ok = P3TryResidualRecovery(base, target, internal, corr, budget)
  Else
    ok = P3TryNewtonWatchdog(base, target, internal, corr, budget)
  End If
  IncoWatchdogRun = Array(ok,TDisp(0),UDisp(0),Disp(0),P6CSRValues(0),P6IterativeLastResidual,P6IterativeLastIterations,ResultStatus,AnalysisErrorNumber,Elem(0).Stmat(0,0),Elem(0).mStmat(0,0),P3GlobalIterationCount,ResidualNormFree,budget,mLineSearchFailure)
  Exit Function
Failed:
  IncoWatchdogRun = Array("ERROR", Err.Number, Err.Description)
End Function
'@)
 $cmTask=$wbTask.VBProject.VBComponents.Item('FEMSolver').CodeModule
 $sTask=$cmTask.Lines(1,$cmTask.CountOfLines)
 $needleTask='Function BandSolver() As Boolean'
 $hookTask=@'
  If IncoWatchdogMode > 0 Then
    If IncoWatchdogMode = 3 Then
      SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "Injected watchdog singular", vbObjectError+3301,-1,-1,0,0
      BandSolver = False: Exit Function
    ElseIf IncoWatchdogMode = 4 Then
      SetAnalysisFailure RESULT_RUNTIME_ERROR, "Injected user interrupt",18,-1,-1,0,0
      BandSolver = False: Exit Function
    End If
    Disp(0) = Force(0)/P6CSRValues(0)
    P6IterativeLastResidual = 0.0009: P6IterativeLastIterations = 2
    BandSolver = True: Exit Function
  End If
'@
 $sTask=$sTask.Replace($needleTask,($needleTask+"`r`n"+($hookTask -replace "`r?`n","`r`n")))
 $sTask=$sTask.Replace('Public Function P6SolveResidualRecovery(ByVal damping As Double) As Boolean', "Public Function P6SolveResidualRecovery(ByVal damping As Double) As Boolean`r`n  If IncoWatchdogMode >= 8 Then P6SolveResidualRecovery = False: Exit Function")
 $sTask=$sTask.Replace('Sub SetBoundaryCondition()',"Sub SetBoundaryCondition()`r`n  If IncoWatchdogMode > 0 Then Exit Sub")
 $cmTask.DeleteLines(1,$cmTask.CountOfLines);$cmTask.AddFromString($sTask)
 $compileTask=$xlTask.VBE.CommandBars.FindControl(1,578);if($compileTask.Enabled){$compileTask.Execute()}
 $recordsTask=@()
 foreach($modeTask in @(1,2,3,4,5,6,7,8,9)){
  $macroTask="'"+$wbTask.Name+"'!IncoWatchdogRun"
  try{$vTask=@($xlTask.Run($macroTask,[double]$modeTask))}catch{Write-Host ('Call: '+$macroTask+' mode='+$modeTask);$slTask=0;$scTask=0;$elTask=0;$ecTask=0;$paneTask=$xlTask.VBE.ActiveCodePane;$paneTask.GetSelection([ref]$slTask,[ref]$scTask,[ref]$elTask,[ref]$ecTask);Write-Host ($paneTask.CodeModule.Name+':'+$slTask+' '+$paneTask.CodeModule.Lines([Math]::Max(1,$slTask-1),3));throw}
  if($modeTask -ge 8){
   if($vTask[0] -or $vTask[13] -ne 100 -or $vTask[11] -ne ($modeTask-8) -or $vTask[14] -ne 'GLOBAL_ITERATION_LIMIT' -or $vTask[1] -ne 0 -or $vTask[2] -ne 0 -or $vTask[10] -ne 7){throw 'LM shared budget exceeded or state changed'}
  }elseif($modeTask -eq 1){
   if(-not $vTask[0] -or $vTask[11] -lt 1 -or $vTask[12] -ge .99 -or [Math]::Abs($vTask[3]-($vTask[1]-.1)) -gt 1e-12){throw ('Combined correction acceptance failed '+($vTask -join '|'))}
  }else{
   $expectedTTask=$(if($modeTask -eq 2 -or $modeTask -ge 6){0}else{.1})
   if($vTask[0] -or [Math]::Abs($vTask[1]-$expectedTTask) -gt 1e-12 -or $vTask[2] -ne $vTask[1] -or $vTask[10] -ne 7){throw ('Rejected state not restored '+($vTask -join '|'))}
   if(($modeTask -eq 2 -or $modeTask -ge 6) -and ($vTask[3] -ne 5 -or [Math]::Abs($vTask[4]) -gt 1e-12 -or $vTask[5] -ne .0123 -or $vTask[6] -ne 17 -or [Math]::Abs($vTask[9]-100) -gt 1e-12)){throw ('Rejected operator/diagnostics not restored '+($vTask -join '|'))}
   if($modeTask -ge 3 -and $modeTask -le 5 -and $vTask[9] -ne 7){throw 'Fatal material checkpoint not restored'}
   if($modeTask -eq 3 -and $vTask[7] -ne 'GLOBAL_SINGULAR'){throw 'Singular failure lost'}
   if($modeTask -eq 4 -and $vTask[8] -ne 18){throw 'User interrupt lost'}
   if($modeTask -eq 5 -and $vTask[7] -ne 'CAPACITY_ERROR'){throw 'Capacity failure lost'}
   if($modeTask -ge 6 -and ($vTask[13] -ne 100 -or $vTask[14] -ne 'GLOBAL_ITERATION_LIMIT' -or $vTask[11] -ne ($modeTask-6))){throw 'Global recovery budget exceeded or reason lost'}
  }
  $recordsTask+=@{mode=$modeTask;values=$vTask}
 }
 if((Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash -ne $hashTask){throw 'Source changed'}
 @{status='PASS';source_sha256=$hashTask;description='Synthetic nonlinear force map with native watchdog control and native material checkpoint restoration';records=$recordsTask}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $OutputRoot 'watchdog_checks.json') -Encoding UTF8
 Write-Host 'PASS: combined descent, rejected rollback, singular, user interrupt, capacity, shared iteration budget'
}finally{if($wbTask){$wbTask.Close($false)};$xlTask.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)}
