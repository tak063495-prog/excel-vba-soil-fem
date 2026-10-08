Option Explicit

Private mActive As Boolean, mSkip As Boolean, mAttempt As Long
Private mTrial As Long, mStage As Long, mSrm As Long, mInc As Long, mRetry As Long
Private mKind As String, mFs As Double, mLs As Double, mLt As Double, mPlan As Double
Private mBaseline As Boolean, mPredictor As Boolean, mEligible As Boolean, mV2Enabled As Boolean
Private mV2Path As String, mPrevV2 As Long, mV2Used As Boolean, mFirstQ As Double, mHaveQ As Boolean
Private mGates As Object, mFail As String, mStartT As Double, mPrevStep As Double
Private mN0 As Long, mF0 As Long, mR0 As Long, mM0 As Double, mLines As Collection
Private mAttemptCount As Long, mAcceptedCount As Long, mBaselineCount As Long, mCutbackCount As Long, mAbortCount As Long
Private mRecoveryEnabled As Boolean, mRecoveryProbe As Boolean
Private mCorrectionRatio As Double, mCorrectionKnown As Boolean, mCuts0 As Long, mRebuilds As Object
Private mMinAlpha As Double, mAlphaKnown As Boolean

Private Function ILNum(ByVal x As Double) As String
  ILNum = Replace(Format$(x, "0.000000000000E+00"), Application.International(xlDecimalSeparator), ".")
End Function
Private Function ILCsv(ByVal s As String) As String
  Dim q As String: q = Chr$(34)
  If InStr(s, ",") > 0 Or InStr(s, q) > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then ILCsv = q & Replace(s, q, q & q) & q Else ILCsv = s
End Function
Private Function ILVal(ByVal ok As Boolean, ByVal x As Double) As String
  If ok Then ILVal = ILNum(x) Else ILVal = "NA"
End Function
Private Function ILPath(ByVal p As Long) As String
  If p = 1 Then
    ILPath = "SHORT"
  ElseIf p = 2 Then
    ILPath = "LONG"
  Else
    ILPath = "NONE"
  End If
End Function
Private Function ILElapsed(ByVal t As Double) As Double
  ILElapsed = Timer - t: If ILElapsed < 0# Then ILElapsed = ILElapsed + 86400#
  ILElapsed = ILElapsed * 1000#
End Function
Private Function ILCounter(ByVal x As Long, ByVal zero As Long) As Long
  ILCounter = x - zero: If ILCounter < 0 Then ILCounter = 0
End Function

Public Sub IncrementLogResetRun()
  Set mLines = New Collection: Set mGates = CreateObject("Scripting.Dictionary")
  mAttempt = 0: mPrevStep = 0#: mActive = False: mSkip = False
  mAttemptCount = 0: mAcceptedCount = 0: mBaselineCount = 0: mCutbackCount = 0: mAbortCount = 0
End Sub
Public Sub IncrementLogResetStage()
  If mActive Then IncrementLogClosePending "RESET_STAGE"
  mPrevStep = 0#
End Sub
Public Sub IncrementLogBegin(ByVal incrementNo As Long, ByVal LambdaStart As Double, ByVal lambdaTarget As Double, ByVal plannedStep As Double, ByVal retryCount As Long, ByVal baselineRetry As Boolean)
  If mLines Is Nothing Then IncrementLogResetRun
  If UCase$(CStr(femIoMode)) = "OFF" Then
    mSkip = True
    mActive = False
    Exit Sub
  End If
  If mActive Then IncrementLogClosePending "REPLACED_ATTEMPT"
  mAttempt = mAttempt + 1: mAttemptCount = mAttemptCount + 1
  mInc = incrementNo: mLs = LambdaStart: mLt = lambdaTarget: mPlan = plannedStep: mRetry = retryCount: mBaseline = baselineRetry
  mSkip = False
  mTrial = PolicyLogTrialId: mStage = P3RunLogStageNo: mKind = P3RunLogKind: mFs = P3CurrentStrengthFactor
  If P3SrmTrialRunning Then mSrm = P3SrmLogStageNo Else mSrm = 0
  mV2Enabled = AccelV2a: mEligible = AccelIncrementEligible: mPrevV2 = AccelV2aPreviousIterations: mV2Path = "NONE": mV2Used = False
  mRecoveryEnabled = StepRecoveryEnabled
  mRecoveryProbe = False
  If mRecoveryEnabled Then mRecoveryProbe = StepRecoveryAttemptProbe()
  If mEligible Then
    If mPrevV2 > 5 Then
      mV2Path = "LONG"
    Else
      mV2Path = "SHORT"
    End If
  End If
  mPredictor = False: mHaveQ = False: mFirstQ = 0#: Set mGates = CreateObject("Scripting.Dictionary"): mFail = ""
  mCorrectionRatio = 0#: mCorrectionKnown = False: Set mRebuilds = CreateObject("Scripting.Dictionary")
  mMinAlpha = 1#: mAlphaKnown = False
  mStartT = Timer: mN0 = P6PerfNewtonCount: mF0 = P6FactorizationCount: mR0 = P6FactorizationReuseCount: mM0 = P6ProfFactorMs: mCuts0 = P3LineSearchCuts
  mActive = True
End Sub
Public Sub IncrementLogCorrectionRatio(ByVal ratio As Double, ByVal known As Boolean)
  If mActive And Not mSkip Then mCorrectionRatio = ratio: mCorrectionKnown = known
End Sub
Public Sub IncrementLogNoteRebuild(ByVal reason As String)
  If Not mActive Or mSkip Or Len(reason) = 0 Then Exit Sub
  If Not mRebuilds.Exists(reason) Then mRebuilds.Add reason, 1& Else mRebuilds(reason) = CLng(mRebuilds(reason)) + 1
End Sub
Public Sub IncrementLogLineSearchAlpha(ByVal alpha As Double)
  If Not mActive Or mSkip Or alpha <= 0# Or alpha > 1# Then Exit Sub
  If Not mAlphaKnown Or alpha < mMinAlpha Then mMinAlpha = alpha
  mAlphaKnown = True
End Sub
Public Sub IncrementLogPredictorUsed(ByVal used As Boolean): If mActive And Not mSkip Then mPredictor = used
End Sub
Public Sub IncrementLogV2Use(ByVal pathId As Long): If mActive And Not mSkip Then mV2Path = ILPath(pathId): mV2Used = True
End Sub
Public Sub IncrementLogV2Gate(ByVal pathId As Long, ByVal reason As String)
  Dim k As String
  If Not mActive Or mSkip Then Exit Sub
  k = ILPath(pathId) & "." & reason
  If Not mGates.Exists(k) Then
    mGates.Add k, 1&
  Else
    mGates(k) = CLng(mGates(k)) + 1
  End If
End Sub
Public Sub IncrementLogV2Q(ByVal pathId As Long, ByVal q As Double)
  If mActive And Not mSkip And Not mHaveQ Then mFirstQ = q: mHaveQ = True
End Sub
Public Sub IncrementLogNoteCorrectionFailure(ByVal reason As String)
  If mActive And Not mSkip And Len(mFail) = 0 And Len(reason) > 0 Then mFail = reason
End Sub
Public Sub IncrementLogFinish(ByVal Accepted As Boolean, ByVal iterations As Long, ByVal Outcome As String, ByVal nextStep As Double, ByVal stepAction As String, ByVal reason As String)
  Dim line As String, k As Variant, gates As String, rebuilds As String, elapsed As Double, delta As Double, ratio As String
  If Not mActive Then Exit Sub
  If mSkip Then mActive = False: Exit Sub
  If Len(reason) > 0 Then mFail = reason
  For Each k In mGates.Keys
    If Len(gates) > 0 Then gates = gates & ";"
    gates = gates & CStr(k) & ":" & CStr(mGates(k))
  Next k
  For Each k In mRebuilds.Keys
    If Len(rebuilds) > 0 Then rebuilds = rebuilds & ";"
    rebuilds = rebuilds & CStr(k) & ":" & CStr(mRebuilds(k))
  Next k
  If Len(gates) = 0 Then
    If Not mV2Enabled Then
      gates = "NONE.NOT_ENABLED:1"
    ElseIf Not mEligible Then
      gates = "NONE.BASIC_INELIGIBLE:1"
    Else
      gates = "NONE.NOT_REQUESTED:1"
    End If
  End If
  delta = mLt - mLs
  If mPrevStep > 0# Then ratio = ILNum(delta / mPrevStep) Else ratio = "NA"
  elapsed = ILElapsed(mStartT)
  line = FEM_BUILD_STAMP & "," & CStr(mTrial) & "," & CStr(mStage) & "," & ILCsv(mKind)
  line = line & "," & ILNum(mFs) & "," & CStr(mSrm) & "," & CStr(mInc) & "," & CStr(mAttempt)
  line = line & "," & ILNum(mLs) & "," & ILNum(mLt) & "," & ILNum(delta) & "," & ILNum(mPlan)
  line = line & "," & ILVal(mPrevStep > 0#, mPrevStep) & "," & ratio & "," & CStr(mRetry) & "," & CStr(mBaseline)
  line = line & "," & CStr(mPredictor) & "," & CStr(mV2Enabled) & "," & CStr(mEligible) & "," & ILCsv(mV2Path)
  line = line & "," & CStr(mPrevV2) & "," & CStr(mV2Used) & "," & ILCsv(gates) & "," & ILVal(mHaveQ, mFirstQ)
  line = line & "," & CStr(iterations) & "," & CStr(ILCounter(P6PerfNewtonCount, mN0)) & "," & CStr(ILCounter(P6FactorizationCount, mF0)) & "," & CStr(ILCounter(P6FactorizationReuseCount, mR0))
  line = line & "," & ILNum(P6ProfFactorMs - mM0) & "," & ILNum(elapsed) & "," & CStr(Accepted) & "," & ILCsv(Outcome)
  line = line & "," & ILVal(nextStep >= 0#, nextStep) & "," & ILCsv(stepAction) & "," & ILCsv(mFail) & "," & ILNum(RelativeResidualFree)
  line = line & "," & CStr(mRecoveryEnabled) & "," & CStr(mRecoveryProbe)
  line = line & "," & ILNum(StepRecoveryRatio()) & "," & ILCsv(StepRecoveryReason())
  line = line & "," & CStr(StepRecoveryStableCount()) & "," & CStr(StepRecoveryCooldown())
  line = line & "," & StepRecoveryDiagnosticCsv()
  line = line & "," & ILVal(mCorrectionKnown, mCorrectionRatio) & "," & CStr(ILCounter(P3LineSearchCuts, mCuts0)) & "," & ILCsv(rebuilds)
  line = line & "," & ILVal(mAlphaKnown, mMinAlpha)
  mLines.Add line
  If Accepted Then
    mAcceptedCount = mAcceptedCount + 1
    mPrevStep = delta
  Else
    Select Case UCase$(Outcome)
      Case "BASELINE_RETRY"
        mBaselineCount = mBaselineCount + 1
      Case "CUTBACK"
        mCutbackCount = mCutbackCount + 1
      Case "ABORT", "FINAL_FAILURE"
        mAbortCount = mAbortCount + 1
    End Select
  End If
  mActive = False
End Sub
Public Sub IncrementLogClosePending(ByVal Outcome As String)
  If mActive Then IncrementLogFinish False, CurrentIteration, Outcome, -1#, "NONE", P3AccelFailureReason
End Sub
Public Function IncrementLogHeader() As String
  IncrementLogHeader = "ver,trial_id,stage,kind,fs,srm_stage,increment,attempt_id,lambda_start,lambda_target,delta_lambda,planned_step,previous_step,step_ratio,retry_count,baseline_retry,predictor_used,v2a_enabled,v2a_eligible,v2a_path,previous_v2_iterations,v2a_used,v2a_gates,v2a_first_q,iterations,newton,factorizations,reuses,factor_ms,elapsed_ms,accepted,outcome,next_step,step_action,failure_reason,relative_residual,step_recovery_enabled,step_recovery_probe,step_recovery_ratio,step_recovery_reason,step_recovery_stable,step_recovery_cooldown," & StepRecoveryDiagnosticHeader() & ",correction_ratio,line_search_cuts,rebuild_reasons,line_search_min_alpha"
End Function
Public Function IncrementLogDrain() As String
  Dim a() As String, i As Long
  If mLines Is Nothing Then Exit Function
  If mLines.count = 0 Then Exit Function
  ReDim a(0 To mLines.count - 1): For i = 1 To mLines.count: a(i - 1) = CStr(mLines(i)): Next i
  IncrementLogDrain = Join(a, vbCrLf): Set mLines = New Collection
End Function
Public Function IncrementLogSummary() As String
  IncrementLogSummary = "IncrementLogAttemptCount=" & CStr(mAttemptCount) & vbCrLf & "IncrementLogAcceptedCount=" & CStr(mAcceptedCount) & vbCrLf & "IncrementLogBaselineRetryCount=" & CStr(mBaselineCount) & vbCrLf & "IncrementLogCutbackCount=" & CStr(mCutbackCount) & vbCrLf & "IncrementLogAbortCount=" & CStr(mAbortCount)
End Function

