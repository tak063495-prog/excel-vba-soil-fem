Option Explicit
' Experimental online selection. Eligibility never replaces numerical safety checks.
Public AdaptiveEnabled As Boolean
Private mCooldown(1 To 5) As Long, mUseCount(1 To 5) As Long, mSkipCount(1 To 5) As Long
Private mUsed(1 To 5) As Boolean
Private mPauseFresh(1 To 5) As Boolean
Private mRecovery As Long, mIncrementStart As Double, mLambdaStart As Double, mLambdaTarget As Double
Private mPlainQ As Double, mPlainMs As Double, mPlainSamples As Long
Private mRate As Double, mPreviousStep As Double, mSamples As Long, mAATouched As Boolean
Private mSwitchCount As Long, mFailedAttempts As Long, mAcceptedIncrements As Long
Private mAASampleCount As Long, mAAProfitable As Long, mAAMs As Double
Private mGMRESSampleCount As Long, mGMRESMs As Double, mGMRESBudgetStops As Long
Private mV2Good As Long, mV2Poor As Long, mLastGain As Double
Private mIncrementCostMs As Double, mFailedIncrementCostMs As Double
Private mV2Cooldown(1 To 2) As Long, mV2PauseFresh(1 To 2) As Boolean
Private mV2Skips(1 To 2) As Long, mV2LossStreak(1 To 2) As Long
Private mV2UsedPath As Long, mV2Active As Boolean, mV2Sample As PolicyV2Sample
Private mV2Start As Double, mV2FactorStart As Long, mV2NewtonStart As Long, mV2FactorMsStart As Double
Private mV2Reasons As Object, mV2LossThisIncrement As Boolean
Private mFreshQ As Double, mFreshSamples As Long
Private mNaturalStalls As Long, mMethodStalls As Long
Private mV2FirstHandled As Boolean, mV2HadAA As Boolean, mV2FreshSamples As Long
Private Const ADAPT_COST_MIN_FACTOR_MS As Double = 31.25

Private Function AdaptiveMethodName(ByVal methodId As Long) As String
  Select Case methodId
    Case 1: AdaptiveMethodName = "V1"
    Case 2: AdaptiveMethodName = "V2a"
    Case 3: AdaptiveMethodName = "V2b"
    Case 4: AdaptiveMethodName = "V4"
    Case 5: AdaptiveMethodName = "V5"
  End Select
End Function

Public Sub AdaptiveResetRun()
  Dim i As Long
  AdaptiveEnabled = (P6ReadSetting("ACCEL_ADAPTIVE", 0#) = 1#)
  For i = 1 To 5: mUseCount(i) = 0: mSkipCount(i) = 0: Next i
  mSwitchCount = 0: mFailedAttempts = 0: mAcceptedIncrements = 0
  mAASampleCount = 0: mAAProfitable = 0: mAAMs = 0#
  mGMRESSampleCount = 0: mGMRESMs = 0#: mGMRESBudgetStops = 0
  mV2Good = 0: mV2Poor = 0
  mV2Active = False
  PolicyLogResetRun
  StepRecoveryResetRun
  IncrementLogResetRun
  For i = 1 To 2: mV2Skips(i) = 0: Next i
  mNaturalStalls = 0: mMethodStalls = 0
  mIncrementCostMs = 0#: mFailedIncrementCostMs = 0#
  AdaptiveResetStage
End Sub

Public Sub AdaptiveResetStage()
  Dim i As Long
  AdaptiveV2aEndAttempt False, "STAGE_RESET"
  For i = 1 To 5: mCooldown(i) = 0: mUsed(i) = False: mPauseFresh(i) = False: Next i
  mRecovery = 0: mPlainQ = 0#: mPlainMs = 0#: mPlainSamples = 0
  mRate = 0#: mPreviousStep = 0#: mSamples = 0: mAATouched = False
  mLambdaStart = 0#: mLambdaTarget = 0#: mLastGain = 0#
  For i = 1 To 2: mV2Cooldown(i) = 0: mV2PauseFresh(i) = False: mV2LossStreak(i) = 0: Next i
  mV2UsedPath = 0: mV2LossThisIncrement = False
  mFreshQ = 0#: mFreshSamples = 0
End Sub

Public Sub AdaptiveBeginIncrement(ByVal LambdaStart As Double, ByVal lambdaTarget As Double)
  Dim i As Long
  AdaptiveV2aEndAttempt False, "NEXT_INCREMENT"
  For i = 1 To 5: mUsed(i) = False: mPauseFresh(i) = False: Next i
  For i = 1 To 2: mV2PauseFresh(i) = False: Next i
  mV2UsedPath = 0: mV2LossThisIncrement = False
  mIncrementStart = Timer: mLambdaStart = LambdaStart: mLambdaTarget = lambdaTarget
End Sub

Public Sub AdaptiveBeginIteration()
  mAATouched = False
End Sub

Public Function AdaptiveAllowMethod(ByVal methodId As Long) As Boolean
  AdaptiveAllowMethod = True
  If Not AdaptiveEnabled Then Exit Function
  AdaptiveAllowMethod = False
  If methodId < 1 Or methodId > 5 Then Exit Function
  If methodId = 2 Then
    AdaptiveAllowMethod = AdaptiveAllowV2a()
    Exit Function
  End If
  If mRecovery = 0 And mCooldown(methodId) = 0 Then
    AdaptiveAllowMethod = True
  Else
    mSkipCount(methodId) = mSkipCount(methodId) + 1
  End If
End Function

Public Sub AdaptiveUseMethod(ByVal methodId As Long)
  If Not AdaptiveEnabled Then Exit Sub
  If methodId < 1 Or methodId > 5 Then Exit Sub
  mUsed(methodId) = True: mUseCount(methodId) = mUseCount(methodId) + 1
End Sub

Public Function AdaptiveIncrementUsed() As Boolean
  Dim i As Long
  For i = 1 To 5
    If mUsed(i) Then AdaptiveIncrementUsed = True: Exit Function
  Next i
End Function

Private Sub AdaptiveSuspend(ByVal methodId As Long, ByVal successfulIncrements As Long, ByVal reason As String)
  If Not AdaptiveEnabled Then Exit Sub
  If methodId = 2 Then
    If mV2UsedPath > 0 Then AdaptiveSuspendV2a mV2UsedPath, successfulIncrements, reason
    Exit Sub
  End If
  If successfulIncrements <= mCooldown(methodId) Then Exit Sub
  mCooldown(methodId) = successfulIncrements: mSwitchCount = mSwitchCount + 1
  mPauseFresh(methodId) = True
  P6SolverEvent "ADAPT_SWITCH", "method=" & AdaptiveMethodName(methodId) & ";pause_successes=" & CStr(successfulIncrements) & ";reason=" & reason & ";lambda=" & Format$(mLambdaStart, "0.000000000000")
End Sub

Public Function AdaptiveV2aPath() As Long
  If AccelV2aPreviousIterations > 5 Then AdaptiveV2aPath = 2 Else AdaptiveV2aPath = 1
End Function

Private Function AdaptiveV2aName(ByVal pathId As Long) As String
  If pathId = 2 Then AdaptiveV2aName = "LONG" Else AdaptiveV2aName = "SHORT"
End Function

Private Function AdaptiveAllowV2a() As Boolean
  Dim pathId As Long, why As String
  pathId = AdaptiveV2aPath()
  AdaptiveAllowV2a = True
  If Not AdaptiveEnabled Then Exit Function
  If mRecovery > 0 Then
    why = "RECOVERY"
  ElseIf mV2Cooldown(pathId) > 0 Then
    why = "PATH_PAUSE"
  Else
    Exit Function
  End If
  AdaptiveAllowV2a = False
  mV2Skips(pathId) = mV2Skips(pathId) + 1
  mSkipCount(2) = mSkipCount(2) + 1
  PolicyLogGate pathId, why
End Function

Private Sub AdaptiveSuspendV2a(ByVal pathId As Long, ByVal successes As Long, ByVal why As String)
  If Not AdaptiveEnabled Or pathId < 1 Or pathId > 2 Then Exit Sub
  If successes <= mV2Cooldown(pathId) Then Exit Sub
  mV2Cooldown(pathId) = successes: mV2PauseFresh(pathId) = True
  mSwitchCount = mSwitchCount + 1
  P6SolverEvent "ADAPT_SWITCH", "method=V2a;path=" & AdaptiveV2aName(pathId) & ";pause_successes=" & CStr(successes) & ";reason=" & why & ";lambda=" & Format$(mLambdaStart, "0.000000000000")
End Sub

Public Sub AdaptiveV2aBeginUse()
  Dim blankSample As PolicyV2Sample
  AdaptiveV2aEndAttempt False, "REPLACED_SAMPLE"
  mV2UsedPath = AdaptiveV2aPath(): mV2Active = True
  mV2Sample = blankSample
  mV2Sample.CostStatus = "UNASSESSED"
  mV2Sample.TrialId = PolicyLogTrialId()
  mV2Sample.Stage = P3RunLogStageNo: mV2Sample.kind = P3RunLogKind
  mV2Sample.Fs = P3CurrentStrengthFactor
  If P3SrmTrialRunning Then mV2Sample.SrmStage = P3SrmLogStageNo
  mV2Sample.Increment = CurrentIncrement
  mV2Sample.LambdaStart = mLambdaStart: mV2Sample.lambdaTarget = mLambdaTarget
  mV2Sample.PreviousIterations = AccelV2aPreviousIterations
  mV2Sample.PreviousStep = AccelV2aPreviousStep
  If AccelV2aPreviousStep > 0# Then mV2Sample.StepRatio = (mLambdaTarget - mLambdaStart) / AccelV2aPreviousStep
  mV2FirstHandled = False: mV2HadAA = False
  mV2Sample.FreshQ = mFreshQ: mV2FreshSamples = mFreshSamples
  If P6FactorizationCount > 0 Then mV2Sample.FactorPriceMs = P6ProfFactorMs / P6FactorizationCount
  mV2Start = Timer: mV2FactorStart = P6FactorizationCount
  mV2FactorMsStart = P6ProfFactorMs: mV2NewtonStart = P6PerfNewtonCount
  Set mV2Reasons = CreateObject("Scripting.Dictionary")
  IncrementLogV2Use mV2UsedPath
  PolicyLogGate mV2UsedPath, "REUSE"
End Sub

Public Sub AdaptiveV2aRecordRebuild(ByVal why As String)
  If Not mV2Active Then Exit Sub
  If Not mV2Reasons.Exists(why) Then mV2Reasons.Add why, 0&
  mV2Reasons(why) = CLng(mV2Reasons(why)) + 1
End Sub

Private Sub AdaptiveV2aAssess(ByVal q As Double, ByVal CorrectionMs As Double)
  Dim extraIters As Double
  If Not mV2Active Or mV2Sample.Assessed Then Exit Sub
  mV2FirstHandled = True
  mV2Sample.Assessed = True: mV2Sample.q = q: mV2Sample.CorrectionMs = CorrectionMs
  IncrementLogV2Q mV2UsedPath, q
  If q >= 0# And q <= 0.7 Then
    mV2Good = mV2Good + 1
  Else
    mV2Poor = mV2Poor + 1: mV2LossThisIncrement = True
    AdaptiveSuspendV2a mV2UsedPath, 16, "V2A_POOR_REDUCTION"
  End If
  ' Estimated counterfactual cost, NOT a measured saving. Require local fresh-q evidence.
  ' Sub-tick factor prices on tiny systems cannot support cost-based decisions.
  If mAATouched Then
    mV2Sample.CostStatus = "AA_OVERLAP"
  ElseIf mV2FreshSamples < 2 Then
    mV2Sample.CostStatus = "FRESH_Q_WARMUP"
  ElseIf mV2Sample.FreshQ <= 0# Or mV2Sample.FreshQ >= 0.7 Then
    mV2Sample.CostStatus = "FRESH_Q_UNSUITABLE"
  ElseIf mV2Sample.FactorPriceMs < ADAPT_COST_MIN_FACTOR_MS Then
    mV2Sample.CostStatus = "FACTOR_TIMER_FLOOR"
  ElseIf CorrectionMs <= 0# Then
    mV2Sample.CostStatus = "CORRECTION_TIMER_ZERO"
  Else
    mV2Sample.CostStatus = "ESTIMATED"
  End If
  If mV2Sample.CostStatus = "ESTIMATED" Then
    mV2Sample.CostKnown = True
    mV2Sample.CostModel = "ADAPT03_REPEATED_Q_ESTIMATE"
    If q >= 0# And q < 1# Then
      extraIters = 0#
      If q > 0# Then extraIters = Log(mV2Sample.FreshQ) / Log(q) - 1#
      If extraIters < 0# Then extraIters = 0#
      mV2Sample.EstimatedNetMs = mV2Sample.FactorPriceMs - extraIters * CorrectionMs
    Else
      mV2Sample.EstimatedNetMs = -mV2Sample.FactorPriceMs
    End If
    If mV2Sample.EstimatedNetMs < 0# Then
      mV2LossStreak(mV2UsedPath) = mV2LossStreak(mV2UsedPath) + 1
      mV2LossThisIncrement = True
      If mV2LossStreak(mV2UsedPath) >= 2 Then AdaptiveSuspendV2a mV2UsedPath, 8, "V2A_COST"
    Else
      mV2LossStreak(mV2UsedPath) = 0
    End If
  Else
    mV2LossStreak(mV2UsedPath) = 0
  End If
End Sub

Public Sub AdaptiveV2aEndAttempt(ByVal Accepted As Boolean, ByVal Outcome As String)
  Dim key As Variant, reasonText As String
  If Not mV2Active Then Exit Sub
  mV2Sample.Accepted = Accepted: mV2Sample.Outcome = Outcome
  If mV2HadAA Then mV2Sample.Outcome = Outcome & "_AA_OVERLAP"
  mV2Sample.Corrections = P6PerfNewtonCount - mV2NewtonStart
  mV2Sample.Factorizations = P6FactorizationCount - mV2FactorStart
  mV2Sample.FactorMs = P6ProfFactorMs - mV2FactorMsStart
  mV2Sample.ElapsedMs = P6ElapsedMs(mV2Start)
  For Each key In mV2Reasons.Keys
    If Len(reasonText) > 0 Then reasonText = reasonText & ";"
    reasonText = reasonText & CStr(key) & ":" & CStr(mV2Reasons(key))
  Next key
  mV2Sample.RebuildReasons = reasonText
  mV2Sample.Cooldown = mV2Cooldown(mV2UsedPath)
  PolicyLogSample mV2UsedPath, mV2Sample
  mV2Active = False
End Sub

Public Sub AdaptiveV2aRejectedCorrection(Optional ByVal why As String = "CORRECTION_REJECTED")
  If Not mV2Active Or mV2FirstHandled Then Exit Sub
  mV2Sample.FailureReason = why
  mV2FirstHandled = True: mV2LossThisIncrement = True
  AdaptiveSuspendV2a mV2UsedPath, 16, "V2A_CORRECTION_REJECT"
End Sub

Public Sub AdaptiveNoteCorrection(ByVal beforeResidual As Double, ByVal afterResidual As Double, ByVal rebuilt As Boolean, ByVal ElapsedMs As Double, Optional ByVal FactorMs As Double = 0#)
  Dim q As Double, CorrectionMs As Double
  If beforeResidual <= 1E-30 Then Exit Sub
  q = afterResidual / beforeResidual
  CorrectionMs = ElapsedMs - FactorMs
  If CorrectionMs < 0# Then CorrectionMs = 0#
  If mV2Active And Not mV2FirstHandled Then AdaptiveV2aAssess q, CorrectionMs
  If rebuilt And Not mAATouched And q > 0# And q < 0.7 Then
    If mFreshSamples = 0 Then mFreshQ = q Else mFreshQ = 0.75 * mFreshQ + 0.25 * q
    mFreshSamples = mFreshSamples + 1
  End If
  If Not AdaptiveEnabled Then Exit Sub
  If rebuilt Or mAATouched Or AccelBaselineRetry Or q <= 0# Or q >= 1# Or ElapsedMs <= 0# Then Exit Sub
  If mPlainSamples = 0 Then
    mPlainQ = q: mPlainMs = ElapsedMs
  Else
    mPlainQ = 0.75 * mPlainQ + 0.25 * q
    mPlainMs = 0.75 * mPlainMs + 0.25 * ElapsedMs
  End If
  mPlainSamples = mPlainSamples + 1
End Sub

Public Function AdaptiveAllowAA(ByVal relativeResidual As Double, ByVal engineeringTolerance As Double) As Boolean
  AdaptiveAllowAA = True
  If Not AdaptiveEnabled Then Exit Function
  AdaptiveAllowAA = False
  If Not AdaptiveAllowMethod(4) Then Exit Function
  ' Probe only when ordinary frozen-tangent corrections are measurably slow.
  If mPlainSamples < 2 Or mPlainMs <= 0# Then Exit Function
  If mPlainQ < 0.55 Or mPlainQ > 0.98 Then Exit Function
  If relativeResidual <= 20# * engineeringTolerance Then Exit Function
  AdaptiveAllowAA = True
End Function

Public Sub AdaptiveNoteAA(ByVal plainResidual As Double, ByVal candidateResidual As Double, ByVal ElapsedMs As Double, ByVal Accepted As Boolean, Optional ByVal reason As String = "")
  Dim savedIters As Double, savedMs As Double
  If Not AdaptiveEnabled Then Exit Sub
  If mV2Active And Not mV2FirstHandled Then mV2HadAA = True
  mAATouched = True: mAASampleCount = mAASampleCount + 1
  mAAMs = mAAMs + ElapsedMs: mLastGain = 0#
  If Accepted And plainResidual > 1E-30 And candidateResidual > 1E-30 Then
    If candidateResidual < plainResidual And mPlainQ > 0# And mPlainQ < 1# Then
      mLastGain = Log(plainResidual / candidateResidual)
      savedIters = mLastGain / (-Log(mPlainQ))
      savedMs = savedIters * mPlainMs
    End If
  ElseIf Accepted And candidateResidual <= 1E-30 Then
    savedMs = 1E+30
  End If
  ' Keep an already-paid, numerically valid improvement. Pause FUTURE probes if uneconomic.
  If Accepted And savedMs > 1.25 * ElapsedMs Then
    mAAProfitable = mAAProfitable + 1
  Else
    If Accepted Then AdaptiveSuspend 4, 8, "AA_COST" Else AdaptiveSuspend 4, 8, "AA_REJECT_" & reason
  End If
  P6SolverEvent "ADAPT_AA", "accepted=" & CStr(Accepted) & ";sample_ms=" & Format$(ElapsedMs, "0.000") & ";estimated_saved_ms=" & Format$(savedMs, "0.000") & ";log_gain=" & Format$(mLastGain, "0.000000") & ";reason=" & reason
End Sub

Public Function AdaptiveGMRESBudgetMs() As Double
  If Not AdaptiveEnabled Then Exit Function
  If Not AdaptiveAllowMethod(5) Then AdaptiveGMRESBudgetMs = -1#: Exit Function
  ' Obtain a direct-LU price before probing; reserve headroom for copy and fallback costs.
  If P6FactorizationCount < 2 Or P6ProfFactorMs <= 0# Then AdaptiveGMRESBudgetMs = -1#: Exit Function
  AdaptiveGMRESBudgetMs = 0.6 * P6ProfFactorMs / P6FactorizationCount
End Function

Public Sub AdaptiveNoteGMRES(ByVal ElapsedMs As Double, ByVal Accepted As Boolean, ByVal budgetStopped As Boolean)
  If Not AdaptiveEnabled Then Exit Sub
  mGMRESSampleCount = mGMRESSampleCount + 1: mGMRESMs = mGMRESMs + ElapsedMs
  If budgetStopped Then mGMRESBudgetStops = mGMRESBudgetStops + 1
  If Not Accepted Then AdaptiveSuspend 5, 64, "GMRES_COST_OR_FAILURE"
  P6SolverEvent "ADAPT_GMRES", "accepted=" & CStr(Accepted) & ";elapsed_ms=" & Format$(ElapsedMs, "0.000") & ";budget_stop=" & CStr(budgetStopped)
End Sub

Public Sub AdaptiveFailedAttempt()
  Dim i As Long
  If AdaptiveEnabled Then
    mFailedAttempts = mFailedAttempts + 1
    For i = 1 To 5
      If mUsed(i) Then AdaptiveSuspend i, 8, "INCREMENT_RETRY"
    Next i
    mRecovery = 3
  End If
  AdaptiveV2aEndAttempt False, "BASELINE_RETRY"
End Sub

Public Sub AdaptiveEndIncrement(ByVal Accepted As Boolean, ByVal iterations As Long, ByVal baselineRetried As Boolean)
  Dim ElapsedMs As Double, delta As Double, rate As Double, i As Long, profile As String, stalled As Boolean
  If Accepted Then AdaptiveV2aEndAttempt True, "ACCEPT" Else AdaptiveV2aEndAttempt False, "CUTBACK"
  If Not AdaptiveEnabled Then Exit Sub
  ElapsedMs = P6ElapsedMs(mIncrementStart): delta = mLambdaTarget - mLambdaStart
  mIncrementCostMs = mIncrementCostMs + ElapsedMs
  If Not Accepted Then mFailedIncrementCostMs = mFailedIncrementCostMs + ElapsedMs
  For i = 1 To 5
    If mUsed(i) Then
      If Len(profile) > 0 Then profile = profile & "+"
      profile = profile & AdaptiveMethodName(i)
    End If
  Next i
  If Len(profile) = 0 Then profile = "BASE"
  If Accepted Then
    mAcceptedIncrements = mAcceptedIncrements + 1
    If ElapsedMs > 0# And delta > 0# Then
      rate = delta / ElapsedMs
      stalled = (mSamples >= 3 And mRate > 0# And rate < 0.25 * mRate And delta < 0.5 * mPreviousStep And iterations >= 6)
      If stalled Then
        ' Step shrinkage is not evidence that an unused or profitable method is harmful.
        If mV2LossThisIncrement Then
          mMethodStalls = mMethodStalls + 1
          P6SolverEvent "ADAPT_STALL", "class=METHOD_LOSS;path=" & AdaptiveV2aName(mV2UsedPath) & ";profile=" & profile & ";action=PATH_EVIDENCE_ONLY"
        Else
          mNaturalStalls = mNaturalStalls + 1
          P6SolverEvent "ADAPT_STALL", "class=STEP_REDUCTION;profile=" & profile & ";action=KEEP_ELIGIBLE"
        End If
      End If
      If mSamples = 0 Then mRate = rate Else mRate = 0.8 * mRate + 0.2 * rate
      mPreviousStep = delta: mSamples = mSamples + 1
    End If
    For i = 1 To 5
      If mCooldown(i) > 0 And Not mPauseFresh(i) Then mCooldown(i) = mCooldown(i) - 1
    Next i
    For i = 1 To 2
      If mV2Cooldown(i) > 0 And Not mV2PauseFresh(i) Then mV2Cooldown(i) = mV2Cooldown(i) - 1
    Next i
    If mRecovery > 0 Then mRecovery = mRecovery - 1
  Else
    mRecovery = 3
    ' Global recovery stays; only actually used methods receive additional cooldowns.
    For i = 1 To 5
      If mUsed(i) Then AdaptiveSuspend i, 8, "CUTBACK"
    Next i
  End If
  If AccelTrace Or Not Accepted Or baselineRetried Or stalled Or (mAcceptedIncrements Mod 20) = 0 Then
    P6SolverEvent "ADAPT_PROGRESS", "accepted=" & CStr(Accepted) & ";profile=" & profile & ";lambda_start=" & Format$(mLambdaStart, "0.000000000000") & ";lambda_target=" & Format$(mLambdaTarget, "0.000000000000") & ";elapsed_ms=" & Format$(ElapsedMs, "0.000") & ";progress_per_ms=" & Format$(rate, "0.000000000000") & ";iterations=" & CStr(iterations) & ";baseline_retry=" & CStr(baselineRetried)
  End If
End Sub

Public Function AdaptiveSummary() As String
  Dim s As String, i As Long
  s = "AdaptiveEnabled=" & CStr(AdaptiveEnabled) & vbCrLf
  s = s & "AdaptiveSwitchCount=" & CStr(mSwitchCount) & vbCrLf
  s = s & "AdaptiveFailedAttempts=" & CStr(mFailedAttempts) & vbCrLf
  s = s & "AdaptiveAcceptedIncrements=" & CStr(mAcceptedIncrements) & vbCrLf
  s = s & "AdaptiveIncrementCostMs=" & Format$(mIncrementCostMs, "0.000") & vbCrLf
  s = s & "AdaptiveFailedIncrementCostMs=" & Format$(mFailedIncrementCostMs, "0.000") & vbCrLf
  For i = 1 To 5
    s = s & "AdaptiveMethod_" & AdaptiveMethodName(i) & "_Used=" & CStr(mUseCount(i)) & vbCrLf
    s = s & "AdaptiveMethod_" & AdaptiveMethodName(i) & "_Skipped=" & CStr(mSkipCount(i)) & vbCrLf
  Next i
  s = s & "AdaptiveV2Good=" & CStr(mV2Good) & vbCrLf & "AdaptiveV2Poor=" & CStr(mV2Poor) & vbCrLf
  s = s & "AdaptiveAASamples=" & CStr(mAASampleCount) & vbCrLf & "AdaptiveAAProfitable=" & CStr(mAAProfitable) & vbCrLf
  s = s & "AdaptiveAACostMs=" & Format$(mAAMs, "0.000") & vbCrLf
  s = s & "AdaptiveGMRESSamples=" & CStr(mGMRESSampleCount) & vbCrLf
  s = s & "AdaptiveGMRESCostMs=" & Format$(mGMRESMs, "0.000") & vbCrLf
  s = s & "AdaptiveGMRESBudgetStops=" & CStr(mGMRESBudgetStops) & vbCrLf
  s = s & "AdaptiveV2ShortSkipped=" & CStr(mV2Skips(1)) & vbCrLf
  s = s & "AdaptiveV2LongSkipped=" & CStr(mV2Skips(2)) & vbCrLf
  s = s & "AdaptiveStepReductionStalls=" & CStr(mNaturalStalls) & vbCrLf
  s = s & "AdaptiveMethodLossStalls=" & CStr(mMethodStalls) & vbCrLf
  s = s & PolicyLogSummary() & vbCrLf
  s = s & "AdaptivePolicy=ADAPT_03_RESTORED" & vbCrLf
  s = s & "AdaptiveV2PausePolicy=PATH_LOCAL_Q16_COST8" & vbCrLf
  s = s & IncrementLogSummary() & vbCrLf
  s = s & StepRecoverySummary() & vbCrLf
  AdaptiveSummary = s
End Function

' Diagnostic access for optional recovery; it does not change ADAPT_03 decisions.
Public Function AdaptiveMethodLossThisIncrement() As Boolean
  AdaptiveMethodLossThisIncrement = mV2LossThisIncrement
End Function
