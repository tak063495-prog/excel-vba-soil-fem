Option Explicit
' Optional scheduling only; accepted material states and numerical tests are unchanged.
Public StepRecoveryEnabled As Boolean
Private Const SR_GROW As Double = 1.1
Private Const SR_SMALL_FRACTION As Double = 0.01
Private Const SR_STABLE_REQUIRED As Long = 5
Private Const SR_OBSERVE_REQUIRED As Long = 8
Private Const SR_CUMULATIVE_CAP As Double = 1.25
Private Const SR_PROFIT_RATIO As Double = 0.95
Private Const SR_MIN_REFERENCE_MS As Double = 625#
Private Const SR_FAILURE_PAUSE As Long = 16
Private Const SR_SOFT_PAUSE As Long = 8
Private mSmallLimit As Double, mStableStep As Double, mPlateauAnchor As Double
Private mHasShrunk As Boolean, mProbePending As Boolean, mAttemptProbe As Boolean
Private mPauseFresh As Boolean, mRejected As Boolean, mFailed As Boolean
Private mStable As Long, mCooldown As Long, mCutsStart As Long
Private mRatio As Double, mReason As String, mEvents As String
Private mChecks As Long, mGrowths As Long, mProbes As Long, mProbeGood As Long
Private mProbeSlow As Long, mProbeFailed As Long, mPauses As Long
Private mObserving As Boolean, mAttemptObserved As Boolean, mBadPending As Boolean
Private mStageOff As Boolean, mEpisodeBadCounted As Boolean, mBadStreak As Long
Private mFreshLUEnabled As Boolean, mEpisode As Long, mEpisodeAnchor As Double
Private mRestoreLimit As Double, mRestoreApplied As Boolean
Private mAttemptStart As Double, mAttemptFactorStart As Long, mAttemptNewtonStart As Long
Private mAttemptRecorded As Boolean
Private mHistMs(1 To 5) As Double, mHistDelta(1 To 5) As Double
Private mHistFactor(1 To 5) As Long, mHistNewton(1 To 5) As Long
Private mPreMs As Double, mPreDelta As Double, mPreFactor As Long, mPreNewton As Long
Private mPostMs As Double, mPostDelta As Double, mPostFactor As Long, mPostNewton As Long, mPostSuccesses As Long
Private mProfitable As Long, mUneconomic As Long, mUnmeasured As Long, mIncomplete As Long
Private mBadEpisodes As Long, mStageDisables As Long, mRestores As Long, mFreshLUSkips As Long

Public Sub StepRecoveryResetRun()
  StepRecoveryEnabled = AdaptiveEnabled And (P6ReadSetting("ACCEL_STEP_RECOVERY", 0#) = 1#)
  mFreshLUEnabled = (P6ReadSetting("ACCEL_STEP_FRESH_LU", 0#) = 1#)
  mChecks = 0: mGrowths = 0: mProbes = 0: mProbeGood = 0
  mProbeSlow = 0: mProbeFailed = 0: mPauses = 0: mEpisode = 0
  mProfitable = 0: mUneconomic = 0: mUnmeasured = 0: mIncomplete = 0
  mBadEpisodes = 0: mStageDisables = 0: mRestores = 0: mFreshLUSkips = 0
  mObserving = False: mBadPending = False
  StepRecoveryResetStage 0#
End Sub
Private Sub SRClearStable()
  mStable = 0: mStableStep = 0#
End Sub
Private Sub SREvent(ByVal why As String)
  If Len(why) = 0 Then Exit Sub
  If InStr(";" & mEvents & ";", ";" & why & ";") = 0 Then
    If Len(mEvents) > 0 Then mEvents = mEvents & ";"
    mEvents = mEvents & why
  End If
End Sub
Private Sub SRReason(ByVal why As String)
  mReason = why
  SREvent why
End Sub
Public Sub StepRecoveryResetStage(ByVal initialStep As Double)
  If mObserving Then
    If mBadPending Then
      Call SRCloseBadWindow
    Else
      mIncomplete = mIncomplete + 1
    End If
  End If
  mSmallLimit = initialStep * SR_SMALL_FRACTION
  mHasShrunk = False: mPlateauAnchor = 0#: Call SRClearStable
  mCooldown = 0: mProbePending = False: mAttemptProbe = False: mPauseFresh = False
  mRejected = False: mFailed = False: mRatio = 1#: mReason = "STAGE_RESET": mEvents = ""
  mObserving = False: mAttemptObserved = False: mBadPending = False
  mStageOff = False: mEpisodeBadCounted = False: mBadStreak = 0
  mEpisodeAnchor = 0#: mRestoreLimit = 0#: mRestoreApplied = False
  mPreMs = 0#: mPreDelta = 0#: mPreFactor = 0: mPreNewton = 0
  mPostMs = 0#: mPostDelta = 0#: mPostFactor = 0: mPostNewton = 0: mPostSuccesses = 0
End Sub
Public Sub StepRecoveryBeginAttempt()
  mAttemptProbe = StepRecoveryEnabled And mProbePending
  mAttemptObserved = StepRecoveryEnabled And mObserving
  mProbePending = False
  If mAttemptProbe Then mProbes = mProbes + 1
  mRejected = False: mFailed = False: mRatio = 1#: mEvents = ""
  If Not mBadPending Then mRestoreLimit = 0#
  mRestoreApplied = False
  If StepRecoveryEnabled Then mReason = "ATTEMPT" Else mReason = "DISABLED"
  mCutsStart = P3LineSearchCuts
  mAttemptStart = Timer: mAttemptFactorStart = P6FactorizationCount: mAttemptNewtonStart = P6PerfNewtonCount
  mAttemptRecorded = False
End Sub
Private Sub SRPause(ByVal successes As Long, ByVal why As String)
  If successes > mCooldown Then mCooldown = successes: mPauses = mPauses + 1
  mPauseFresh = True: Call SRClearStable
  mProbePending = False: mRatio = 1#: SRReason why
End Sub
Private Sub SRBadEpisode(ByVal why As String)
  If Not mObserving Then Exit Sub
  If Not mEpisodeBadCounted Then
    mEpisodeBadCounted = True: mBadStreak = mBadStreak + 1: mBadEpisodes = mBadEpisodes + 1
    If mBadStreak >= 2 And Not mStageOff Then mStageOff = True: mStageDisables = mStageDisables + 1
  End If
  mBadPending = True: mRestoreLimit = mEpisodeAnchor
  SRPause SR_FAILURE_PAUSE, why
  If mStageOff Then SREvent "STAGE_OFF_AFTER_TWO_BAD_EPISODES"
End Sub
Public Sub StepRecoveryRejectedCorrection(ByVal why As String)
  If Not StepRecoveryEnabled Then Exit Sub
  mRejected = True
  If mObserving Then SRBadEpisode "CORRECTION_REJECT:" & why Else SRPause SR_FAILURE_PAUSE, "CORRECTION_REJECT:" & why
End Sub
Private Function SRRecordAttempt() As Double
  Dim attemptMs As Double, nf As Long, nn As Long
  attemptMs = P6ElapsedMs(mAttemptStart)
  SRRecordAttempt = attemptMs
  If mAttemptRecorded Then Exit Function
  mAttemptRecorded = True
  nf = P6FactorizationCount - mAttemptFactorStart: If nf < 0 Then nf = 0
  nn = P6PerfNewtonCount - mAttemptNewtonStart: If nn < 0 Then nn = 0
  If mObserving Then
    mPostMs = mPostMs + attemptMs: mPostFactor = mPostFactor + nf: mPostNewton = mPostNewton + nn
  End If
End Function
Public Sub StepRecoveryFailedAttempt(ByVal why As String)
  Dim attemptMs As Double
  If Not StepRecoveryEnabled Then Exit Sub
  If Not mFailed And mAttemptProbe Then mProbeFailed = mProbeFailed + 1
  mFailed = True: attemptMs = SRRecordAttempt()
  If mObserving Then SRBadEpisode "ATTEMPT_FAILED:" & why Else SRPause SR_FAILURE_PAUSE, "ATTEMPT_FAILED:" & why
  ' Current attempt flags survive until BeginAttempt, including for fallback/logging.
End Sub
Private Sub SRCloseBadWindow()
  mUneconomic = mUneconomic + 1: mObserving = False: mBadPending = False
End Sub
Public Sub StepRecoveryTerminateWindow()
  If Not mObserving Then Exit Sub
  If mBadPending Then
    Call SRCloseBadWindow
  Else
    mObserving = False: mIncomplete = mIncomplete + 1
  End If
  SREvent "WINDOW_TERMINATED"
End Sub
Public Sub StepRecoveryCutback(Optional ByVal cutWidth As Double = 0#)
  If Not StepRecoveryEnabled Then Exit Sub
  mHasShrunk = True: Call SRClearStable
  If cutWidth > 0# Then mPlateauAnchor = cutWidth
  If mObserving Then
    SRBadEpisode "CUTBACK_IN_WINDOW"
    Call SRCloseBadWindow
  End If
End Sub
Private Sub SRTickPause()
  If mCooldown > 0 Then
    If Not mPauseFresh Then mCooldown = mCooldown - 1
  End If
  mPauseFresh = False
End Sub
Private Sub SRHistory(ByVal actualStep As Double, ByVal nominalStep As Double, ByVal attemptMs As Double)
  Dim j As Long, nf As Long, nn As Long
  If mStableStep > 0# Then
    If Abs(nominalStep - mStableStep) > nominalStep * 0.00000001 Then SRClearStable
  End If
  If mStable = SR_STABLE_REQUIRED Then
    For j = 1 To 4
      mHistMs(j) = mHistMs(j + 1): mHistDelta(j) = mHistDelta(j + 1)
      mHistFactor(j) = mHistFactor(j + 1): mHistNewton(j) = mHistNewton(j + 1)
    Next j
  Else
    mStable = mStable + 1
  End If
  nf = P6FactorizationCount - mAttemptFactorStart: If nf < 0 Then nf = 0
  nn = P6PerfNewtonCount - mAttemptNewtonStart: If nn < 0 Then nn = 0
  mHistMs(mStable) = attemptMs: mHistDelta(mStable) = actualStep
  mHistFactor(mStable) = nf: mHistNewton(mStable) = nn
  mStableStep = nominalStep
End Sub
Public Function StepRecoveryAccepted(ByVal iterationCount As Long, ByVal actualStep As Double, ByVal nominalStep As Double, _
    ByVal lambdaEnd As Double, ByVal correctionSize As Double, ByVal retries As Long, ByVal baselineRetried As Boolean, _
    ByVal boundarySatisfied As Boolean, ByVal methodLoss As Boolean) As Double
  Dim why As String, desiredStep As Double, attemptMs As Double, j As Long
  Dim refMs As Double, refDelta As Double, refFactor As Long, refNewton As Long
  StepRecoveryAccepted = 1#: mRatio = 1#
  If Not StepRecoveryEnabled Then mReason = "DISABLED": Exit Function
  mChecks = mChecks + 1: attemptMs = SRRecordAttempt()
  If iterationCount >= 6 Then mHasShrunk = True
  If mAttemptProbe Then
    If iterationCount <= 4 And Not baselineRetried And Not mRejected And Not methodLoss And P3LineSearchCuts = mCutsStart Then
      mProbeGood = mProbeGood + 1
    Else
      mProbeSlow = mProbeSlow + 1
    End If
  End If
  If mObserving Then
    mPostSuccesses = mPostSuccesses + 1: mPostDelta = mPostDelta + actualStep
    If baselineRetried Or retries > 0 Then
      SRBadEpisode "ACCEPT_AFTER_RETRY"
    ElseIf mRejected Then
      SREvent "ACCEPT_AFTER_REJECTION"
    ElseIf methodLoss Then
      SRBadEpisode "METHOD_LOSS_IN_WINDOW"
    ElseIf P3LineSearchCuts <> mCutsStart Then
      SRBadEpisode "LINESEARCH_DAMPING_IN_WINDOW"
    ElseIf iterationCount >= 6 Then
      SRBadEpisode "NATIVE_SHRINK_IN_WINDOW"
    ElseIf mAttemptProbe And iterationCount >= 5 Then
      SRBadEpisode "PROBE_SLOW_RESTORE"
    ElseIf Not boundarySatisfied Or RelativeResidualFree > P3_ENGINEERING_RESIDUAL Or correctionSize > P3_ENGINEERING_CORRECTION Then
      SRBadEpisode "ACCEPTANCE_MARGIN_IN_WINDOW"
    End If
    If mBadPending Then
      Call SRCloseBadWindow
    ElseIf lambdaEnd >= 1# - 0.000000000001 Then
      mObserving = False: mIncomplete = mIncomplete + 1: SRReason "STAGE_COMPLETE_WINDOW"
    ElseIf mPostSuccesses < SR_OBSERVE_REQUIRED Then
      SRReason "OBSERVE_WAIT"
    ElseIf mPreMs < SR_MIN_REFERENCE_MS Or mPostMs < SR_MIN_REFERENCE_MS Then
      mObserving = False: mUnmeasured = mUnmeasured + 1: mRestoreLimit = mEpisodeAnchor
      SRPause SR_SOFT_PAUSE, "WINDOW_TIME_RESOLUTION"
    ElseIf mPostDelta > 0# And mPreDelta > 0# And _
        mPostMs / mPostDelta <= SR_PROFIT_RATIO * mPreMs / mPreDelta And _
        CDbl(mPostFactor) / mPostDelta <= CDbl(mPreFactor) / mPreDelta Then
      mObserving = False: mProfitable = mProfitable + 1: mBadStreak = 0
      SRReason "WINDOW_PROFITABLE"
    Else
      SRBadEpisode "WINDOW_UNECONOMIC_RESTORE"
      Call SRCloseBadWindow
    End If
    Call SRClearStable: Call SRTickPause
    Exit Function
  End If
  If baselineRetried Or retries > 0 Then
    SRPause SR_FAILURE_PAUSE, "ACCEPT_AFTER_RETRY"
  ElseIf mRejected Then
    ' Precise rejection reason was already recorded.
  ElseIf methodLoss Then
    SRPause SR_SOFT_PAUSE, "METHOD_LOSS"
  ElseIf P3LineSearchCuts <> mCutsStart Then
    SRPause SR_SOFT_PAUSE, "LINESEARCH_DAMPING"
  End If
  If mStageOff Then
    SRReason "STAGE_OFF": Call SRClearStable: Call SRTickPause
    Exit Function
  End If
  If mCooldown > 0 Then
    If Not mPauseFresh Then SRReason "COOLDOWN"
    Call SRClearStable: Call SRTickPause
    Exit Function
  End If
  If P6MixedUP Or P3HasJointElements Or P6ConsolActive Then
    why = "UNSUPPORTED_MODEL"
  ElseIf lambdaEnd >= 1# - 0.000000000001 Then
    why = "STAGE_COMPLETE"
  ElseIf Not mHasShrunk Then
    why = "NO_PRIOR_SHRINK"
  ElseIf nominalStep <= 0# Or mSmallLimit <= 0# Or nominalStep > mSmallLimit Then
    why = "NOT_SMALL_STEP"
  ElseIf Abs(actualStep - nominalStep) > nominalStep * 0.00000001 Then
    why = "CLIPPED_STEP"
  ElseIf iterationCount < 3 Or iterationCount > 4 Then
    why = "ITERATION_RANGE"
  ElseIf Not boundarySatisfied Or RelativeResidualFree > P3_ENGINEERING_RESIDUAL Or correctionSize > P3_ENGINEERING_CORRECTION Then
    why = "ACCEPTANCE_MARGIN"
  End If
  If Len(why) > 0 Then SRReason why: Call SRClearStable: Exit Function
  SRHistory actualStep, nominalStep, attemptMs
  SRReason "STABLE_WAIT"
  If mStable < SR_STABLE_REQUIRED Then Exit Function
  desiredStep = nominalStep * SR_GROW
  If mPlateauAnchor <= 0# Then mPlateauAnchor = nominalStep
  If desiredStep > mSmallLimit Or desiredStep > P3_MAX_STEP_SIZE Or desiredStep > 1# - lambdaEnd Or _
      desiredStep > mPlateauAnchor * SR_CUMULATIVE_CAP Then
    SRReason "RECOVERY_CAP"
    Exit Function
  End If
  For j = 1 To 5
    refMs = refMs + mHistMs(j): refDelta = refDelta + mHistDelta(j)
    refFactor = refFactor + mHistFactor(j): refNewton = refNewton + mHistNewton(j)
  Next j
  If refMs < SR_MIN_REFERENCE_MS Or refDelta <= 0# Then SRReason "REFERENCE_TIME_RESOLUTION": Exit Function
  mPreMs = refMs: mPreDelta = refDelta: mPreFactor = refFactor: mPreNewton = refNewton
  mPostMs = 0#: mPostDelta = 0#: mPostFactor = 0: mPostNewton = 0: mPostSuccesses = 0
  mEpisode = mEpisode + 1: mEpisodeAnchor = nominalStep: mEpisodeBadCounted = False
  mBadPending = False: mObserving = True: mProbePending = True
  mRatio = SR_GROW: StepRecoveryAccepted = SR_GROW: mGrowths = mGrowths + 1
  SRReason "RECOVER_1_1": Call SRClearStable
End Function
Public Function StepRecoveryBoundNextStep(ByVal proposedStep As Double) As Double
  StepRecoveryBoundNextStep = proposedStep
  If mRestoreLimit > 0# And proposedStep > mRestoreLimit Then
    StepRecoveryBoundNextStep = mRestoreLimit
    If Not mRestoreApplied Then mRestores = mRestores + 1: mRestoreApplied = True
    SREvent "WIDTH_RESTORE"
  End If
End Function
Public Sub StepRecoveryNoteNextStep(ByVal action As String, ByVal nextWidth As Double)
  If Not StepRecoveryEnabled Then Exit Sub
  If left$(action, 6) = "SHRINK" Or left$(action, 8) = "GROW_1_5" Then mPlateauAnchor = nextWidth
End Sub
Public Function StepRecoveryAttemptProbe() As Boolean
  StepRecoveryAttemptProbe = mAttemptProbe
End Function
Public Function StepRecoveryAttemptObserved() As Boolean
  StepRecoveryAttemptObserved = mAttemptObserved
End Function
Public Function StepRecoveryFreshLU() As Boolean
  StepRecoveryFreshLU = StepRecoveryEnabled And mFreshLUEnabled And mAttemptProbe
End Function
Public Sub StepRecoveryNoteFreshLUSkip()
  mFreshLUSkips = mFreshLUSkips + 1
  SREvent "PROBE_FRESH_LU"
End Sub
Public Function StepRecoveryRatio() As Double
  StepRecoveryRatio = mRatio
End Function
Public Function StepRecoveryReason() As String
  StepRecoveryReason = mReason
End Function
Public Function StepRecoveryStableCount() As Long
  StepRecoveryStableCount = mStable
End Function
Public Function StepRecoveryCooldown() As Long
  StepRecoveryCooldown = mCooldown
End Function
Private Function SRState() As String
  If Not StepRecoveryEnabled Then
    SRState = "DISABLED"
  ElseIf mStageOff Then
    SRState = "STAGE_OFF"
  ElseIf mBadPending Then
    SRState = "RESTORE_PENDING"
  ElseIf mObserving Then
    SRState = "OBSERVE"
  ElseIf mCooldown > 0 Then
    SRState = "PAUSE"
  Else
    SRState = "CANDIDATE"
  End If
End Function
Private Function SRNum(ByVal v As Double) As String
  SRNum = Replace(Format$(v, "0.000000000000E+00"), Application.International(xlDecimalSeparator), ".")
End Function
Private Function SRValue(ByVal known As Boolean, ByVal v As Double) As String
  If known Then SRValue = SRNum(v) Else SRValue = "NA"
End Function
Private Function SRCsv(ByVal v As String) As String
  Dim q As String: q = Chr$(34)
  If InStr(v, ",") > 0 Or InStr(v, q) > 0 Or InStr(v, vbCr) > 0 Or InStr(v, vbLf) > 0 Then SRCsv = q & Replace(v, q, q & q) & q Else SRCsv = v
End Function
Public Function StepRecoveryDiagnosticHeader() As String
  Dim h As String
  h = "step_recovery_state,step_recovery_episode,step_recovery_observing,step_recovery_anchor,step_recovery_cap,step_recovery_restore"
  h = h & ",step_recovery_window_successes,step_recovery_window_ms,step_recovery_window_progress,step_recovery_pre_ms_per_lambda,step_recovery_post_ms_per_lambda,step_recovery_cost_ratio"
  h = h & ",step_recovery_pre_lu_per_lambda,step_recovery_post_lu_per_lambda,step_recovery_lu_ratio,step_recovery_pre_newton_per_lambda,step_recovery_post_newton_per_lambda"
  StepRecoveryDiagnosticHeader = h & ",step_recovery_events,step_recovery_bad_streak,step_recovery_fresh_lu"
End Function
Public Function StepRecoveryDiagnosticCsv() As String
  Dim s As String, preC As Double, postC As Double, preF As Double, postF As Double, preN As Double, postN As Double
  If mPreDelta > 0# Then preC = mPreMs / mPreDelta: preF = mPreFactor / mPreDelta: preN = mPreNewton / mPreDelta
  If mPostDelta > 0# Then postC = mPostMs / mPostDelta: postF = mPostFactor / mPostDelta: postN = mPostNewton / mPostDelta
  s = SRState() & "," & CStr(mEpisode) & "," & CStr(mObserving)
  s = s & "," & SRValue(mEpisodeAnchor > 0#, mEpisodeAnchor) & "," & SRValue(mPlateauAnchor > 0#, mPlateauAnchor * SR_CUMULATIVE_CAP)
  s = s & "," & SRValue(mRestoreLimit > 0#, mRestoreLimit) & "," & CStr(mPostSuccesses) & "," & SRNum(mPostMs) & "," & SRNum(mPostDelta)
  s = s & "," & SRValue(mPreDelta > 0#, preC) & "," & SRValue(mPostDelta > 0#, postC)
  If preC > 0# And mPostDelta > 0# Then s = s & "," & SRNum(postC / preC) Else s = s & ",NA"
  s = s & "," & SRValue(mPreDelta > 0#, preF) & "," & SRValue(mPostDelta > 0#, postF)
  If preF > 0# And mPostDelta > 0# Then s = s & "," & SRNum(postF / preF) Else s = s & ",NA"
  s = s & "," & SRValue(mPreDelta > 0#, preN) & "," & SRValue(mPostDelta > 0#, postN)
  StepRecoveryDiagnosticCsv = s & "," & SRCsv(mEvents) & "," & CStr(mBadStreak) & "," & CStr(StepRecoveryFreshLU())
End Function
Public Function StepRecoverySummary() As String
  Dim s As String
  s = "StepRecoveryEnabled=" & CStr(StepRecoveryEnabled) & vbCrLf
  s = s & "StepRecoveryPolicy=OBSERVE_8_COST_0_95_LU_NONINCREASE_CAP_1_25" & vbCrLf
  s = s & "StepRecoveryFreshLUEnabled=" & CStr(mFreshLUEnabled) & vbCrLf
  s = s & "StepRecoveryReferenceMinMs=625" & vbCrLf
  s = s & "StepRecoveryChecks=" & CStr(mChecks) & vbCrLf & "StepRecoveryGrowths=" & CStr(mGrowths) & vbCrLf
  s = s & "StepRecoveryProbes=" & CStr(mProbes) & vbCrLf & "StepRecoveryProbeGood=" & CStr(mProbeGood) & vbCrLf
  s = s & "StepRecoveryProbeSlow=" & CStr(mProbeSlow) & vbCrLf & "StepRecoveryProbeFailed=" & CStr(mProbeFailed) & vbCrLf
  s = s & "StepRecoveryPauses=" & CStr(mPauses) & vbCrLf & "StepRecoveryWindowsProfitable=" & CStr(mProfitable) & vbCrLf
  s = s & "StepRecoveryWindowsUneconomic=" & CStr(mUneconomic) & vbCrLf & "StepRecoveryWindowsUnmeasured=" & CStr(mUnmeasured) & vbCrLf
  s = s & "StepRecoveryWindowsIncomplete=" & CStr(mIncomplete) & vbCrLf & "StepRecoveryBadEpisodes=" & CStr(mBadEpisodes) & vbCrLf
  s = s & "StepRecoveryStageDisables=" & CStr(mStageDisables) & vbCrLf & "StepRecoveryWidthRestores=" & CStr(mRestores) & vbCrLf
  StepRecoverySummary = s & "StepRecoveryFreshLUSkips=" & CStr(mFreshLUSkips)
End Function





