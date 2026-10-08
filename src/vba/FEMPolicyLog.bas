Option Explicit

Public Type PolicyV2Sample
  TrialId As Long
  Stage As Long
  kind As String
  Fs As Double
  SrmStage As Long
  Increment As Long
  LambdaStart As Double
  lambdaTarget As Double
  PreviousIterations As Long
  Assessed As Boolean
  q As Double
  CorrectionMs As Double
  FreshQ As Double
  FactorPriceMs As Double
  CostKnown As Boolean
  CostStatus As String
  EstimatedNetMs As Double
  Corrections As Long
  Factorizations As Long
  FactorMs As Double
  ElapsedMs As Double
  Accepted As Boolean
  Outcome As String
  RebuildReasons As String
  Cooldown As Long
  PreviousStep As Double
  StepRatio As Double
  Probe As Boolean
  BadStreakBefore As Long
  BadStreakAfter As Long
  CostModel As String
  FailureReason As String
End Type

Private mTrialId As Long
Private mSamples As Collection, mGates As Object, mGateMark As Object
Private mS(1 To 2) As Long, mGood(1 To 2) As Long, mPoor(1 To 2) As Long
Private mUnknown(1 To 2) As Long, mAccepted(1 To 2) As Long, mFailed(1 To 2) As Long
Private mCost(1 To 2) As Long, mPositive(1 To 2) As Long, mNegative(1 To 2) As Long
Private mNet(1 To 2) As Double

Private Function PLNum(ByVal value As Double) As String
  PLNum = Replace(Format$(value, "0.000000000000E+00"), Application.International(xlDecimalSeparator), ".")
End Function

Private Function PLCsv(ByVal value As String) As String
  Dim quote As String
  quote = Chr$(34)
  If InStr(value, ",") > 0 Or InStr(value, quote) > 0 Or InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then
    PLCsv = quote & Replace(value, quote, quote & quote) & quote
  Else
    PLCsv = value
  End If
End Function

Private Function PLValue(ByVal known As Boolean, ByVal value As Double) As String
  If known Then PLValue = PLNum(value) Else PLValue = "NA"
End Function

Private Function PLPath(ByVal pathId As Long) As String
  If pathId = 2 Then PLPath = "LONG" Else PLPath = "SHORT"
End Function

Public Sub PolicyLogResetRun()
  Dim i As Long
  Set mSamples = New Collection
  Set mGates = CreateObject("Scripting.Dictionary")
  Set mGateMark = CreateObject("Scripting.Dictionary")
  mTrialId = 0
  For i = 1 To 2
    mS(i) = 0: mGood(i) = 0: mPoor(i) = 0: mUnknown(i) = 0
    mAccepted(i) = 0: mFailed(i) = 0: mCost(i) = 0
    mPositive(i) = 0: mNegative(i) = 0: mNet(i) = 0#
  Next i
End Sub

Public Sub PolicyLogMark()
  Dim key As Variant
  If mGates Is Nothing Then PolicyLogResetRun
  mTrialId = mTrialId + 1
  Set mGateMark = CreateObject("Scripting.Dictionary")
  For Each key In mGates.Keys
    mGateMark.Add CStr(key), CLng(mGates(key))
  Next key
End Sub

Public Function PolicyLogTrialId() As Long
  PolicyLogTrialId = mTrialId
End Function

Public Sub PolicyLogGate(ByVal pathId As Long, ByVal reason As String)
  Dim key As String
  If mGates Is Nothing Then PolicyLogResetRun
  If pathId < 1 Or pathId > 2 Then Exit Sub
  IncrementLogV2Gate pathId, reason
  key = CStr(pathId) & "|" & reason
  If Not mGates.Exists(key) Then mGates.Add key, 0&
  mGates(key) = CLng(mGates(key)) + 1
End Sub

Public Sub PolicyLogSample(ByVal pathId As Long, ByRef sample As PolicyV2Sample)
  Dim lineText As String
  If mSamples Is Nothing Then PolicyLogResetRun
  If pathId < 1 Or pathId > 2 Then Exit Sub
  mS(pathId) = mS(pathId) + 1
  If Not sample.Assessed Then
    mUnknown(pathId) = mUnknown(pathId) + 1
  ElseIf sample.q >= 0# And sample.q <= 0.7 Then
    mGood(pathId) = mGood(pathId) + 1
  Else
    mPoor(pathId) = mPoor(pathId) + 1
  End If
  If sample.Accepted Then mAccepted(pathId) = mAccepted(pathId) + 1 Else mFailed(pathId) = mFailed(pathId) + 1
  If sample.CostKnown Then
    mCost(pathId) = mCost(pathId) + 1
    If sample.EstimatedNetMs >= 0# Then mPositive(pathId) = mPositive(pathId) + 1 Else mNegative(pathId) = mNegative(pathId) + 1
    mNet(pathId) = mNet(pathId) + sample.EstimatedNetMs
  End If
  lineText = FEM_BUILD_STAMP & "," & CStr(sample.TrialId) & "," & PLPath(pathId)
  lineText = lineText & "," & CStr(sample.Stage) & "," & PLCsv(sample.kind) & "," & PLNum(sample.Fs) & "," & CStr(sample.SrmStage)
  lineText = lineText & "," & CStr(sample.Increment) & "," & PLNum(sample.LambdaStart) & "," & PLNum(sample.lambdaTarget) & "," & CStr(sample.PreviousIterations)
  lineText = lineText & "," & CStr(sample.Assessed) & "," & PLValue(sample.Assessed, sample.q) & "," & PLValue(sample.Assessed, sample.CorrectionMs)
  lineText = lineText & "," & PLValue(sample.FreshQ > 0#, sample.FreshQ) & "," & PLValue(sample.FactorPriceMs > 0#, sample.FactorPriceMs)
  lineText = lineText & "," & CStr(sample.CostKnown) & "," & PLCsv(sample.CostStatus) & "," & PLValue(sample.CostKnown, sample.EstimatedNetMs)
  lineText = lineText & "," & CStr(sample.Corrections) & "," & CStr(sample.Factorizations) & "," & PLNum(sample.FactorMs) & "," & PLNum(sample.ElapsedMs)
  lineText = lineText & "," & CStr(sample.Accepted) & "," & PLCsv(sample.Outcome) & "," & PLCsv(sample.RebuildReasons) & "," & CStr(sample.Cooldown)
  lineText = lineText & "," & PLValue(sample.PreviousStep > 0#, sample.PreviousStep)
  lineText = lineText & "," & PLValue(sample.PreviousStep > 0#, sample.StepRatio) & "," & CStr(sample.Probe)
  lineText = lineText & "," & CStr(sample.BadStreakBefore) & "," & CStr(sample.BadStreakAfter)
  lineText = lineText & "," & PLCsv(sample.CostModel) & "," & PLCsv(sample.FailureReason)
  mSamples.Add lineText
End Sub

Public Function PolicyLogDrain() As String
  Dim i As Long, lines() As String
  If mSamples Is Nothing Then Exit Function
  If mSamples.count = 0 Then Exit Function
  ReDim lines(0 To mSamples.count - 1)
  For i = 1 To mSamples.count
    lines(i - 1) = CStr(mSamples(i))
  Next i
  PolicyLogDrain = Join(lines, vbCrLf)
  Set mSamples = New Collection
End Function

Public Function PolicyLogGateRows(ByVal statusText As String) As String
  Dim key As Variant, oldCount As Long, delta As Long, lines As String, separator As Long
  Dim pathId As Long, reason As String
  If mGates Is Nothing Then PolicyLogResetRun
  For Each key In mGates.Keys
    oldCount = 0
    If mGateMark.Exists(key) Then oldCount = CLng(mGateMark(key))
    delta = CLng(mGates(key)) - oldCount
    If delta > 0 Then
      separator = InStr(CStr(key), "|")
      pathId = CLng(left$(CStr(key), separator - 1)): reason = mid$(CStr(key), separator + 1)
      If Len(lines) > 0 Then lines = lines & vbCrLf
      lines = lines & FEM_BUILD_STAMP & "," & CStr(mTrialId) & "," & PLNum(P3CurrentStrengthFactor) & "," & PLCsv(statusText) & "," & PLPath(pathId) & "," & PLCsv(reason) & "," & CStr(delta)
    End If
    mGateMark(key) = CLng(mGates(key))
  Next key
  PolicyLogGateRows = lines
End Function

Public Function PolicyLogSampleHeader() As String
  PolicyLogSampleHeader = "ver,trial_id,path,stage,kind,fs,srm_stage,increment,lambda_start,lambda_target,previous_iterations,assessed,q,correction_ms,fresh_q,factor_price_ms,cost_known,cost_status,estimated_net_ms,corrections,factorizations,factor_ms,elapsed_ms,accepted,outcome,rebuild_reasons,cooldown,previous_step,step_ratio,probe,bad_streak_before,bad_streak_after,cost_model,failure_reason"
End Function

Public Function PolicyLogGateHeader() As String
  PolicyLogGateHeader = "ver,trial_id,fs_exact,status,path,reason,count"
End Function

Public Function PolicyLogSummary() As String
  Dim i As Long, prefix As String, text As String
  For i = 1 To 2
    prefix = "PolicyV2_" & PLPath(i) & "_"
    text = text & prefix & "samples=" & CStr(mS(i)) & vbCrLf
    text = text & prefix & "assessed_good=" & CStr(mGood(i)) & vbCrLf
    text = text & prefix & "assessed_poor=" & CStr(mPoor(i)) & vbCrLf
    text = text & prefix & "unassessed=" & CStr(mUnknown(i)) & vbCrLf
    text = text & prefix & "accepted=" & CStr(mAccepted(i)) & vbCrLf
    text = text & prefix & "failed=" & CStr(mFailed(i)) & vbCrLf
    text = text & prefix & "cost_known=" & CStr(mCost(i)) & vbCrLf
    text = text & prefix & "predicted_positive=" & CStr(mPositive(i)) & vbCrLf
    text = text & prefix & "predicted_negative=" & CStr(mNegative(i)) & vbCrLf
    text = text & prefix & "estimated_net_ms=" & PLNum(mNet(i)) & vbCrLf
  Next i
  PolicyLogSummary = text
End Function
