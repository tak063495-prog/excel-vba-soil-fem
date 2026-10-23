Option Explicit
' Diagnostic only. No values in this module feed the solver or checkpoint restore.
Private mReady As Boolean, mComplete As Boolean, mSteps As Long
Private mTotal() As Double, mMultiplier() As Double
Private mSnapTotal() As Double, mSnapMultiplier() As Double
Private mSnapReady As Boolean, mSnapComplete As Boolean, mSnapSteps As Long

Public Sub PlasticLogResetRun()
  mReady = False: mSnapReady = False: mSteps = 0: mComplete = True
End Sub
Public Sub PlasticLogResetState()
  mReady = False: mSteps = 0: mComplete = True
  If NumberOfElement < 1 Then Exit Sub
  ReDim mTotal(0 To 3, 0 To 3, 0 To NumberOfElement - 1)
  ReDim mMultiplier(0 To 3, 0 To NumberOfElement - 1)
  mReady = True
End Sub
Public Sub PlasticLogResetElement(ByVal elementId As Long)
  Dim comp As Long, gp As Long
  If Not mReady Then Exit Sub
  If elementId < 0 Or elementId > UBound(mTotal, 3) Then Exit Sub
  For gp = 0 To 3
    For comp = 0 To 3: mTotal(comp, gp, elementId) = 0#: Next comp
    mMultiplier(gp, elementId) = 0#
  Next gp
End Sub
Public Sub PlasticLogAccepted()
  Dim elementId As Long, gp As Long, comp As Long
  If Not mReady Then PlasticLogResetState
  If Not mReady Then Exit Sub
  For elementId = 0 To NumberOfElement - 1
    If P3IsElementActive(elementId) Then
      For gp = 0 To 3
        For comp = 0 To 3
          mTotal(comp, gp, elementId) = mTotal(comp, gp, elementId) + Elem(elementId).P2PlasticStrain(comp, gp)
        Next comp
        mMultiplier(gp, elementId) = mMultiplier(gp, elementId) + Elem(elementId).P2PlasticMultiplier(gp)
      Next gp
    End If
  Next elementId
  mSteps = mSteps + 1
End Sub
Public Sub PlasticLogCaptureSuccess()
  If Not mReady Then PlasticLogResetState
  If Not mReady Then mSnapReady = False: Exit Sub
  mSnapTotal = mTotal: mSnapMultiplier = mMultiplier
  mSnapSteps = mSteps: mSnapComplete = mComplete: mSnapReady = True
End Sub

Public Function PlasticLogCaptureBranch() As Variant
  If Not mReady Then PlasticLogResetState
  PlasticLogCaptureBranch = Array(mTotal, mMultiplier, mSteps, mComplete)
End Function

Public Sub PlasticLogRestoreBranch(ByVal saved As Variant)
  mTotal = saved(0): mMultiplier = saved(1): mSteps = saved(2): mComplete = saved(3): mReady = True
End Sub
Public Sub PlasticLogRestoreSuccess()
  If Not mSnapReady Then Exit Sub
  mTotal = mSnapTotal: mMultiplier = mSnapMultiplier
  mSteps = mSnapSteps: mComplete = mSnapComplete: mReady = True
End Sub
Public Sub PlasticLogResumeUnknown()
  ' Existing checkpoints contain local increments, so prior accumulation is unavailable.
  PlasticLogResetState
  mComplete = False
End Sub
' Read-only accessors for the result display (20261022). They never change the log.
Public Function PlasticLogIsReady() As Boolean
  PlasticLogIsReady = mReady
End Function
Public Function PlasticLogIsComplete() As Boolean
  PlasticLogIsComplete = mComplete
End Function
Public Function PlasticLogTotal(ByVal comp As Long, ByVal gp As Long, ByVal elementId As Long) As Double
  If Not mReady Then Exit Function
  If elementId < 0 Or elementId > UBound(mTotal, 3) Then Exit Function
  PlasticLogTotal = mTotal(comp, gp, elementId)
End Function
Public Function PlasticLogMultiplierTotal(ByVal gp As Long, ByVal elementId As Long) As Double
  If Not mReady Then Exit Function
  If elementId < 0 Or elementId > UBound(mMultiplier, 2) Then Exit Function
  PlasticLogMultiplierTotal = mMultiplier(gp, elementId)
End Function
Private Function PLNum(ByVal v As Double) As String
  PLNum = Replace(Format$(v, "0.000000000000E+00"), Application.International(xlDecimalSeparator), ".")
End Function
Public Function PlasticLogWrite(ByVal folder As String) As Boolean
  Dim fileNo As Integer, elementId As Long, gp As Long, comp As Long, line As String, coverage As String
  Dim failedNumber As Long
  On Error GoTo Failed
  If Not mReady Then Exit Function
  If mComplete Then coverage = "FROM_ZERO_STATE" Else coverage = "PARTIAL_FROM_RESUME"
  fileNo = FreeFile
  Open folder & Application.PathSeparator & "plastic_cumulative.csv" For Output As #fileNo
  Print #fileNo, "elem,gp,ex_pl_cumulative,ey_pl_cumulative,gxy_pl_cumulative,ez_pl_cumulative,mult_cumulative,coverage,successful_increments"
  For elementId = 0 To NumberOfElement - 1
    For gp = 0 To 3
      line = CStr(elementId + 1) & "," & CStr(gp)
      For comp = 0 To 3: line = line & "," & PLNum(mTotal(comp, gp, elementId)): Next comp
      Print #fileNo, line & "," & PLNum(mMultiplier(gp, elementId)) & "," & coverage & "," & CStr(mSteps)
    Next gp
  Next elementId
  Close #fileNo
  PlasticLogWrite = True
  Exit Function
Failed:
  failedNumber = Err.number
  On Error Resume Next
  If fileNo > 0 Then Close #fileNo
  On Error GoTo 0
  If failedNumber = 18 Then
    P3UserCancel = True: P3SearchFatal = True
    Err.Raise 18
  End If
End Function

Public Function PlasticLogRead(ByVal folder As String) As Boolean
  Dim f As Integer, line As String, fields() As String, e As Long, g As Long, j As Long
  Dim steps As Long, complete As Boolean, totals() As Double, multipliers() As Double
  On Error GoTo Failed
  f = FreeFile
  Open folder & Application.PathSeparator & "plastic_cumulative.csv" For Input As #f
  Line Input #f, line
  If line <> "elem,gp,ex_pl_cumulative,ey_pl_cumulative,gxy_pl_cumulative,ez_pl_cumulative,mult_cumulative,coverage,successful_increments" Then GoTo Failed
  ReDim totals(0 To 3, 0 To 3, 0 To NumberOfElement - 1)
  ReDim multipliers(0 To 3, 0 To NumberOfElement - 1)
  For e = 0 To NumberOfElement - 1
    For g = 0 To 3
      If EOF(f) Then GoTo Failed
      Line Input #f, line: fields = Split(line, ",")
      If UBound(fields) <> 8 Then GoTo Failed
      If fields(0) <> CStr(e + 1) Or fields(1) <> CStr(g) Then GoTo Failed
      For j = 2 To 6
        If Not IsNumeric(fields(j)) Then GoTo Failed
        If Not P2IsFinite(Val(fields(j))) Then GoTo Failed
      Next j
      If fields(7) <> "FROM_ZERO_STATE" And fields(7) <> "PARTIAL_FROM_RESUME" Then GoTo Failed
      If Not IsNumeric(fields(8)) Then GoTo Failed
      If Val(fields(8)) < 0# Or Val(fields(8)) <> Fix(Val(fields(8))) Then GoTo Failed
      If e = 0 And g = 0 Then
        steps = CLng(Val(fields(8))): complete = (fields(7) = "FROM_ZERO_STATE")
      ElseIf steps <> CLng(Val(fields(8))) Or complete <> (fields(7) = "FROM_ZERO_STATE") Then
        GoTo Failed
      End If
      For j = 0 To 3: totals(j, g, e) = Val(fields(j + 2)): Next j
      multipliers(g, e) = Val(fields(6))
    Next g
  Next e
  If Not EOF(f) Then GoTo Failed
  Close #f
  mTotal = totals: mMultiplier = multipliers: mSteps = steps: mComplete = complete: mReady = True
  PlasticLogRead = True
  Exit Function
Failed:
  On Error Resume Next
  Close #f
End Function
