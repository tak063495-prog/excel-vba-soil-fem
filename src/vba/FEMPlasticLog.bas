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
Private Function PLNum(ByVal v As Double) As String
  PLNum = Replace(Format$(v, "0.000000000000E+00"), Application.International(xlDecimalSeparator), ".")
End Function
Public Function PlasticLogWrite(ByVal folder As String) As Boolean
  Dim fileNo As Integer, elementId As Long, gp As Long, comp As Long, line As String, coverage As String
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
  On Error Resume Next
  If fileNo > 0 Then Close #fileNo
  Err.Clear
End Function
