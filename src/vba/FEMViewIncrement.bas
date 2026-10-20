Option Explicit

' Observer only: no element UDT/DLL ABI or solver decision is changed.
' total = B * (trial displacement - increment-start committed displacement).
' P2PlasticStrain is the current material update's plastic increment, not cumulative.
Private incData() As Double, snapData() As Double
Private incActive() As Boolean, snapActive() As Boolean
Private incReady As Boolean, snapReady As Boolean
Private incFs As Double, snapFs As Double, incLambda As Double, snapLambda As Double
Private incNo As Long, snapNo As Long
Public FEMIncrementLastError As String

Public Sub FEMIncrementClear()
  incReady = False: Erase incData: Erase incActive
End Sub

Public Sub FEMIncrementReset()
  FEMIncrementClear
  snapReady = False: Erase snapData: Erase snapActive
  FEMIncrementLastError = ""
  Dim ws As Worksheet
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("結果増分"): On Error GoTo 0
  If Not ws Is Nothing Then ws.Cells.ClearContents
End Sub

Public Sub FEMIncrementCapture(ByVal loadFactor As Double)
  Dim e As Long, gp As Long, c As Long, j As Long, dof As Long, s As Double
  On Error GoTo Failed
  incReady = False
  If NumberOfElement < 1 Or lastDof < 0 Then Exit Sub
  ReDim incData(0 To 6, 0 To 3, 0 To NumberOfElement - 1)
  ReDim incActive(0 To NumberOfElement - 1)
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Not Elem(e).IsJoint Then
      incActive(e) = True
      For gp = 0 To 3
        For c = 0 To 2
          s = 0#
          For j = 0 To 15
            dof = Elem(e).ElNode(j)
            s = s + Elem(e).Bmat(c, j, gp) * (TDisp(dof) - P3CommittedDisp(dof))
          Next j
          incData(c, gp, e) = s
        Next c
        For c = 0 To 3
          incData(c + 3, gp, e) = Elem(e).P2PlasticStrain(c, gp)
        Next c
      Next gp
    End If
  Next e
  incFs = P3CurrentStrengthFactor: incLambda = loadFactor: incNo = CurrentIncrement
  incReady = True: FEMIncrementLastError = ""
  Exit Sub
Failed:
  FEMIncrementLastError = "増分保存の準備に失敗: " & CStr(Err.Number) & " " & Err.Description
  incReady = False
End Sub

Public Sub FEMIncrementSnapshot()
  snapReady = incReady
  If incReady Then
    snapData = incData: snapActive = incActive
    snapFs = incFs: snapLambda = incLambda: snapNo = incNo
  End If
End Sub

Public Sub FEMIncrementRestore()
  incReady = snapReady
  If snapReady Then
    incData = snapData: incActive = snapActive
    incFs = snapFs: incLambda = snapLambda: incNo = snapNo
  End If
End Sub

Public Sub FEMIncrementSave()
  Dim ws As Worksheet, rows() As Variant, headers As Variant, e As Long, gp As Long, c As Long, n As Long, r As Long
  Dim x As Double, y As Double, hi As Double, lo As Double, angle As Double, stageNo As Long
  On Error GoTo Failed
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("結果増分"): On Error GoTo Failed
  If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("結果要素")): ws.Name = "結果増分"
  stageNo = P3ResultStageNumber()
  If UCase$(P6ReadTextSetting("OUTPUT_STAGE_MODE", "FINAL")) <> "ALL" Then ws.Cells.ClearContents
  headers = Array("ステージ番号", "要素番号", "積分点", "要素中心X", "要素中心Y", "全ひずみ増分X", "全ひずみ増分Y", "工学せん断ひずみ増分XY", "塑性ひずみ増分X", "塑性ひずみ増分Y", "塑性工学せん断ひずみ増分XY", "塑性ひずみ増分Z", "最大工学せん断ひずみ増分(3D)", "最大塑性工学せん断ひずみ増分(3D)", "面内塑性主ひずみ増分max", "面内塑性主ひずみ増分min", "面内max主軸角rad", "Fs", "荷重増分番号", "荷重係数")
  For c = 0 To UBound(headers): ws.Cells(1, c + 1).Value2 = headers(c): Next c
  ws.Range("V1").Value2 = "LAST_CONVERGED_INCREMENT_V1"
  If Not incReady Then
    ws.Range("V2").Value2 = "増分結果なし。新しい版で最初から再解析してください。" & FEMIncrementLastError
    Exit Sub
  End If
  ReDim rows(1 To NumberOfElement * 4, 1 To 20)
  For e = 0 To NumberOfElement - 1
    If incActive(e) Then
      x = 0#: y = 0#
      For c = 0 To 3: x = x + Elem(e).x(c) * 0.25: y = y + Elem(e).y(c) * 0.25: Next c
      For gp = 0 To 3
        n = n + 1
        rows(n, 1) = stageNo: rows(n, 2) = e + 1: rows(n, 3) = gp + 1
        rows(n, 4) = x: rows(n, 5) = y
        For c = 0 To 6: rows(n, c + 6) = incData(c, gp, e): Next c
        rows(n, 13) = FEMStrainMaxShear(incData(0, gp, e), incData(1, gp, e), incData(2, gp, e))
        rows(n, 14) = FEMStrainMaxShear(incData(3, gp, e), incData(4, gp, e), incData(5, gp, e), incData(6, gp, e))
        FEMStrainPrincipal2D incData(3, gp, e), incData(4, gp, e), incData(5, gp, e), hi, lo, angle
        rows(n, 15) = hi: rows(n, 16) = lo: rows(n, 17) = angle
        rows(n, 18) = incFs: rows(n, 19) = incNo: rows(n, 20) = incLambda
      Next gp
    End If
  Next e
  r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1: If r < 2 Then r = 2
  If n > 0 Then ws.Range(ws.Cells(r, 1), ws.Cells(r + n - 1, 20)).Value2 = rows
  ws.Range("V2").Value2 = "最終収束荷重増分。SRMでは最後の成功Fsに対応。引張正、XYは工学せん断ひずみ。"
  ws.Rows(1).Font.Bold = True: ws.Columns("A:C").ColumnWidth = 14
  ws.Range("F:T").NumberFormat = "0.00000E+00"
  Exit Sub
Failed:
  FEMIncrementLastError = "増分結果の出力に失敗: " & CStr(Err.Number) & " " & Err.Description
  If Not ws Is Nothing Then ws.Range("V1").Value2 = "INVALID": ws.Range("V2").Value2 = FEMIncrementLastError
End Sub

' Returns e -> Array(max shear, max plastic shear, mean plastic ex,ey,gxy, Fs,increment,lambda).
' All four Gauss points must exist exactly once; old/missing output is never fabricated.
Public Function FEMIncrementRead(ByVal stageNo As Long) As Object
  Dim ws As Worksheet, data As Variant, result As Object, counts As Object, seen As Object
  Dim r As Long, e As Long, gp As Long, k As String, v As Variant, j As Long
  Dim frame As Variant, frameReady As Boolean
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("結果増分"): On Error GoTo 0
  If ws Is Nothing Then GoTo Missing
  If CStr(ws.Range("V1").Value2) <> "LAST_CONVERGED_INCREMENT_V1" Then GoTo Missing
  r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
  If r < 2 Then GoTo Missing
  data = ws.Range("A1:T" & CStr(r)).Value2
  Set result = CreateObject("Scripting.Dictionary"): Set counts = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
  For r = 2 To UBound(data, 1)
    For j = 1 To 3
      If IsError(data(r, j)) Or IsEmpty(data(r, j)) Then GoTo Missing
      If Not IsNumeric(data(r, j)) Then GoTo Missing
      If CDbl(data(r, j)) <> Fix(CDbl(data(r, j))) Then GoTo Missing
    Next j
    If CLng(data(r, 1)) = stageNo Then
      e = CLng(data(r, 2)): gp = CLng(data(r, 3)): k = CStr(e) & ":" & CStr(gp)
      If gp < 1 Or gp > 4 Or e < 1 Or seen.Exists(k) Then Err.Raise vbObjectError + 3911, , "増分結果の積分点番号・重複を確認してください。"
      seen(k) = True
      For j = 6 To 20
        If IsError(data(r, j)) Or IsEmpty(data(r, j)) Then GoTo Missing
        If Not IsNumeric(data(r, j)) Then GoTo Missing
        If Abs(CDbl(data(r, j))) > 1E+100 Then GoTo Missing
      Next j
      If CDbl(data(r, 13)) < 0# Or CDbl(data(r, 14)) < 0# Then GoTo Missing
      If CDbl(data(r, 19)) <> Fix(CDbl(data(r, 19))) Then GoTo Missing
      If Not frameReady Then frame = Array(CDbl(data(r, 18)), CDbl(data(r, 19)), CDbl(data(r, 20))): frameReady = True
      If frame(0) <> CDbl(data(r, 18)) Or frame(1) <> CDbl(data(r, 19)) Or frame(2) <> CDbl(data(r, 20)) Then GoTo Missing
      If Not result.Exists(e) Then result(e) = Array(0#, 0#, 0#, 0#, 0#, CDbl(data(r, 18)), CLng(data(r, 19)), CDbl(data(r, 20))): counts(e) = 0
      v = result(e)
      If CDbl(data(r, 13)) > v(0) Then v(0) = CDbl(data(r, 13))
      If CDbl(data(r, 14)) > v(1) Then v(1) = CDbl(data(r, 14))
      For j = 2 To 4: v(j) = v(j) + CDbl(data(r, j + 7)) * 0.25: Next j
      If v(5) <> CDbl(data(r, 18)) Or v(6) <> CLng(data(r, 19)) Or v(7) <> CDbl(data(r, 20)) Then GoTo Missing
      result(e) = v: counts(e) = counts(e) + 1
    End If
  Next r
  For Each v In result.Keys
    If counts(v) <> 4 Then GoTo Missing
  Next v
  If result.Count = 0 Then GoTo Missing
  Set FEMIncrementRead = result
  Exit Function
Missing:
  Err.Raise vbObjectError + 3912, , "この結果には表示に必要な最終収束増分がありません。新しい版で最初から再解析してください（結果増分）。"
End Function
