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
  If Not ws Is Nothing Then ws.cells.ClearContents
End Sub

Public Sub FEMIncrementCapture(ByVal loadFactor As Double)
  Dim e As Long, gp As Long, c As Long, j As Long, dof As Long, s As Double
  Dim zeroLocal As Boolean
  On Error GoTo Failed
  incReady = False
  If NumberOfElement < 1 Or lastDof < 0 Then Exit Sub
  ReDim incData(0 To 6, 0 To 3, 0 To NumberOfElement - 1)
  ReDim incActive(0 To NumberOfElement - 1)
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Not Elem(e).IsJoint Then
      incActive(e) = True
      ' Same local-displacement test as P3EvalElements and native material reuse.
      ' Do not report stale plastic increments when no material update occurred.
      zeroLocal = True
      For j = 0 To 15
        dof = Elem(e).elnode(j)
        If dof >= 0 Then
          If Abs(TDisp(dof) - P3CommittedDisp(dof)) > 1E-30 Then zeroLocal = False
        End If
      Next j
      For gp = 0 To 3
        For c = 0 To 2
          s = 0#
          For j = 0 To 15
            dof = Elem(e).elnode(j)
            If dof >= 0 Then s = s + Elem(e).bMat(c, j, gp) * (TDisp(dof) - P3CommittedDisp(dof))
          Next j
          incData(c, gp, e) = s
        Next c
        For c = 0 To 3
          If Not zeroLocal Then incData(c + 3, gp, e) = Elem(e).P2PlasticStrain(c, gp)
        Next c
      Next gp
    End If
  Next e
  incFs = P3CurrentStrengthFactor: incLambda = loadFactor: incNo = CurrentIncrement
  incReady = True: FEMIncrementLastError = ""
  Exit Sub
Failed:
  FEMIncrementLastError = "増分保存の準備に失敗: " & CStr(Err.number) & " " & Err.description
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
  FEMPlasticStateSave
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("結果増分"): On Error GoTo Failed
  If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("結果要素")): ws.name = "結果増分"
  stageNo = P3ResultStageNumber()
  If UCase$(P6ReadTextSetting("OUTPUT_STAGE_MODE", "FINAL")) <> "ALL" Then ws.cells.ClearContents
  headers = Array("ステージ番号", "要素番号", "積分点", "要素中心X", "要素中心Y", "全ひずみ増分X", "全ひずみ増分Y", "工学せん断ひずみ増分XY", "塑性ひずみ増分X", "塑性ひずみ増分Y", "塑性工学せん断ひずみ増分XY", "塑性ひずみ増分Z", "最大工学せん断ひずみ増分(3D)", "最大塑性工学せん断ひずみ増分(3D)", "面内塑性主ひずみ増分max", "面内塑性主ひずみ増分min", "面内max主軸角rad", "Fs", "荷重増分番号", "荷重係数")
  For c = 0 To UBound(headers): ws.cells(1, c + 1).value2 = headers(c): Next c
  ws.range("V1").value2 = "LAST_CONVERGED_INCREMENT_V1"
  If Not incReady Then
    ws.range("V2").value2 = "増分結果なし。新しい版で最初から再解析してください。" & FEMIncrementLastError
    Exit Sub
  End If
  ReDim rows(1 To NumberOfElement * 4, 1 To 20)
  For e = 0 To NumberOfElement - 1
    If incActive(e) Then
      x = 0#: y = 0#
      If FEMElementIsT6(e) Then
        For c = 0 To 2: x = x + Elem(e).x(c) / 3#: y = y + Elem(e).y(c) / 3#: Next c
      Else
        For c = 0 To 3: x = x + Elem(e).x(c) * 0.25: y = y + Elem(e).y(c) * 0.25: Next c
      End If
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
  r = ws.cells(ws.rows.count, 1).End(xlUp).row + 1: If r < 2 Then r = 2
  If n > 0 Then ws.range(ws.cells(r, 1), ws.cells(r + n - 1, 20)).value2 = rows
  ws.range("V2").value2 = "最終収束荷重増分。SRMでは最後の成功Fsに対応。圧縮正、XYは工学せん断ひずみ。"
  ws.rows(1).Font.bOld = True: ws.columns("A:C").ColumnWidth = 14
  ws.range("F:T").NumberFormat = "0.00000E+00"
  Exit Sub
Failed:
  FEMIncrementLastError = "増分結果の出力に失敗: " & CStr(Err.number) & " " & Err.description
  If Not ws Is Nothing Then ws.range("V1").value2 = "INVALID": ws.range("V2").value2 = FEMIncrementLastError
End Sub

' Returns e -> Array(max shear, max plastic shear, mean plastic ex,ey,gxy, Fs,increment,lambda).
' All four Gauss points must exist exactly once; old/missing output is never fabricated.
Public Function FEMIncrementRead(ByVal stageNo As Long) As Object
  Dim ws As Worksheet, data As Variant, result As Object, counts As Object, seen As Object
  Dim r As Long, e As Long, gp As Long, k As String, v As Variant, j As Long
  Dim frame As Variant, frameReady As Boolean
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("結果増分"): On Error GoTo 0
  If ws Is Nothing Then GoTo Missing
  If CStr(ws.range("V1").value2) <> "LAST_CONVERGED_INCREMENT_V1" Then GoTo Missing
  r = ws.cells(ws.rows.count, 1).End(xlUp).row
  If r < 2 Then GoTo Missing
  data = ws.range("A1:T" & CStr(r)).value2
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
      For j = 2 To 4: v(j) = v(j) + CDbl(data(r, j + 7)) * FEMGaussFraction(e - 1, gp - 1, True): Next j
      If v(5) <> CDbl(data(r, 18)) Or v(6) <> CLng(data(r, 19)) Or v(7) <> CDbl(data(r, 20)) Then GoTo Missing
      result(e) = v: counts(e) = counts(e) + 1
    End If
  Next r
  For Each v In result.keys
    If counts(v) <> 4 Then GoTo Missing
  Next v
  If result.count = 0 Then GoTo Missing
  Set FEMIncrementRead = result
  Exit Function
Missing:
  Err.Raise vbObjectError + 3912, , "この結果には表示に必要な最終収束増分がありません。新しい版で最初から再解析してください（結果増分）。"
End Function

' ===== 20261022: plastic state / cumulative plastic strain for the result display (post-processing only) =====
' Writes sheet 結果塑性 from the converged state already held in memory. Nothing here feeds the solver.
' State per Gauss point: 2 = yielded now (P2Yielded, same test as the plastic point count),
' 1 = plastic strain accumulated earlier but not yielded now, 0 = no plastic strain so far, -1 = no cumulative log.
Public Sub FEMPlasticStateSave()
  Dim ws As Worksheet, rows() As Variant, e As Long, gp As Long, c As Long, n As Long, r As Long, stageNo As Long
  Dim logReady As Boolean, coverage As String, cum(0 To 3) As Double, mult As Double, hasPast As Boolean, st As Long
  On Error GoTo Failed
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("結果塑性"): On Error GoTo Failed
  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = "結果塑性"
  End If
  stageNo = P3ResultStageNumber()
  If UCase$(P6ReadTextSetting("OUTPUT_STAGE_MODE", "FINAL")) <> "ALL" Then ws.cells.ClearContents
  ws.range("A1:M1").value2 = Array("ステージ番号", "要素番号", "積分点", "降伏関数f", "現在降伏(1/0)", "累積塑性ひずみX", "累積塑性ひずみY", "累積塑性工学せん断ひずみXY", "累積塑性ひずみZ", "累積塑性乗数", "累積最大塑性工学せん断ひずみ(3D)", "塑性状態(2現在降伏/1過去に降伏/0弾性/-1不明)", "累積の範囲")
  ws.range("V1").value2 = "PLASTIC_STATE_V1"
  logReady = PlasticLogIsReady()
  If Not logReady Then
    coverage = "NONE"
  ElseIf PlasticLogIsComplete() Then
    coverage = "FROM_ZERO_STATE"
  Else
    coverage = "PARTIAL_FROM_RESUME"
  End If
  If NumberOfElement < 1 Then Exit Sub
  ReDim rows(1 To NumberOfElement * 4, 1 To 13)
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Not Elem(e).IsJoint Then
      For gp = 0 To 3
        n = n + 1
        For c = 0 To 3: cum(c) = PlasticLogTotal(c, gp, e): Next c
        mult = PlasticLogMultiplierTotal(gp, e)
        hasPast = (mult > 0#) Or cum(0) <> 0# Or cum(1) <> 0# Or cum(2) <> 0# Or cum(3) <> 0#
        If Elem(e).P2Yielded(gp) Then
          st = 2
        ElseIf Not logReady Then
          st = -1
        ElseIf hasPast Then
          st = 1
        Else
          st = 0
        End If
        rows(n, 1) = stageNo: rows(n, 2) = e + 1: rows(n, 3) = gp + 1
        rows(n, 4) = Elem(e).Stmat(9, gp): rows(n, 5) = IIf(Elem(e).P2Yielded(gp), 1, 0)
        For c = 0 To 3: rows(n, 6 + c) = cum(c): Next c
        rows(n, 10) = mult
        rows(n, 11) = FEMStrainMaxShear(cum(0), cum(1), cum(2), cum(3))
        rows(n, 12) = st: rows(n, 13) = coverage
      Next gp
    End If
  Next e
  r = ws.cells(ws.rows.count, 1).End(xlUp).row + 1: If r < 2 Then r = 2
  If n > 0 Then ws.range(ws.cells(r, 1), ws.cells(r + n - 1, 13)).value2 = rows
  ws.range("V2").value2 = "表示している工程の最終状態。現在降伏=降伏点数と同じ判定（P2Yielded）。累積はBIRTH・RESETで0から。"
  ws.range("D:K").NumberFormat = "0.00000E+00"
  Exit Sub
Failed:
  FEMIncrementLastError = "塑性状態の出力に失敗: " & CStr(Err.number) & " " & Err.description
End Sub

' Returns e -> Array(state, cumulative max plastic shear 4-point max, yielded-now count, past count, unknown count)
Public Function FEMPlasticStateRead(ByVal stageNo As Long) As Object
  Dim ws As Worksheet, data As Variant, result As Object, r As Long, e As Long, v As Variant, st As Long
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("結果塑性"): On Error GoTo 0
  If ws Is Nothing Then GoTo Missing
  If CStr(ws.range("V1").value2) <> "PLASTIC_STATE_V1" Then GoTo Missing
  r = ws.cells(ws.rows.count, 1).End(xlUp).row
  If r < 2 Then GoTo Missing
  data = ws.range("A1:M" & CStr(r)).value2
  Set result = CreateObject("Scripting.Dictionary")
  For r = 2 To UBound(data, 1)
    If IsNumeric(data(r, 1)) And Not IsEmpty(data(r, 1)) Then
      If CLng(data(r, 1)) = stageNo Then
        e = CLng(data(r, 2)): st = CLng(data(r, 12))
        If Not result.Exists(e) Then result(e) = Array(-2&, 0#, 0&, 0&, 0&)
        v = result(e)
        If CDbl(data(r, 11)) > v(1) Then v(1) = CDbl(data(r, 11))
        Select Case st
          Case 2: v(2) = v(2) + 1
          Case 1: v(3) = v(3) + 1
          Case -1: v(4) = v(4) + 1
        End Select
        result(e) = v
      End If
    End If
  Next r
  For Each v In result.keys
    e = CLng(v): data = result(e)
    If data(2) > 0 Then
      data(0) = 2
    ElseIf data(4) > 0 Then
      data(0) = -1
    ElseIf data(3) > 0 Then
      data(0) = 1
    Else
      data(0) = 0
    End If
    result(e) = data
  Next v
  If result.count = 0 Then GoTo Missing
  Set FEMPlasticStateRead = result
  Exit Function
Missing:
  Err.Raise vbObjectError + 3914, , "この結果には塑性状態の保存がありません。20261022版で最初から再解析してください（結果塑性）。"
End Function
