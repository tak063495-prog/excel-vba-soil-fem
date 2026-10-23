Option Explicit

Public Sub GWDrawWaterView(ByVal ws As Worksheet, ByRef pressures() As Double, ByRef surfaceForce() As Double, ByVal titleText As String)
  Dim e As Long, i As Long, g As Long, n As Long, shapeCount As Long, maxP As Double, maxF As Double, forceNodes As Long, profilePoints As Long
  Dim minX As Double, maxX As Double, minY As Double, maxY As Double, x0 As Double, y0 As Double, sx As Double, sy As Double
  Dim nodeX() As Double, nodeY() As Double, haveNode() As Boolean, p As Double, value As Double, px As Double, py As Double
  Dim sh As shape, key As Variant, prof As String, r As Long, x As Double, y As Double, wsIn As Worksheet
  Dim drawStride As Long, sampleNote As String, canvasTop As Double, canvasHeight As Double, plotHeight As Double
  Dim mode As String, originText As String
  mode = GWViewMode()
  If mode = "FIXED_STRESS" Then
    originText = "時間連成：受理時刻t=" & Format$(FSTime, "0.###") & " sの計算水圧・全水頭・流束。入力水位は境界条件で、内部過剰水圧は固定していません。"
  ElseIf mode = "STEADY" Then
    originText = "浸透計算済：h・qから水圧を算出。破線は入力境界用水位で、自由水面の計算結果ではありません。"
  Else
    originText = "指定場：入力水位から静水圧を換算。浸透計算・水位の再計算は行いません。"
  End If
  If NumberOfElement < 1 Or NumberOfFreeNode < 1 Then Err.Raise vbObjectError + 7950, "GWDrawWaterView", "No mesh to draw."
  If UBound(pressures, 1) < 3 Or UBound(pressures, 2) < NumberOfElement - 1 Then Err.Raise vbObjectError + 7950, "GWDrawWaterView", "Pressure array is too small."
  If UBound(surfaceForce) < NumberOfFreeNode * 2 - 1 Then Err.Raise vbObjectError + 7950, "GWDrawWaterView", "Surface force array is too small."
  ReDim nodeX(0 To NumberOfFreeNode - 1): ReDim nodeY(0 To NumberOfFreeNode - 1): ReDim haveNode(0 To NumberOfFreeNode - 1)
  maxP = 0#: maxF = 0#: shapeCount = 0
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Not Elem(e).IsJoint Then
      shapeCount = shapeCount + 5
      For i = 0 To 7
        n = Elem(e).node(i)
        If n >= 0 And n < NumberOfFreeNode Then
          nodeX(n) = Elem(e).x(i): nodeY(n) = Elem(e).y(i): haveNode(n) = True
        End If
      Next i
      For g = 0 To 3
        p = pressures(g, e): If Abs(p) > maxP Then maxP = Abs(p)
      Next g
    End If
  Next e
  For n = 0 To NumberOfFreeNode - 1
    value = Sqr(surfaceForce(2 * n) ^ 2 + surfaceForce(2 * n + 1) ^ 2)
    If value > maxF Then maxF = value
  Next n
  If shapeCount = 0 Then
    GWClearWaterView ws
    ws.range("A8").value2 = "初期の有効要素はありません。BIRTH後に水圧場を再計算します。"
    Exit Sub
  End If
  GWViewBounds nodeX, nodeY, haveNode, minX, maxX, minY, maxY
  Set wsIn = Nothing
  On Error Resume Next
  Set wsIn = ThisWorkbook.Worksheets("水条件")
  On Error GoTo 0
  If Not wsIn Is Nothing Then
    For r = 12 To 211
      If Len(Trim$(CStr(wsIn.cells(r, 1).value2))) > 0 Then
        profilePoints = profilePoints + 1
        If IsNumeric(wsIn.cells(r, 2).value2) And IsNumeric(wsIn.cells(r, 3).value2) Then GWExpandBounds CDbl(wsIn.cells(r, 2).value2), CDbl(wsIn.cells(r, 3).value2), minX, maxX, minY, maxY
      End If
    Next r
  End If
  For n = 0 To NumberOfFreeNode - 1
    If haveNode(n) Then forceNodes = forceNodes + 1
  Next n
  If mode = "STEADY" Or mode = "FIXED_STRESS" Then shapeCount = shapeCount * 10
  shapeCount = shapeCount + forceNodes + profilePoints + 12
  drawStride = 1
  If shapeCount > 3500 Then
    drawStride = (shapeCount + 3499) \ 3500
    sampleNote = " （概観表示: 要素1/" & CStr(drawStride) & "、数値表は全点）"
  End If
  GWClearWaterView ws
  canvasTop = ws.range("A8").Top
  canvasHeight = ws.range("A35").Top - canvasTop - 12#
  plotHeight = canvasHeight - 165#
  If plotHeight < 90# Then Err.Raise vbObjectError + 7950, , "水圧図の表示領域が狭すぎます。8～34行の高さを戻してください。"
  x0 = ws.range("A8").left + 20#: y0 = canvasTop + plotHeight + 35#
  GWViewScale minX, maxX, minY, maxY, 600#, plotHeight, sx, sy
  Set sh = ws.Shapes.AddShape(msoShapeRectangle, ws.range("A8").left, canvasTop, 680#, canvasHeight)
  sh.name = "GW_VIEW_CANVAS": sh.fill.ForeColor.RGB = RGB(255, 255, 255): sh.line.ForeColor.RGB = RGB(210, 210, 210)
  sh.ZOrder msoSendToBack
  Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x0, canvasTop + 2#, 650#, 22#)
  sh.name = "GW_VIEW_TITLE": sh.TextFrame2.TextRange.text = titleText & sampleNote
  ws.range("A4").value2 = originText
  ws.range("A5").value2 = "入力線：橙破線 / GP水圧：青点 / 水圧外力：橙矢印。数値は35行目以降。"
  ws.range("A6").value2 = "HYDROSTATIC=指定水位から換算 / STEADY=定常浸透 / FIXED_STRESS=飽和時間連成。未知自由水面は対象外。"
  ws.rows("4:6").RowHeight = 18
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Not Elem(e).IsJoint And e Mod drawStride = 0 Then GWDrawElement ws, e, x0, y0, sx, sy, minX, minY
  Next e
  GWDrawProfiles ws, x0, y0, sx, sy, minX, minY
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Not Elem(e).IsJoint Then
      For g = 0 To 3
        If e Mod drawStride <> 0 Then Exit For
        FEMGaussPhysicalXY e, g, px, py
        p = pressures(g, e): GWDrawPressurePoint ws, px, py, p, maxP, x0, y0, sx, sy, minX, minY, e, g
      Next g
    End If
  Next e
  If mode = "STEADY" Or mode = "FIXED_STRESS" Then GWDrawHydraulicSolution ws, drawStride, x0, y0, sx, sy, minX, minY
  GWDrawForces ws, nodeX, nodeY, haveNode, surfaceForce, maxF, x0, y0, sx, sy, minX, minY
  GWDrawLegend ws, x0, y0, maxP, maxF
End Sub

Public Sub GWClearWaterView(ByVal ws As Worksheet)
  Dim i As Long
  For i = ws.Shapes.count To 1 Step -1
    If left$(ws.Shapes(i).name, 8) = "GW_VIEW_" Then ws.Shapes(i).Delete
  Next i
End Sub

Private Sub GWDrawElement(ByVal ws As Worksheet, ByVal e As Long, ByVal x0 As Double, ByVal y0 As Double, ByVal sx As Double, ByVal sy As Double, ByVal minX As Double, ByVal minY As Double)
  Dim i As Long, a As Long, path As FreeformBuilder, sh As shape
  Set path = ws.Shapes.BuildFreeform(msoEditingAuto, x0 + (Elem(e).x(0) - minX) * sx, y0 + (minY - Elem(e).y(0)) * sy)
  For i = 0 To 3
    a = i + 4
    path.AddNodes msoSegmentLine, msoEditingAuto, x0 + (Elem(e).x(a) - minX) * sx, y0 + (minY - Elem(e).y(a)) * sy
    a = (i + 1) Mod 4
    path.AddNodes msoSegmentLine, msoEditingAuto, x0 + (Elem(e).x(a) - minX) * sx, y0 + (minY - Elem(e).y(a)) * sy
  Next i
  Set sh = path.ConvertToShape
  sh.name = "GW_VIEW_MESH_" & CStr(e): sh.fill.Visible = msoFalse: sh.line.ForeColor.RGB = RGB(150, 150, 150)
End Sub

Private Sub GWDrawPressurePoint(ByVal ws As Worksheet, ByVal px As Double, ByVal py As Double, ByVal p As Double, ByVal maxP As Double, ByVal x0 As Double, ByVal y0 As Double, ByVal sx As Double, ByVal sy As Double, ByVal minX As Double, ByVal minY As Double, ByVal elementId As Long, ByVal gaussId As Long)
  Dim sh As shape, ratio As Double, col As Long, d As Double
  ratio = 0#: If maxP > 0# Then ratio = Abs(p) / maxP
  col = RGB(255 - CLng(180# * ratio), 255 - CLng(180# * ratio), 255)
  d = 7#: Set sh = ws.Shapes.AddShape(msoShapeOval, x0 + (px - minX) * sx - d / 2#, y0 + (minY - py) * sy - d / 2#, d, d)
  sh.name = "GW_VIEW_GP_" & CStr(elementId) & "_" & CStr(gaussId): sh.fill.ForeColor.RGB = col: sh.line.Visible = msoFalse
End Sub

Private Sub GWDrawProfiles(ByVal ws As Worksheet, ByVal x0 As Double, ByVal y0 As Double, ByVal sx As Double, ByVal sy As Double, ByVal minX As Double, ByVal minY As Double)
  Dim r As Long, prof As String, prevProf As String, px As Double, py As Double, qx As Double, qy As Double, n As Long
  Dim sh As shape, wsIn As Worksheet
  On Error Resume Next: Set wsIn = ThisWorkbook.Worksheets("水条件"): On Error GoTo 0
  If wsIn Is Nothing Then Exit Sub
  For r = 12 To 211
    prof = Trim$(CStr(wsIn.cells(r, 1).value2))
    If Len(prof) > 0 And IsNumeric(wsIn.cells(r, 2).value2) And IsNumeric(wsIn.cells(r, 3).value2) Then
      qx = CDbl(wsIn.cells(r, 2).value2): qy = CDbl(wsIn.cells(r, 3).value2)
      If prof = prevProf Then
        GWLine ws, "GW_VIEW_PROFILE_" & r, x0 + (px - minX) * sx, y0 + (minY - py) * sy, x0 + (qx - minX) * sx, y0 + (minY - qy) * sy, RGB(195, 100, 0)
        ws.Shapes("GW_VIEW_PROFILE_" & r).line.DashStyle = msoLineDash
      Else
        n = n + 1
        GWProfileLabel ws, prof, qx, qy, x0, y0, sx, sy, minX, minY, n
      End If
      px = qx: py = qy: prevProf = prof
    End If
  Next r
End Sub

Private Sub GWProfileLabel(ByVal ws As Worksheet, ByVal textValue As String, ByVal x As Double, ByVal y As Double, ByVal x0 As Double, ByVal y0 As Double, ByVal sx As Double, ByVal sy As Double, ByVal minX As Double, ByVal minY As Double, ByVal indexValue As Long)
  Dim sh As shape
  Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x0 + (x - minX) * sx + 3#, y0 + (minY - y) * sy - 10#, 100#, 16#)
  sh.name = "GW_VIEW_PROFILE_LABEL_" & indexValue: sh.TextFrame2.TextRange.text = "入力線：" & textValue
  sh.TextFrame2.TextRange.Font.Size = 9
  sh.TextFrame2.TextRange.Font.fill.ForeColor.RGB = RGB(160, 75, 0)
End Sub

Private Sub GWDrawForces(ByVal ws As Worksheet, ByRef nodeX() As Double, ByRef nodeY() As Double, ByRef haveNode() As Boolean, ByRef force() As Double, ByVal maxF As Double, ByVal x0 As Double, ByVal y0 As Double, ByVal sx As Double, ByVal sy As Double, ByVal minX As Double, ByVal minY As Double)
  Dim n As Long, scaleF As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double
  If maxF <= 0# Then Exit Sub
  scaleF = 80# / maxF
  For n = 0 To UBound(haveNode)
    If haveNode(n) Then
      x1 = x0 + (nodeX(n) - minX) * sx: y1 = y0 + (minY - nodeY(n)) * sy
      x2 = x1 + force(2 * n) * scaleF: y2 = y1 - force(2 * n + 1) * scaleF
      If Abs(x2 - x1) + Abs(y2 - y1) > 0.01 Then
        GWLine ws, "GW_VIEW_FORCE_" & n, x1, y1, x2, y2, RGB(255, 180, 0)
        ws.Shapes("GW_VIEW_FORCE_" & n).line.EndArrowheadStyle = msoArrowheadTriangle
      End If
    End If
  Next n
End Sub

Private Sub GWDrawLegend(ByVal ws As Worksheet, ByVal x0 As Double, ByVal y0 As Double, ByVal maxP As Double, ByVal maxF As Double)
  Dim sh As shape
  Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x0, y0 + 40#, 630#, 70#)
  sh.name = "GW_VIEW_LEGEND"
  sh.TextFrame2.TextRange.Font.Size = 9
  sh.TextFrame2.TextRange.text = "橙破線＝登録した入力水位線（境界用。計算結果の水面ではない）" & vbLf & _
    "青点＝GP水圧 max=" & Format$(maxP, "0.000") & " kPa / 橙矢印＝水圧の等価節点荷重 max=" & Format$(maxF, "0.000") & " kN"
  If GWViewMode() = "STEADY" Or GWViewMode() = "FIXED_STRESS" Then
    sh.TextFrame2.TextRange.text = sh.TextFrame2.TextRange.text & vbLf & "紫実線＝計算全水頭h [m] の等値線（表示用補間） / 紫菱形＝固定水頭境界" & vbLf & _
      "緑矢印＝計算Darcy流束q [m/s]（外力矢印とは別）。飽和領域は入力メッシュで固定。"
  Else
    sh.TextFrame2.TextRange.text = sh.TextFrame2.TextRange.text & vbLf & "HYDROSTATIC：水圧p=γw max(入力水位-y,0)。水頭・流束は解いていません。"
  End If
End Sub

Private Sub GWDrawHydraulicSolution(ByVal ws As Worksheet, ByVal stride As Long, ByVal x0 As Double, ByVal y0 As Double, ByVal sx As Double, ByVal sy As Double, ByVal minX As Double, ByVal minY As Double)
  Dim e As Long, j As Long, g As Long, k As Long, a As Long, b As Long, n As Long, counter As Long
  Dim h(0 To 7) As Double, low As Double, high As Double, have As Boolean, locked As Boolean, value As Double
  Dim qx As Double, qy As Double, maxQ As Double, fx As Double, fy As Double, cx As Double, cy As Double, ch As Double
  Dim order As Variant, level As Double, seen As Object, labels As Object, levels As Object, key As String, sh As shape
  Set seen = CreateObject("Scripting.Dictionary"): Set labels = CreateObject("Scripting.Dictionary")
  Set levels = CreateObject("Scripting.Dictionary")
  order = Array(0, 4, 1, 5, 2, 6, 3, 7)
  For e = 0 To NumberOfElement - 1
    For j = 0 To 7
      If GWViewHead(e, j, value, locked) Then
        If Not have Then low = value: high = value: have = True
        If value < low Then low = value
        If value > high Then high = value
      End If
    Next j
    For g = 0 To 3
      If GWViewFlux(e, g, qx, qy) Then
        value = Sqr(qx * qx + qy * qy): If value > maxQ Then maxQ = value
      End If
    Next g
  Next e
  If Not have Then Exit Sub
  For e = 0 To NumberOfElement - 1
    If e Mod stride <> 0 Then GoTo NextHydraulicElement
    If Not GWViewHead(e, 0, value, locked) Then GoTo NextHydraulicElement
    cx = 0#: cy = 0#: ch = 0#: fx = 0#: fy = 0#
    For j = 0 To 7
      Call GWViewHead(e, j, h(j), locked)
      value = 0.5: If j < 4 Then value = -0.25
      If FEMElementIsT6(e) Then
        value = -1# / 9#
        If j = 4 Or j = 5 Or j = 7 Then value = 4# / 9#
        If j = 3 Or j = 6 Then value = 0#
      End If
      cx = cx + value * Elem(e).x(j): cy = cy + value * Elem(e).y(j): ch = ch + value * h(j)
      n = Elem(e).node(j): key = CStr(n)
      If locked And Not seen.Exists(key) Then
        seen.Add key, True
        Set sh = ws.Shapes.AddShape(msoShapeDiamond, x0 + (Elem(e).x(j) - minX) * sx - 3#, y0 + (minY - Elem(e).y(j)) * sy - 3#, 6#, 6#)
        sh.name = "GW_VIEW_HEAD_FIXED_" & key: sh.fill.ForeColor.RGB = RGB(130, 40, 170): sh.line.Visible = msoFalse
        key = Format$(h(j), "0.000000")
        If Not labels.Exists(key) Then
          labels.Add key, True
          Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x0 + (Elem(e).x(j) - minX) * sx + 5#, y0 + (minY - Elem(e).y(j)) * sy, 110#, 17#)
          sh.name = "GW_VIEW_HEAD_LABEL_" & CStr(labels.count): sh.TextFrame2.TextRange.text = "固定h=" & Format$(h(j), "0.###") & " m": sh.TextFrame2.TextRange.Font.Size = 9
        End If
      End If
    Next j
    If high - low > 0.000000001 Then
      For k = 1 To 5
        level = low + (high - low) * CDbl(k) / 6#
        For j = 0 To 7
          a = CLng(order(j)): b = CLng(order((j + 1) Mod 8))
          If Elem(e).node(a) <> Elem(e).node(b) Then
          GWHeadTriangle ws, cx, cy, ch, Elem(e).x(a), Elem(e).y(a), h(a), Elem(e).x(b), Elem(e).y(b), h(b), level, counter, levels, x0, y0, sx, sy, minX, minY
          End If
        Next j
      Next k
    End If
    For g = 0 To 3
      If GWViewFlux(e, g, qx, qy) Then fx = fx + qx * FEMGaussFraction(e, g): fy = fy + qy * FEMGaussFraction(e, g)
    Next g
    If high - low > 0.000000001 And maxQ > 1E-24 And Abs(fx) + Abs(fy) > maxQ * 0.000001 Then
      GWLine ws, "GW_VIEW_FLUX_" & CStr(e), x0 + (cx - minX) * sx, y0 + (minY - cy) * sy, x0 + (cx - minX) * sx + 30# * fx / maxQ, y0 + (minY - cy) * sy - 30# * fy / maxQ, RGB(0, 135, 70)
      ws.Shapes("GW_VIEW_FLUX_" & CStr(e)).line.EndArrowheadStyle = msoArrowheadTriangle
    End If
NextHydraulicElement:
  Next e
  Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x0, y0 + 20#, 630#, 18#)
  sh.name = "GW_VIEW_HEAD_RANGE": sh.TextFrame2.TextRange.Font.Size = 9
  sh.TextFrame2.TextRange.text = "計算h=" & Format$(low, "0.###") & "～" & Format$(high, "0.###") & " m / max|q|=" & Format$(maxQ, "0.00E+00") & " m/s"
  If high - low <= 0.000000001 Then sh.TextFrame2.TextRange.text = sh.TextFrame2.TextRange.text & "（一定水頭：流向矢印を省略）"
End Sub

Private Sub GWHeadTriangle(ByVal ws As Worksheet, ByVal ax As Double, ByVal ay As Double, ByVal ah As Double, ByVal bx As Double, ByVal by As Double, ByVal bh As Double, ByVal cx As Double, ByVal cy As Double, ByVal ch As Double, ByVal level As Double, ByRef counter As Long, ByVal levels As Object, ByVal x0 As Double, ByVal y0 As Double, ByVal sx As Double, ByVal sy As Double, ByVal minX As Double, ByVal minY As Double)
  Dim x As Variant, y As Variant, h As Variant, i As Long, j As Long, count As Long, t As Double, px(0 To 1) As Double, py(0 To 1) As Double
  x = Array(ax, bx, cx): y = Array(ay, by, cy): h = Array(ah, bh, ch)
  For i = 0 To 2
    j = (i + 1) Mod 3
    If (h(i) <= level And h(j) > level) Or (h(j) <= level And h(i) > level) Then
      t = (level - h(i)) / (h(j) - h(i))
      If count < 2 Then px(count) = x(i) + t * (x(j) - x(i)): py(count) = y(i) + t * (y(j) - y(i))
      count = count + 1
    End If
  Next i
  If count = 2 Then
    counter = counter + 1
    GWLine ws, "GW_VIEW_HEAD_CONTOUR_" & CStr(counter), x0 + (px(0) - minX) * sx, y0 + (minY - py(0)) * sy, x0 + (px(1) - minX) * sx, y0 + (minY - py(1)) * sy, RGB(130, 40, 170)
    ws.Shapes("GW_VIEW_HEAD_CONTOUR_" & CStr(counter)).AlternativeText = "計算全水頭h=" & Format$(level, "0.###") & " m"
    If Not levels.Exists(CStr(level)) Then
      levels.Add CStr(level), True
      Dim label As shape
      Set label = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x0 + ((px(0) + px(1)) / 2# - minX) * sx, y0 + (minY - (py(0) + py(1)) / 2#) * sy, 75#, 15#)
      label.name = "GW_VIEW_CONTOUR_VALUE_" & CStr(levels.count): label.TextFrame2.TextRange.text = Format$(level, "0.###") & " m"
      label.TextFrame2.TextRange.Font.Size = 8: label.TextFrame2.TextRange.Font.fill.ForeColor.RGB = RGB(130, 40, 170)
      label.fill.Visible = msoFalse: label.line.Visible = msoFalse
    End If
  End If
End Sub

Private Sub GWLine(ByVal ws As Worksheet, ByVal nameText As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal colorValue As Long)
  Dim sh As shape: Set sh = ws.Shapes.AddLine(x1, y1, x2, y2): sh.name = nameText: sh.line.ForeColor.RGB = colorValue: sh.line.weight = 1#
End Sub

Private Sub GWViewBounds(ByRef nx() As Double, ByRef ny() As Double, ByRef have() As Boolean, ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double)
  Dim i As Long, first As Boolean
  For i = LBound(have) To UBound(have)
    If have(i) Then
      If Not first Then
        minX = nx(i): maxX = nx(i): minY = ny(i): maxY = ny(i): first = True
      Else
        GWExpandBounds nx(i), ny(i), minX, maxX, minY, maxY
      End If
    End If
  Next i
  If Not first Then Err.Raise vbObjectError + 7950, "GWDrawWaterView", "No active mesh geometry."
End Sub

Private Sub GWExpandBounds(ByVal x As Double, ByVal y As Double, ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double)
  If x < minX Then minX = x
  If x > maxX Then maxX = x
  If y < minY Then minY = y
  If y > maxY Then maxY = y
End Sub

Private Sub GWViewScale(ByVal minX As Double, ByVal maxX As Double, ByVal minY As Double, ByVal maxY As Double, ByVal widthValue As Double, ByVal heightValue As Double, ByRef sx As Double, ByRef sy As Double)
  Dim dx As Double, dy As Double, s As Double
  dx = maxX - minX: dy = maxY - minY
  If dx < 0.000000000001 Then dx = 1#
  If dy < 0.000000000001 Then dy = 1#
  s = widthValue / dx
  If heightValue / dy < s Then s = heightValue / dy
  sx = s: sy = s
End Sub
