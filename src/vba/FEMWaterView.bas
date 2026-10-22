Option Explicit

Public Sub GWDrawWaterView(ByVal ws As Worksheet, ByRef pressures() As Double, ByRef surfaceForce() As Double, ByVal titleText As String)
  Dim e As Long, i As Long, g As Long, n As Long, shapeCount As Long, maxP As Double, maxF As Double, forceNodes As Long, profilePoints As Long
  Dim minX As Double, maxX As Double, minY As Double, maxY As Double, x0 As Double, y0 As Double, sx As Double, sy As Double
  Dim nodeX() As Double, nodeY() As Double, haveNode() As Boolean, p As Double, value As Double, px As Double, py As Double
  Dim sh As shape, key As Variant, prof As String, r As Long, x As Double, y As Double, wsIn As Worksheet
  Dim drawStride As Long, sampleNote As String, canvasTop As Double, canvasHeight As Double, plotHeight As Double
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
  shapeCount = shapeCount + forceNodes + profilePoints + 9
  drawStride = 1
  If shapeCount > 3500 Then
    drawStride = (shapeCount + 3499) \ 3500
    sampleNote = " （概観表示: 要素1/" & CStr(drawStride) & "、数値表は全点）"
  End If
  GWClearWaterView ws
  canvasTop = ws.range("A8").Top
  canvasHeight = ws.range("A35").Top - canvasTop - 12#
  plotHeight = canvasHeight - 90#
  If plotHeight < 90# Then Err.Raise vbObjectError + 7950, , "水圧図の表示領域が狭すぎます。8～34行の高さを戻してください。"
  x0 = ws.range("A8").left + 20#: y0 = canvasTop + plotHeight + 20#
  GWViewScale minX, maxX, minY, maxY, 600#, plotHeight, sx, sy
  Set sh = ws.Shapes.AddShape(msoShapeRectangle, ws.range("A8").left, canvasTop, 680#, canvasHeight)
  sh.name = "GW_VIEW_CANVAS": sh.fill.ForeColor.RGB = RGB(255, 255, 255): sh.line.ForeColor.RGB = RGB(210, 210, 210)
  sh.ZOrder msoSendToBack
  Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x0, ws.range("A8").Top - 25#, 680#, 22#)
  sh.name = "GW_VIEW_TITLE": sh.TextFrame2.TextRange.text = titleText & "  (水圧・等価節点荷重)" & sampleNote
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
        GWLine ws, "GW_VIEW_PROFILE_" & r, x0 + (px - minX) * sx, y0 + (minY - py) * sy, x0 + (qx - minX) * sx, y0 + (minY - qy) * sy, RGB(0, 150, 220)
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
  sh.name = "GW_VIEW_PROFILE_LABEL_" & indexValue: sh.TextFrame2.TextRange.text = textValue
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
  Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x0, y0 + 12#, 320#, 40#)
  sh.name = "GW_VIEW_LEGEND": sh.TextFrame2.TextRange.text = "GP水圧 max=" & Format$(maxP, "0.000") & " kPa / 等価荷重 max=" & Format$(maxF, "0.000") & " kN"
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
