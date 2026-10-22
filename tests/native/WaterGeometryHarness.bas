Option Explicit

Public Function GWTestWetGeometry() As String
  Dim px(0 To 1) As Double, py(0 To 1) As Double, f() As Double
  Dim side As Long, i As Long, fx As Double, fy As Double, moment As Double
  On Error GoTo Failed
  GWBuildGeometryElement
  px(0) = -1#: px(1) = 3#: py(0) = 3#: py(1) = 3#
  ReDim f(0 To 15)
  For side = 0 To 3
    GWWetEdgeLoad 0, side, px, py, 10#, f
  Next side
  For i = 0 To 7
    fx = fx + f(2 * i): fy = fy + f(2 * i + 1)
    moment = moment + GWNodeX(i) * f(2 * i + 1) - GWNodeY(i) * f(2 * i)
  Next i
  If Abs(fx) > 1E-10 Or Abs(fy - 60#) > 1E-10 Then GWGeometryFail "closed resultant"
  If Abs(moment - 60#) > 1E-10 Then GWGeometryFail "closed moment"
  ReDim f(0 To 15): py(0) = 0.5: py(1) = 0.5
  GWWetEdgeLoad 0, 1, px, py, 10#, f
  fx = 0#: moment = 0#
  For i = 0 To 7
    fx = fx + f(2 * i)
    moment = moment + GWNodeX(i) * f(2 * i + 1) - GWNodeY(i) * f(2 * i)
  Next i
  If Abs(fx + 3.75) > 1E-10 Then GWGeometryFail "partial force"
  If Abs(moment - 0.625) > 1E-10 Then GWGeometryFail "partial moment"
  GWTestWetGeometry = "PASS closedFy=" & Format$(fy, "0.000000") & " partialFx=" & Format$(fx, "0.000000")
  Exit Function
Failed:
  GWTestWetGeometry = "FAIL " & Err.Description
End Function

Private Sub GWBuildGeometryElement()
  Dim i As Long, x As Variant, y As Variant
  NumberOfElement = 1: NumberOfFreeNode = 8: NumberOfMaterial = 1
  lastDof = 15: nDof = 16
  ReDim Elem(0 To 0): ReDim Material(0 To 0)
  Material(0).thickness = 3#
  x = Array(0#, 2#, 2#, 0#, 1#, 2#, 1#, 0#)
  y = Array(0#, 0#, 1#, 1#, 0#, 0.5, 1#, 0.5)
  For i = 0 To 7: Elem(0).node(i) = i: Elem(0).x(i) = x(i): Elem(0).y(i) = y(i): Elem(0).ElNode(2 * i) = 2 * i: Elem(0).ElNode(2 * i + 1) = 2 * i + 1: Next i
End Sub

Private Function GWNodeX(ByVal i As Long) As Double
  GWNodeX = Elem(0).x(i)
End Function

Private Function GWNodeY(ByVal i As Long) As Double
  GWNodeY = Elem(0).y(i)
End Function

Private Sub GWGeometryFail(ByVal messageText As String)
  Err.Raise vbObjectError + 7940, "GWTestWetGeometry", messageText
End Sub
