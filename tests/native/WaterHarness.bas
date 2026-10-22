Option Explicit

Public Function GWTestDarcy() As String
  Dim conduct(0 To 1) As Double, fixed(0 To 12) As Boolean, prescribed(0 To 12) As Double
  Dim h() As Double, fx() As Double, fy() As Double, reaction() As Double
  Dim rr As Double, balance As Double, i As Long, e As Long, g As Long, expected As Double
  Dim nodeX(0 To 12) As Double
  On Error GoTo Failed
  GWBuildTwoElementMesh
  nodeX(0) = 0#: nodeX(1) = 1#: nodeX(2) = 2#: nodeX(3) = 0#: nodeX(4) = 1#: nodeX(5) = 2#
  nodeX(6) = 0.5: nodeX(7) = 1#: nodeX(8) = 0.5: nodeX(9) = 0#: nodeX(10) = 1.5: nodeX(11) = 2#: nodeX(12) = 1.5
  conduct(0) = 0.00001: conduct(1) = 0.00001
  fixed(0) = True: fixed(3) = True: fixed(9) = True
  fixed(2) = True: fixed(5) = True: fixed(11) = True
  prescribed(0) = 3#: prescribed(3) = 3#: prescribed(9) = 3#
  prescribed(2) = 2#: prescribed(5) = 2#: prescribed(11) = 2#
  GWSeepSolve conduct, fixed, prescribed, h, fx, fy, reaction, rr, balance
  For i = 0 To 12
    expected = 3# - 0.5# * nodeX(i)
    If Abs(h(i) - expected) > 1E-10 Then GWFailTest "head"
  Next i
  For e = 0 To 1
    For g = 0 To 3
      If Abs(fx(g, e) - 0.000005) > 1E-12 Then GWFailTest "fluxX"
      If Abs(fy(g, e)) > 1E-12 Then GWFailTest "fluxY"
    Next g
  Next e
  If rr > 1E-10 Then GWFailTest "residual"
  If balance > 1E-10 Then GWFailTest "balance"
  GWTestDarcy = "PASS h0=" & Format$(h(0), "0.000000") & " h2=" & Format$(h(2), "0.000000") & _
                " qx=" & Format$(fx(0, 0), "0.000000E+00") & " rr=" & Format$(rr, "0.000E+00") & _
                " bal=" & Format$(balance, "0.000E+00")
  Exit Function
Failed:
  GWTestDarcy = "FAIL " & Err.Description
End Function

Public Function GWTestDarcyCases() As String
  Dim conduct(0 To 1) As Double, fixed(0 To 12) As Boolean, prescribed(0 To 12) As Double
  Dim h() As Double, fx() As Double, fy() As Double, reaction() As Double
  Dim rr As Double, balance As Double, i As Long, rejected As Boolean
  On Error GoTo Failed
  GWBuildTwoElementMesh
  conduct(0) = 0.00001: conduct(1) = 0.00001
  For i = 0 To 12
    fixed(i) = True: prescribed(i) = 7#
  Next i
  GWSeepSolve conduct, fixed, prescribed, h, fx, fy, reaction, rr, balance
  If rr > 1E-10 Then GWFailTest "all-dirichlet residual"
  For i = 0 To 12
    If Abs(h(i) - 7#) > 1E-10 Then GWFailTest "all-dirichlet head"
  Next i
  GWBuildTwoElementMesh
  conduct(0) = 0.00001: conduct(1) = 0#
  For i = 0 To 12
    fixed(i) = False
    prescribed(i) = 0#
  Next i
  rejected = False
  On Error Resume Next
  GWSeepSolve conduct, fixed, prescribed, h, fx, fy, reaction, rr, balance
  rejected = (Err.Number <> 0): Err.Clear
  On Error GoTo Failed
  If Not rejected Then GWFailTest "no datum"
  GWBuildTwoElementMesh
  conduct(0) = -0.00001: conduct(1) = 0.00001
  rejected = False
  On Error Resume Next
  GWSeepSolve conduct, fixed, prescribed, h, fx, fy, reaction, rr, balance
  rejected = (Err.Number <> 0): Err.Clear
  On Error GoTo Failed
  If Not rejected Then GWFailTest "negative conductivity"
  GWTestDarcyCases = "PASS allDirichlet/disconnected/negative"
  Exit Function
Failed:
  GWTestDarcyCases = "FAIL " & Err.Description
End Function

Private Sub GWBuildTwoElementMesh()
  Dim i As Long, e As Long, nodes As Variant
  NumberOfFreeNode = 13
  NumberOfNode = 13
  NumberOfElement = 2
  NumberOfMaterial = 1
  ReDim Elem(0 To 1)
  ReDim Material(0 To 0)
  Material(0).thickness = 1#
  nodes = Array(Array(0, 1, 4, 3, 6, 7, 8, 9), Array(1, 2, 5, 4, 10, 11, 12, 7))
  For e = 0 To 1
    Elem(e).MatNo = 0
    For i = 0 To 7
      Elem(e).node(i) = nodes(e)(i)
    Next i
  Next e
  GWSetNode Elem(0), 0, 0#, 0#: GWSetNode Elem(0), 1, 1#, 0#
  GWSetNode Elem(0), 2, 1#, 1#: GWSetNode Elem(0), 3, 0#, 1#
  GWSetNode Elem(0), 4, 0.5, 0#: GWSetNode Elem(0), 5, 1#, 0.5
  GWSetNode Elem(0), 6, 0.5, 1#: GWSetNode Elem(0), 7, 0#, 0.5
  GWSetNode Elem(1), 0, 1#, 0#: GWSetNode Elem(1), 1, 2#, 0#
  GWSetNode Elem(1), 2, 2#, 1#: GWSetNode Elem(1), 3, 1#, 1#
  GWSetNode Elem(1), 4, 1.5, 0#: GWSetNode Elem(1), 5, 2#, 0.5
  GWSetNode Elem(1), 6, 1.5, 1#: GWSetNode Elem(1), 7, 1#, 0.5
End Sub

Private Sub GWSetNode(ByRef element As Element_Data, ByVal localNode As Long, ByVal xValue As Double, ByVal yValue As Double)
  element.x(localNode) = xValue
  element.y(localNode) = yValue
End Sub

Private Sub GWFailTest(ByVal label As String)
  Err.Raise vbObjectError + 7921, "GWTestDarcy", label
End Sub

Public Function GWTestInvalidInputs() As String
  Dim ws As Worksheet, cases As Variant, item As Variant, saved As Variant, message As String, count As Long
  On Error GoTo Failed
  Set ws = ThisWorkbook.Worksheets("水条件")
  cases = Array(Array("B4", -1#), Array("B3", "BAD_MODE"), Array("F12", 1.5), Array("F12", ""), Array("B13", -1#), Array("B5", "MISSING"))
  For Each item In cases
    saved = ws.Range(CStr(item(0))).Value2
    ws.Range(CStr(item(0))).Value2 = item(1)
    message = ""
    On Error Resume Next
    GWPreviewCompute
    message = Err.Description: Err.Clear
    On Error GoTo Failed
    ws.Range(CStr(item(0))).Value2 = saved
    If Len(message) = 0 Then GWFailTest "accepted invalid " & CStr(item(0))
    count = count + 1
  Next item
  GWPreviewCompute
  ws.Range("B7").Value2 = GWInputSignature()
  GWTestInvalidInputs = "PASS rejected=" & CStr(count)
  Exit Function
Failed:
  GWTestInvalidInputs = "FAIL " & Err.Description
End Function

Public Function GWTestDarcyPatches() As String
  Dim k(0 To 1) As Double, fixed(0 To 12) As Boolean, prescribed(0 To 12) As Double
  Dim h() As Double, qx() As Double, qy() As Double, reaction() As Double, face() As Double
  Dim rr As Double, balance As Double, localError As Double, correction As Double, flow As Double
  Dim e As Long, j As Long, n As Long, testNo As Long, factor As Double
  On Error GoTo Failed
  For testNo = 1 To 2
    GWBuildTwoElementMesh
    factor = IIf(testNo = 1, 0.00001, 0.000000000000001)
    k(0) = factor: k(1) = factor: Material(0).thickness = 2.5
    For e = 0 To 1
      For j = 0 To 7
        Elem(e).x(j) = Elem(e).x(j) + 0.3 * Elem(e).y(j)
        n = Elem(e).node(j): fixed(n) = (n <> 7)
        prescribed(n) = 4# - 0.2 * Elem(e).x(j) + 0.1 * Elem(e).y(j)
      Next j
    Next e
    GWSeepSolve k, fixed, prescribed, h, qx, qy, reaction, rr, balance
    For e = 0 To 1
      For j = 0 To 3
        If Abs(qx(j, e) / factor - 0.2) > 0.000000001 Or Abs(qy(j, e) / factor + 0.1) > 0.000000001 Then GWFailTest "skew/scaled flux"
      Next j
    Next e
  Next testNo
  GWBuildTwoElementMesh
  Material(0).thickness = 1#
  k(0) = 0.00001: k(1) = 0.000001
  For n = 0 To 12: fixed(n) = False: prescribed(n) = 0#: Next n
  fixed(0) = True: fixed(3) = True: fixed(9) = True
  fixed(2) = True: fixed(5) = True: fixed(11) = True
  prescribed(0) = 3#: prescribed(3) = 3#: prescribed(9) = 3#
  prescribed(2) = 1#: prescribed(5) = 1#: prescribed(11) = 1#
  GWSeepSolve k, fixed, prescribed, h, qx, qy, reaction, rr, balance
  GWRecoverFaceFlows k, h, fixed, face, localError, correction
  flow = 2# / (1# / k(0) + 1# / k(1))
  For e = 0 To 1
    For j = 0 To 3
      If Abs(qx(j, e) - flow) > 0.000000000001 Or Abs(qy(j, e)) > 0.000000000001 Then GWFailTest "layered flux e=" & CStr(e) & " qx=" & CStr(qx(j, e)) & " qy=" & CStr(qy(j, e)) & " expected=" & CStr(flow) & " heads=" & Join(Array(h(0), h(1), h(2), h(3), h(4), h(5), h(6), h(7), h(8), h(9), h(10), h(11), h(12)), ",") & " rr=" & CStr(rr)
    Next j
  Next e
  If Abs(face(1, 0) - flow) > 0.000000000001 Or Abs(face(3, 1) + flow) > 0.000000000001 Or localError > 0.00000001 Then GWFailTest "conservative face flow"
  GWTestDarcyPatches = "PASS skew/K=1E-15/layered/thickness/face conservation"
  Exit Function
Failed:
  GWTestDarcyPatches = "FAIL " & Err.Description
End Function
