Option Explicit

Public Sub GWWetEdgeLoad(ByVal e As Long, ByVal side As Long, ByRef profileX() As Double, ByRef profileY() As Double, ByVal gamma As Double, ByRef force() As Double)
  Dim a As Long, b As Long, m As Long, i As Long, q As Long, k As Long, nodeDof As Long, nBreak As Long
  Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, xm As Double, ym As Double, scaleValue As Double
  Dim t As Double, z As Double, t0 As Double, t1 As Double, h As Double, p As Double, w As Double, thickness As Double
  Dim gx As Double, gy As Double, xx As Double, yy As Double, dx As Double, dy As Double
  Dim edgeNodes As Variant
  Dim breaks() As Double, px As Double, py As Double, Root As Double, lo As Double, hi As Double
  Dim nn(0 To 2) As Double, gauss(0 To 2) As Double, weights(0 To 2) As Double
  If e < 0 Or e >= NumberOfElement Then GWGeomFail "element out of range"
  If side < 0 Or side > 3 Then GWGeomFail "side out of range"
  If gamma <= 0# Or Not P2IsFinite(gamma) Then GWGeomFail "invalid water unit weight"
  a = side: b = (side + 1) Mod 4: m = side + 4
  edgeNodes = Array(a, b, m)
  x1 = Elem(e).x(a): y1 = Elem(e).y(a): x2 = Elem(e).x(b): y2 = Elem(e).y(b)
  xm = Elem(e).x(m): ym = Elem(e).y(m)
  scaleValue = Abs(x2 - x1) + Abs(y2 - y1) + 1#
  If Sqr((xm - 0.5 * (x1 + x2)) ^ 2 + (ym - 0.5 * (y1 + y2)) ^ 2) > 0.000000001 * scaleValue Then GWGeomFail "curved Q8 wetted edge is unsupported"
  If UBound(profileX) < 1 Or UBound(profileY) < 1 Then GWGeomFail "water profile needs two points"
  If UBound(profileY) <> UBound(profileX) Then GWGeomFail "water profile arrays differ"
  If Not P2IsFinite(profileX(LBound(profileX))) Or Not P2IsFinite(profileY(LBound(profileY))) Then GWGeomFail "water profile is nonfinite"
  For i = LBound(profileX) + 1 To UBound(profileX)
    If Not P2IsFinite(profileX(i)) Or Not P2IsFinite(profileY(i)) Then GWGeomFail "water profile is nonfinite"
    If profileX(i) <= profileX(i - 1) Then GWGeomFail "water profile X is not ascending"
  Next i
  If profileX(LBound(profileX)) > IIf(x1 < x2, x1, x2) + 0.000000001 Then GWGeomFail "water profile does not cover edge"
  If profileX(UBound(profileX)) < IIf(x1 > x2, x1, x2) - 0.000000001 Then GWGeomFail "water profile does not cover edge"
  ReDim breaks(0 To 1): breaks(0) = 0#: breaks(1) = 1#: nBreak = 2
  For i = LBound(profileX) To UBound(profileX)
    If Abs(x2 - x1) > 1E-30 Then
      t = (profileX(i) - x1) / (x2 - x1)
      If t > 0.000000000001 And t < 1# - 0.000000000001 Then GWAddBreak breaks, nBreak, t
    End If
  Next i
  Call GWSortBreaks(breaks, nBreak)
  Dim intervalCount As Long
  intervalCount = nBreak - 1
  For k = 0 To intervalCount - 1
    t0 = breaks(k): t1 = breaks(k + 1)
    lo = GWDepthAt(t0, x1, y1, x2, y2, profileX, profileY)
    hi = GWDepthAt(t1, x1, y1, x2, y2, profileX, profileY)
    If lo * hi < 0# Then
      Root = t0 - lo * (t1 - t0) / (hi - lo)
      If Root > t0 + 0.000000000001 And Root < t1 - 0.000000000001 Then GWAddBreak breaks, nBreak, Root
    End If
  Next k
  Call GWSortBreaks(breaks, nBreak)
  Call GWDedupBreaks(breaks, nBreak)
  gauss(0) = -Sqr(0.6): gauss(1) = 0#: gauss(2) = Sqr(0.6)
  weights(0) = 5# / 9#: weights(1) = 8# / 9#: weights(2) = 5# / 9#
  If UBound(force) < lastDof Then GWGeomFail "force array is too small"
  thickness = Material(Elem(e).MatNo).thickness
  If thickness <= 0# Then GWGeomFail "invalid element thickness"
  For k = 0 To nBreak - 2
    t0 = breaks(k): t1 = breaks(k + 1)
    For q = 0 To 2
      t = 0.5 * (t0 + t1) + 0.5 * (t1 - t0) * gauss(q)
      h = GWDepthAt(t, x1, y1, x2, y2, profileX, profileY): If h <= 0# Then GoTo NextGauss
      p = gamma * h: w = weights(q) * 0.5 * (t1 - t0) * thickness
      z = 2# * t - 1#
      nn(0) = 0.5 * z * (z - 1#): nn(1) = 0.5 * z * (z + 1#): nn(2) = 1# - z * z
      dx = x2 - x1: dy = y2 - y1
      For i = 0 To 2
        nodeDof = Elem(e).elnode(2 * CLng(edgeNodes(i)))
        If nodeDof < 0 Or nodeDof > lastDof Then GWGeomFail "edge DOF out of range"
        force(nodeDof) = force(nodeDof) - nn(i) * p * dy * w
        force(nodeDof + 1) = force(nodeDof + 1) + nn(i) * p * dx * w
      Next i
NextGauss:
    Next q
  Next k
End Sub

Public Function GWBoundaryLoops(ByVal edgeCounts As Object, ByVal owners As Object) As Object
  Dim out As Object, outgoing As Object, incoming As Object, edgeAt As Object, edgeList As Collection, key As Variant, own As Variant
  Dim a As Long, b As Long, startNode As Long, node As Long, nextNode As Long, steps As Long, area As Double, loopKey As Variant
  Dim loopKeys As Collection
  Set out = CreateObject("Scripting.Dictionary"): Set outgoing = CreateObject("Scripting.Dictionary")
  Set incoming = CreateObject("Scripting.Dictionary"): Set edgeAt = CreateObject("Scripting.Dictionary"): Set edgeList = New Collection
  For Each key In edgeCounts.keys
    If CLng(edgeCounts(key)) = 1 Then
      own = owners(key): a = GJTopoNode(Elem(CLng(own(0))).node(CLng(own(1)))): b = GJTopoNode(Elem(CLng(own(0))).node((CLng(own(1)) + 1) Mod 4))
      If outgoing.Exists(CStr(a)) Or incoming.Exists(CStr(b)) Then GWGeomFail "boundary branch"
      outgoing.Add CStr(a), b: incoming.Add CStr(b), a: edgeAt.Add CStr(a), CStr(key): edgeList.Add CStr(key)
    End If
  Next key
  For Each key In edgeList
    If out.Exists(CStr(key)) Then GoTo NextLoop
    Set loopKeys = New Collection
    own = owners(key): startNode = GJTopoNode(Elem(CLng(own(0))).node(CLng(own(1)))): node = startNode: area = 0#: steps = 0
    Do
      If Not outgoing.Exists(CStr(node)) Then GWGeomFail "open boundary loop"
      loopKey = edgeAt(CStr(node)): loopKeys.Add loopKey: out(loopKey) = False
      own = owners(loopKey): nextNode = CLng(outgoing(CStr(node)))
      area = area + Elem(CLng(own(0))).x(CLng(own(1))) * Elem(CLng(own(0))).y((CLng(own(1)) + 1) Mod 4) - Elem(CLng(own(0))).x((CLng(own(1)) + 1) Mod 4) * Elem(CLng(own(0))).y(CLng(own(1)))
      node = nextNode: steps = steps + 1
      If steps > edgeList.count Then GWGeomFail "boundary loop does not close"
    Loop While node <> startNode
    If Abs(area) < 1E-20 Then GWGeomFail "zero-area boundary loop"
    For Each loopKey In loopKeys
      out(loopKey) = (area < 0#)
    Next loopKey
NextLoop:
  Next key
  Set GWBoundaryLoops = out
End Function

Private Function GWDepthAt(ByVal t As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByRef px() As Double, ByRef py() As Double) As Double
  Dim x As Double, i As Long, y As Double
  x = x1 + t * (x2 - x1)
  If x < px(LBound(px)) - 0.000000001 Or x > px(UBound(px)) + 0.000000001 Then GWGeomFail "water profile does not cover edge"
  For i = LBound(px) + 1 To UBound(px)
    If x <= px(i) Then
      y = py(i - 1) + (py(i) - py(i - 1)) * (x - px(i - 1)) / (px(i) - px(i - 1))
      GWDepthAt = y - (y1 + t * (y2 - y1))
      Exit Function
    End If
  Next i
  GWDepthAt = py(UBound(py)) - (y1 + t * (y2 - y1))
End Function

Private Sub GWDedupBreaks(ByRef values() As Double, ByRef count As Long)
  Dim i As Long, writeAt As Long
  writeAt = 1
  For i = 1 To count - 1
    If Abs(values(i) - values(writeAt - 1)) > 0.000000000001 Then
      values(writeAt) = values(i)
      writeAt = writeAt + 1
    End If
  Next i
  count = writeAt
  ReDim Preserve values(0 To count - 1)
End Sub

Private Sub GWAddBreak(ByRef values() As Double, ByRef count As Long, ByVal value As Double)
  ReDim Preserve values(0 To count): values(count) = value: count = count + 1
End Sub

Private Sub GWSortBreaks(ByRef values() As Double, ByVal count As Long)
  Dim i As Long, j As Long, t As Double
  For i = 0 To count - 2
    For j = i + 1 To count - 1
      If values(j) < values(i) Then
        t = values(i): values(i) = values(j): values(j) = t
      End If
    Next j
  Next i
End Sub

Private Function GWLoopArea(ByVal edgeCounts As Object, ByVal owners As Object, ByVal edges As Collection, ByVal startNode As Long) As Double
  Dim key As Variant, own As Variant, node As Long, nxt As Long, area As Double
  Dim outgoing As Object, nodeX As Object, nodeY As Object
  Set outgoing = CreateObject("Scripting.Dictionary")
  Set nodeX = CreateObject("Scripting.Dictionary")
  Set nodeY = CreateObject("Scripting.Dictionary")
  For Each key In edges
    own = owners(key): node = Elem(CLng(own(0))).node(CLng(own(1)))
    nxt = Elem(CLng(own(0))).node((CLng(own(1)) + 1) Mod 4)
    outgoing(CStr(node)) = nxt
    nodeX(CStr(node)) = Elem(CLng(own(0))).x(CLng(own(1)))
    nodeY(CStr(node)) = Elem(CLng(own(0))).y(CLng(own(1)))
    nodeX(CStr(nxt)) = Elem(CLng(own(0))).x((CLng(own(1)) + 1) Mod 4)
    nodeY(CStr(nxt)) = Elem(CLng(own(0))).y((CLng(own(1)) + 1) Mod 4)
  Next key
  node = startNode
  Do
    nxt = CLng(outgoing(CStr(node)))
    area = area + CDbl(nodeX(CStr(node))) * CDbl(nodeY(CStr(nxt))) - CDbl(nodeX(CStr(nxt))) * CDbl(nodeY(CStr(node)))
    node = nxt
  Loop While node <> startNode
  If Abs(area) < 1E-20 Then GWGeomFail "zero-area boundary loop"
  GWLoopArea = 0.5 * area
End Function

Private Sub GWGeomFail(ByVal messageText As String)
  Err.Raise vbObjectError + 7930, "FEMWaterGeometry", messageText
End Sub
