Option Explicit
Private qLo As Double, qHi As Double, qAsp As Double, qJac As Double
Private tLo As Double, tHi As Double, tAsp As Double, scoreFloor As Double, crossFloor As Double
Public MeshT6PatchLastDetail As String

Private Function Key(ByVal a As Long, ByVal b As Long) As String
    If a < b Then Key = CStr(a) & ":" & CStr(b) Else Key = CStr(b) & ":" & CStr(a)
End Function
Private Function AcosDegrees(ByVal cs As Double) As Double
    If cs >= 1# Then Exit Function
    If cs <= -1# Then AcosDegrees = 180#: Exit Function
    AcosDegrees = (1.5707963267949 - Atn(cs / Sqr(1# - cs * cs))) * 57.2957795130823
End Function
Private Function Positive(ByVal v As Double) As Double
    If v > 0# Then Positive = v
End Function

' Straight-sided geometry only. Q8 guards use the exact corner Jacobians of
' the bilinear map, normalized by area. T6 targets are never relaxed.
Private Function ElementCost(ByRef x() As Double, ByRef y() As Double, ByRef rs() As Variant, ByVal e As Long, ByRef penalty As Double, ByRef score As Double, ByRef miss As Boolean) As Double
    Dim j As Long, k As Long, a As Long, b As Long, c As Long, triangle As Boolean
    Dim ux As Double, uy As Double, vx As Double, vy As Double, cross As Double, dot As Double, lu As Double, lv As Double, angle As Double
    Dim lo As Double, hi As Double, short As Double, longest As Double, area As Double, minCross As Double, sumLen As Double, ratio As Double
    Dim innerLo As Double, innerHi As Double, innerAsp As Double, cost As Double
    triangle = CStr(rs(e, 11)) = "T6": k = IIf(triangle, 3, 4)
    lo = 180#: short = 1E+308: minCross = 1E+308: penalty = 0#: score = 1#: miss = False
    innerLo = tLo + 0.5: If innerLo > 60# Then innerLo = 60#
    innerHi = tHi - 0.5: If innerHi < 60# Then innerHi = 60#
    innerAsp = tAsp - 0.1: If innerAsp < 1# Then innerAsp = 1#
    For j = 0 To k - 1
        a = CLng(rs(e, j + 2)): b = CLng(rs(e, (j + 1) Mod k + 2)): c = CLng(rs(e, (j + k - 1) Mod k + 2))
        ux = x(b) - x(a): uy = y(b) - y(a): vx = x(c) - x(a): vy = y(c) - y(a)
        cross = ux * vy - uy * vx
        If cross <= crossFloor Then ElementCost = 1E+100: Exit Function
        If cross < minCross Then minCross = cross
        lu = Sqr(ux * ux + uy * uy): lv = Sqr(vx * vx + vy * vy)
        angle = AcosDegrees((ux * vx + uy * vy) / (lu * lv))
        If angle < lo Then lo = angle
        If angle > hi Then hi = angle
        If lu < short Then short = lu
        If lu > longest Then longest = lu
        sumLen = sumLen + lu * lu: area = area + x(a) * y(b) - x(b) * y(a)
        If triangle Then
            cost = cost + Positive(innerLo - angle) ^ 2 / 400# + Positive(angle - innerHi) ^ 2 / 400#
            penalty = penalty + Positive(tLo - angle) ^ 2 / 400# + Positive(angle - tHi) ^ 2 / 400#
        Else
            cost = cost + 0.0001 * (angle - 90#) ^ 2 / 8100#
        End If
    Next j
    area = area / 2#: ratio = longest / short
    If triangle Then
        score = 4# * Sqr(3#) * area / sumLen
        If score < scoreFloor - 0.00000001 Then ElementCost = 1E+100: Exit Function
        cost = cost + Positive(ratio - innerAsp) ^ 2 / 25# + 0.00001 * (1# - score) ^ 2
        penalty = penalty + Positive(ratio - tAsp) ^ 2 / 25#
        miss = lo < tLo - 0.0000001 Or hi > tHi + 0.0000001 Or ratio > tAsp + 0.0000001
    Else
        If lo < qLo - 0.0000001 Or hi > qHi + 0.0000001 Or ratio > qAsp + 0.0000001 Or minCross / area < qJac - 0.0000001 Then ElementCost = 1E+100: Exit Function
    End If
    ElementCost = cost
End Function

Private Sub Measure(ByRef x() As Double, ByRef y() As Double, ByRef rs() As Variant, ByRef misses As Long, ByRef penalty As Double, ByRef minimum As Double)
    Dim e As Long, p As Double, q As Double, bad As Boolean, cost As Double
    misses = 0: penalty = 0#: minimum = 1#
    For e = 1 To UBound(rs, 1)
        If CStr(rs(e, 11)) = "T6" Then
            cost = ElementCost(x, y, rs, e, p, q, bad)
            If cost >= 1E+90 Then Err.Raise vbObjectError + 3985, , "T6局所改善の候補形状が不正です。"
            If bad Then misses = misses + 1
            penalty = penalty + p
            If q < minimum Then minimum = q
        End If
    Next e
End Sub

Public Sub MeshT6PatchSmooth(ByRef x() As Double, ByRef y() As Double, ByRef rs() As Variant, ByRef conditions As Variant, ByRef moved As Long)
    Dim n As Long, ne As Long, e As Long, j As Long, k As Long, a As Long, b As Long, m As Long, c As Long, pass As Long, limit As Long, ring As Long
    Dim adj() As Collection, neighbors() As Object, bounds() As Object, fixed() As Boolean, active() As Boolean, slideA() As Long, slideB() As Long
    Dim edges As Object, info As Variant, item As Variant, other As Variant, pair As Variant, edgeKey As Variant
    Dim ox() As Double, oy() As Double, bx() As Double, by() As Double, beforeBad As Long, bestBad As Long, bad As Long, changes As Long
    Dim beforePenalty As Double, bestPenalty As Double, penalty As Double, minimum As Double, q As Double, unused As Boolean
    Dim ux As Double, uy As Double, vx As Double, vy As Double, length As Double, ln As Double, mx As Double, my As Double, px As Double, py As Double, t As Double
    Dim cx(0 To 17) As Double, cy(0 To 17) As Double, v0x As Double, v0y As Double, cost As Double, baseCost As Double, bestCost As Double, best As Long, stepSize As Double
    If Not PatchScopeAllowed() Then
        moved = 0: MeshT6PatchLastDetail = "T6局所改善は節点・水理IDの条件保持のため省略": Exit Sub
    End If
    MeshQualityEnsureSettings
    qLo = FEMReadSetting("MESH_QUALITY_MIN_ANGLE_DEG", 45#): qHi = FEMReadSetting("MESH_QUALITY_MAX_ANGLE_DEG", 135#): qAsp = FEMReadSetting("MESH_QUALITY_MAX_ASPECT_RATIO", 5#)
    qJac = FEMReadSetting("MESH_Q8_MIN_JACOBIAN_AREA_RATIO", 0.1)
    tLo = FEMReadSetting("MESH_T6_MIN_ANGLE_DEG", 20#): tHi = FEMReadSetting("MESH_T6_MAX_ANGLE_DEG", 140#): tAsp = FEMReadSetting("MESH_T6_MAX_ASPECT_RATIO", 5#)
    limit = CLng(FEMReadSetting("MESH_T6_PATCH_ITERATIONS", 40#))
    If limit < 0 Or limit > 100 Then Err.Raise vbObjectError + 3980, , "T6局所改善の反復は0～100です。"
    moved = 0: If limit = 0 Then Exit Sub
    n = UBound(x): ne = UBound(rs, 1): ox = x: oy = y
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double, span As Double
    minX = x(1): maxX = x(1): minY = y(1): maxY = y(1)
    For a = 2 To n
        If x(a) < minX Then minX = x(a)
        If x(a) > maxX Then maxX = x(a)
        If y(a) < minY Then minY = y(a)
        If y(a) > maxY Then maxY = y(a)
    Next a
    span = maxX - minX: If maxY - minY > span Then span = maxY - minY
    crossFloor = span * span * 0.00000000000001: If crossFloor < 1E-24 Then crossFloor = 1E-24

    ReDim adj(1 To n): ReDim neighbors(1 To n): ReDim bounds(1 To n): ReDim fixed(1 To n): ReDim active(1 To n): ReDim slideA(1 To n): ReDim slideB(1 To n)
    Set edges = CreateObject("Scripting.Dictionary")
    For a = 1 To n
        Set adj(a) = New Collection: Set neighbors(a) = CreateObject("Scripting.Dictionary"): Set bounds(a) = CreateObject("Scripting.Dictionary")
        If a <= UBound(conditions, 1) Then
            For j = 4 To UBound(conditions, 2): If Val(CStr(conditions(a, j))) <> 0# Then fixed(a) = True
            Next j
        End If
    Next a
    For e = 1 To ne
        k = IIf(CStr(rs(e, 11)) = "T6", 3, 4)
        For j = 0 To k - 1
            a = rs(e, j + 2): b = rs(e, (j + 1) Mod k + 2): m = rs(e, IIf(k = 3 And j = 2, 9, j + 6))
            adj(a).Add e
            If k = 3 Then active(a) = True
            For c = 0 To k - 1
                If rs(e, c + 2) <> a Then neighbors(a)(CStr(rs(e, c + 2))) = CLng(rs(e, c + 2))
            Next c
            edgeKey = Key(a, b)
            If edges.Exists(edgeKey) Then
                info = edges(edgeKey)
                If info(0) <> 1 Or info(4) <> m Then Err.Raise vbObjectError + 3976, , "T6局所改善の共有辺が不整合です。"
                If info(3) <> rs(e, 10) Then bounds(a)(CStr(b)) = b: bounds(b)(CStr(a)) = a
                info(0) = 2: edges(edgeKey) = info
            Else
                edges.Add edgeKey, Array(1, a, b, rs(e, 10), m)
            End If
            If fixed(m) Or Abs(x(m) - (x(a) + x(b)) / 2#) + Abs(y(m) - (y(a) + y(b)) / 2#) > 0.000000001 Then fixed(a) = True: fixed(b) = True
        Next j
    Next e
    For Each edgeKey In edges.Keys
        info = edges(edgeKey): a = info(1): b = info(2)
        If info(0) = 1 Then bounds(a)(CStr(b)) = b: bounds(b)(CStr(a)) = a
    Next edgeKey
    ' Expand one element ring around the T6 vertices; all Q8 remain guarded.
    Dim expanded() As Boolean: expanded = active
    For a = 1 To n
        If active(a) Then
            For Each item In neighbors(a).Items: expanded(CLng(item)) = True: Next item
        End If
    Next a
    active = expanded
    For a = 1 To n
        If bounds(a).Count > 0 Then
            If bounds(a).Count <> 2 Then
                fixed(a) = True
            Else
                pair = bounds(a).Items: b = pair(0): c = pair(1)
                ux = x(b) - x(a): uy = y(b) - y(a): vx = x(c) - x(a): vy = y(c) - y(a)
                If Abs(ux * vy - uy * vx) > 0.0000000001 * Sqr((ux * ux + uy * uy) * (vx * vx + vy * vy)) Or ux * vx + uy * vy >= 0# Then
                    fixed(a) = True
                Else
                    slideA(a) = b: slideB(a) = c
                End If
            End If
        End If
    Next a
    scoreFloor = 0#: Measure x, y, rs, beforeBad, beforePenalty, minimum
    scoreFloor = minimum: bestBad = beforeBad: bestPenalty = beforePenalty: bx = x: by = y
    For pass = 1 To limit
        changes = 0
        For a = 1 To n
            If Not active(a) Or fixed(a) Or neighbors(a).Count = 0 Then GoTo NextNode
            length = 1E+308: mx = 0#: my = 0#
            For Each item In neighbors(a).Items
                b = CLng(item): ln = Sqr((x(b) - x(a)) ^ 2 + (y(b) - y(a)) ^ 2)
                If ln < length Then length = ln
                mx = mx + x(b): my = my + y(b)
            Next item
            cx(0) = x(a): cy(0) = y(a): cx(17) = mx / neighbors(a).Count: cy(17) = my / neighbors(a).Count
            For j = 1 To 16
                stepSize = 0.25: If j > 8 Then stepSize = 0.08
                t = ((j - 1) Mod 8) * 0.785398163397448
                cx(j) = x(a) + length * stepSize * Cos(t): cy(j) = y(a) + length * stepSize * Sin(t)
            Next j
            If slideA(a) > 0 Then
                b = slideA(a): c = slideB(a): ux = x(c) - x(b): uy = y(c) - y(b)
                For j = 0 To 17
                    t = ((cx(j) - x(b)) * ux + (cy(j) - y(b)) * uy) / (ux * ux + uy * uy)
                    If t < 0.001 Then t = 0.001
                    If t > 0.999 Then t = 0.999
                    cx(j) = x(b) + t * ux: cy(j) = y(b) + t * uy
                Next j
            End If
            v0x = x(a): v0y = y(a): bestCost = 1E+100: best = 0: baseCost = 1E+100
            For j = 0 To 17
                x(a) = cx(j): y(a) = cy(j): cost = 0#
                For Each item In adj(a)
                    q = ElementCost(x, y, rs, CLng(item), penalty, minimum, unused)
                    If q >= 1E+90 Then cost = 1E+100: Exit For
                    cost = cost + q
                Next item
                If j = 0 Then baseCost = cost
                If cost < bestCost Then bestCost = cost: best = j
            Next j
            x(a) = v0x: y(a) = v0y
            If bestCost < baseCost - 0.00000000001 Then x(a) = cx(best): y(a) = cy(best): changes = changes + 1
NextNode:
        Next a
        Measure x, y, rs, bad, penalty, minimum
        If bad < bestBad Or (bad = bestBad And penalty < bestPenalty - 0.000000000001) Then bestBad = bad: bestPenalty = penalty: bx = x: by = y
        If changes = 0 Then Exit For
    Next pass
    x = bx: y = by
    For Each edgeKey In edges.Keys
        info = edges(edgeKey): a = info(1): b = info(2): m = info(4)
        If x(a) <> ox(a) Or y(a) <> oy(a) Or x(b) <> ox(b) Or y(b) <> oy(b) Then x(m) = (x(a) + x(b)) / 2#: y(m) = (y(a) + y(b)) / 2#
    Next edgeKey
    For a = 1 To n: If x(a) <> ox(a) Or y(a) <> oy(a) Then moved = moved + 1
    Next a
    MeshT6PatchLastDetail = "T6局所節点改善：目標外 " & CStr(beforeBad) & "→" & CStr(bestBad) & " / 移動節点 " & CStr(moved)
End Sub

Public Sub MeshT6PatchCurrentSheet()
    Dim wn As Worksheet, we As Worksheet, ns As Variant, rs() As Variant, x() As Double, y() As Double, cn() As Long, coords() As Variant
    Dim n As Long, ne As Long, j As Long, e As Long, moved As Long
    If AnalysisRunning Then Err.Raise 5, , "解析中です。"
    If FEMReadSetting("MESH_T6_PATCH_ENABLE", 1#) < 0.5 Or Not FEMMeshSheetHasT6() Then Exit Sub
    If Not PatchScopeAllowed() Then
        MeshT6PatchLastDetail = "T6局所改善は載荷・接合・水理IDの条件保持のため省略"
        FEMWriteTextSetting "MESH_GEN_LAST_T6_PATCH", MeshT6PatchLastDetail
        Exit Sub
    End If
    Set wn = ThisWorkbook.Worksheets("節点データ"): Set we = ThisWorkbook.Worksheets("要素データ")
    n = wn.Cells(wn.Rows.Count, 1).End(xlUp).Row - 1: ne = we.Cells(we.Rows.Count, 1).End(xlUp).Row - 1
    ns = wn.Range(wn.Cells(2, 1), wn.Cells(n + 1, 14)).Value2: rs = we.Range(we.Cells(2, 1), we.Cells(ne + 1, 11)).Value2
    ReDim x(1 To n): ReDim y(1 To n): ReDim cn(1 To ne, 1 To 8): ReDim coords(1 To n, 1 To 2)
    For j = 1 To n: x(j) = ns(j, 2): y(j) = ns(j, 3): Next j
    MeshT6PatchSmooth x, y, rs, ns, moved
    If FEMReadSetting("MESH_T6_PATCH_REBUILD", 1#) >= 0.5 And PatchScopeAllowed() Then MeshT6RebuildPatch x, y, rs, ns
    n = UBound(x): ne = UBound(rs, 1)
    ReDim cn(1 To ne, 1 To 8)
    For e = 1 To ne: For j = 1 To 8: cn(e, j) = rs(e, j + 1): Next j: Next e
    P5ValidateMeshQuality n, ne, cn, x, y
    ReDim coords(1 To n, 1 To 14)
    For j = 1 To n
        coords(j, 1) = j: coords(j, 2) = x(j): coords(j, 3) = y(j)
        If j <= UBound(ns, 1) Then
            For e = 4 To 14: coords(j, e) = ns(j, e): Next e
        End If
    Next j
    wn.Range(wn.Cells(2, 1), wn.Cells(wn.Rows.Count, 14)).ClearContents
    wn.Range(wn.Cells(2, 1), wn.Cells(n + 1, 14)).Value2 = coords
    we.Range(we.Cells(2, 1), we.Cells(we.Rows.Count, 11)).ClearContents
    we.Range(we.Cells(2, 1), we.Cells(ne + 1, 11)).Value2 = rs
    FEMInvalidateRemeshedModel
    FEMWriteTextSetting "MESH_GEN_LAST_T6_PATCH", MeshT6PatchLastDetail
    P5EvaluateCurrentSheet "T6局所節点改善後"
End Sub

' Rebuild a small cavity rather than endlessly bisecting an acute input tip.
' A free interior T6 vertex may be removed after converting at most two
' adjacent Q8. Keep the exact outer/interface edges and require fewer misses.
Private Sub MeshT6RebuildPatch(ByRef x() As Double, ByRef y() As Double, ByRef rs() As Variant, ByRef conditions As Variant)
    Dim edges As Object, adj() As Collection, boundary() As Boolean, candidates As Object, neighbors As Object
    Dim info As Variant, edgeKey As Variant, item As Variant, target As Variant, vertex As Variant
    Dim n As Long, ne As Long, e As Long, j As Long, k As Long, a As Long, b As Long, v As Long, q8 As Long, material As Long
    Dim beforeBad As Long, bad As Long, converted As Long, outCount As Long, penalty As Double, minimum As Double, p As Double, q As Double, miss As Boolean, cost As Double
    Dim mask() As Boolean, cx() As Double, cy() As Double, cr() As Variant, output As Variant, detail As String, safe As Boolean
    Dim nx() As Double, ny() As Double, nr() As Variant, nc As Variant
    ' Compact node/element IDs only when there are no external ID-based tables.
    If ThisWorkbook.Worksheets("接合").Cells(Rows.Count, 1).End(xlUp).Row > 1 Then
        MeshT6PatchLastDetail = MeshT6PatchLastDetail & " / 局所再分割は載荷・接合IDの保持のため省略": Exit Sub
    End If
    Dim loadRow As Long, selector As String, mode As String
    With ThisWorkbook.Worksheets("載荷")
        For loadRow = 2 To .Cells(.Rows.Count, 1).End(xlUp).Row
            If Len(CStr(.Cells(loadRow, 1).Value2)) > 0 Then
                selector = UCase$(Trim$(CStr(.Cells(loadRow, 2).Value2))): mode = UCase$(Trim$(CStr(.Cells(loadRow, 4).Value2)))
                If (selector <> "TOP" And selector <> "BOTTOM" And selector <> "LEFT" And selector <> "RIGHT") Or (mode <> "DISP" And mode <> "FIX") Then
                    MeshT6PatchLastDetail = MeshT6PatchLastDetail & " / 局所再分割は載荷IDの保持のため省略": Exit Sub
                End If
            End If
        Next loadRow
    End With
    If UCase$(Trim$(CStr(ThisWorkbook.Worksheets("水条件").Range("B3").Value2))) <> "NONE" Then
        MeshT6PatchLastDetail = MeshT6PatchLastDetail & " / 局所再分割は水理IDの保持のため省略": Exit Sub
    End If
    n = UBound(x): ne = UBound(rs, 1): ReDim adj(1 To n): ReDim boundary(1 To n)
    Set edges = CreateObject("Scripting.Dictionary"): Set candidates = CreateObject("Scripting.Dictionary")
    For a = 1 To n: Set adj(a) = New Collection: Next a
    Measure x, y, rs, beforeBad, penalty, minimum
    If beforeBad = 0 Then Exit Sub
    For e = 1 To ne
        k = IIf(CStr(rs(e, 11)) = "T6", 3, 4)
        cost = ElementCost(x, y, rs, e, p, q, miss)
        For j = 0 To k - 1
            a = rs(e, j + 2): b = rs(e, (j + 1) Mod k + 2): adj(a).Add e
            If k = 3 And miss Then candidates(CStr(a)) = a
            edgeKey = Key(a, b)
            If edges.Exists(edgeKey) Then
                info = edges(edgeKey)
                If info(1) <> rs(e, 10) Then boundary(a) = True: boundary(b) = True
                info(0) = 2: edges(edgeKey) = info
            Else
                edges.Add edgeKey, Array(1, rs(e, 10), a, b)
            End If
        Next j
    Next e
    For Each edgeKey In edges.Keys
        info = edges(edgeKey): If info(0) = 1 Then boundary(info(2)) = True: boundary(info(3)) = True
    Next edgeKey
    For Each vertex In candidates.Items
        v = CLng(vertex): If boundary(v) Then GoTo NextVertex
        safe = True
        If v <= UBound(conditions, 1) Then
            For j = 4 To UBound(conditions, 2): If Val(CStr(conditions(v, j))) <> 0# Then safe = False
            Next j
        End If
        If Not safe Then GoTo NextVertex
        ReDim mask(1 To ne): q8 = 0: material = 0
        For Each item In adj(v)
            e = CLng(item)
            If material = 0 Then material = rs(e, 10)
            If material <> rs(e, 10) Then safe = False
            If CStr(rs(e, 11)) <> "T6" Then mask(e) = True: q8 = q8 + 1
        Next item
        If Not safe Or q8 > 2 Then GoTo NextVertex
        cx = x: cy = y: cr = rs
        If q8 > 0 Then
            output = MeshAdaptConvertBadQ8ToT6(cx, cy, cr, mask, converted, outCount, detail): cr = output
        End If
        Set neighbors = CreateObject("Scripting.Dictionary")
        For e = 1 To UBound(cr, 1)
            If CStr(cr(e, 11)) = "T6" Then
                If cr(e, 2) = v Or cr(e, 3) = v Or cr(e, 4) = v Then
                    For j = 2 To 4: If cr(e, j) <> v Then neighbors(CStr(cr(e, j))) = CLng(cr(e, j))
                    Next j
                End If
            End If
        Next e
        For Each target In neighbors.Items
            If CollapseCandidate(cx, cy, cr, conditions, v, CLng(target), nx, ny, nr, nc) Then
                Measure nx, ny, nr, bad, penalty, minimum
                If bad < beforeBad And minimum >= scoreFloor - 0.00000001 Then
                    x = nx: y = ny: rs = nr: conditions = nc
                    MeshT6PatchLastDetail = MeshT6PatchLastDetail & " / 局所再分割：Q8→T6 " & CStr(q8) & "個、自由内点1個を除去、目標外 " & CStr(beforeBad) & "→" & CStr(bad)
                    Exit Sub
                End If
            End If
        Next target
NextVertex:
    Next vertex
    MeshT6PatchLastDetail = MeshT6PatchLastDetail & " / 局所再分割：改善候補なし"
End Sub

Private Function CollapseCandidate(ByRef x() As Double, ByRef y() As Double, ByRef source() As Variant, ByRef conditions As Variant, ByVal removed As Long, ByVal target As Long, ByRef nx() As Double, ByRef ny() As Double, ByRef nr() As Variant, ByRef nc As Variant) As Boolean
    Dim oldEdges As Object, newEdges As Object, oldArea As Object, newArea As Object, keyText As String, matKey As Variant
    Dim rows() As Variant, tx() As Double, ty() As Double, used() As Boolean, map() As Long
    Dim e As Long, j As Long, k As Long, a As Long, b As Long, m As Long, count As Long, n As Long, nn As Long, rawCount As Long, r As Long, dropped As Long
    Dim p As Double, q As Double, bad As Boolean, cost As Double, area As Double, valid As Boolean, duplicate As Boolean, h As Long
    On Error GoTo Reject
    Set oldEdges = CreateObject("Scripting.Dictionary"): Set newEdges = CreateObject("Scripting.Dictionary")
    Set oldArea = CreateObject("Scripting.Dictionary"): Set newArea = CreateObject("Scripting.Dictionary")
    n = UBound(x): tx = x: ty = y: ReDim rows(1 To UBound(source, 1), 1 To 11)
    For e = 1 To UBound(source, 1)
        k = IIf(CStr(source(e, 11)) = "T6", 3, 4): area = 0#
        For j = 0 To k - 1
            a = source(e, j + 2): b = source(e, (j + 1) Mod k + 2): m = source(e, IIf(k = 3 And j = 2, 9, j + 6))
            oldEdges(Key(a, b)) = m: area = area + x(a) * y(b) - x(b) * y(a)
        Next j
        matKey = CStr(source(e, 10)): If Not oldArea.Exists(matKey) Then oldArea(matKey) = 0#
        oldArea(matKey) = oldArea(matKey) + area / 2#
        duplicate = False
        For j = 2 To k + 1
            For h = 2 To j - 1
                a = source(e, j): b = source(e, h): If a = removed Then a = target
                If b = removed Then b = target
                If a = b Then duplicate = True
            Next h
        Next j
        If duplicate Then
            If k <> 3 Then GoTo Reject
            dropped = dropped + 1
        Else
            count = count + 1
            For j = 1 To 11: rows(count, j) = source(e, j): Next j
            For j = 2 To k + 1: If rows(count, j) = removed Then rows(count, j) = target
            Next j
            If k = 3 Then rows(count, 5) = rows(count, 4): rows(count, 8) = rows(count, 4)
            cost = ElementCost(tx, ty, rows, count, p, q, bad)
            If cost >= 1E+90 Then GoTo Reject
            area = 0#
            For j = 0 To k - 1
                a = rows(count, j + 2): b = rows(count, (j + 1) Mod k + 2)
                area = area + tx(a) * ty(b) - tx(b) * ty(a)
                keyText = Key(a, b)
                If newEdges.Exists(keyText) Then
                    m = newEdges(keyText)
                ElseIf oldEdges.Exists(keyText) Then
                    m = oldEdges(keyText): newEdges(keyText) = m
                Else
                    ' Old edges may occur later in the input table. Build all
                    ' midpoint references in a second pass below instead.
                    m = 0
                End If
            Next j
            If Not newArea.Exists(matKey) Then newArea(matKey) = 0#
            newArea(matKey) = newArea(matKey) + area / 2#
        End If
    Next e
    If dropped <> 2 Then GoTo Reject
    For Each matKey In oldArea.Keys
        If Not newArea.Exists(matKey) Then GoTo Reject
        If Abs(oldArea(matKey) - newArea(matKey)) > 0.000000001 * (1# + Abs(oldArea(matKey))) Then GoTo Reject
    Next matKey
    Set newEdges = CreateObject("Scripting.Dictionary")
    For e = 1 To count
        k = IIf(CStr(rows(e, 11)) = "T6", 3, 4)
        For j = 0 To k - 1
            a = rows(e, j + 2): b = rows(e, (j + 1) Mod k + 2): keyText = Key(a, b)
            If newEdges.Exists(keyText) Then
                m = newEdges(keyText)
            Else
                If oldEdges.Exists(keyText) Then
                    m = oldEdges(keyText)
                Else
                    n = n + 1: If n > 60000 Then GoTo Reject
                    ReDim Preserve tx(1 To n): ReDim Preserve ty(1 To n)
                    tx(n) = (tx(a) + tx(b)) / 2#: ty(n) = (ty(a) + ty(b)) / 2#: m = n
                End If
                newEdges(keyText) = m
            End If
            rows(e, IIf(k = 3 And j = 2, 9, j + 6)) = m
        Next j
    Next e
    ReDim used(1 To n): ReDim map(1 To n)
    For e = 1 To count: For j = 2 To 9: used(CLng(rows(e, j))) = True: Next j: Next e
    For a = 1 To n
        If used(a) Then
            nn = nn + 1: map(a) = nn
        ElseIf a <= UBound(conditions, 1) Then
            For j = 4 To UBound(conditions, 2): If Val(CStr(conditions(a, j))) <> 0# Then GoTo Reject
            Next j
        End If
    Next a
    ReDim nx(1 To nn): ReDim ny(1 To nn): ReDim nr(1 To count, 1 To 11): ReDim nc(1 To nn, 1 To 14)
    For a = 1 To n
        If map(a) > 0 Then
            b = map(a): nx(b) = tx(a): ny(b) = ty(a): nc(b, 1) = b: nc(b, 2) = tx(a): nc(b, 3) = ty(a)
            If a <= UBound(conditions, 1) Then
                For j = 4 To UBound(conditions, 2): nc(b, j) = conditions(a, j): Next j
            End If
        End If
    Next a
    For e = 1 To count
        nr(e, 1) = e: nr(e, 10) = rows(e, 10): nr(e, 11) = rows(e, 11)
        For j = 2 To 9: nr(e, j) = map(CLng(rows(e, j))): Next j
    Next e
    CollapseCandidate = True
Reject:
    Err.Clear
End Function

' Conservative scope for moving original nodes or compacting IDs. Geometric
' side DISP/FIX loads retain the same boundary nodes. Other selectors, force
' anchors, hydraulic plans and element-ID sets require separate remapping.
Private Function PatchScopeAllowed() As Boolean
    Dim ws As Worksheet, r As Long, selector As String, mode As String
    If UCase$(Trim$(CStr(ThisWorkbook.Worksheets("水条件").Range("B3").Value2))) <> "NONE" Then Exit Function
    If ThisWorkbook.Worksheets("接合").Cells(ThisWorkbook.Worksheets("接合").Rows.Count, 1).End(xlUp).Row > 1 Then Exit Function
    Set ws = ThisWorkbook.Worksheets("載荷")
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Len(CStr(ws.Cells(r, 1).Value2)) > 0 Then
            selector = UCase$(Trim$(CStr(ws.Cells(r, 2).Value2))): mode = UCase$(Trim$(CStr(ws.Cells(r, 4).Value2)))
            If (selector <> "TOP" And selector <> "BOTTOM" And selector <> "LEFT" And selector <> "RIGHT") Or (mode <> "DISP" And mode <> "FIX") Then Exit Function
        End If
    Next r
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("要素セット"): On Error GoTo 0
    If Not ws Is Nothing Then
        If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then Exit Function
    End If
    PatchScopeAllowed = True
End Function
