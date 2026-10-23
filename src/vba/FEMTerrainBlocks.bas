Option Explicit
' Non-Delaunay Q8 blocks. No worksheet is changed until all checks pass.
Private tx() As Double, ty() As Double, tc() As Long, tm() As Long
Private tn As Long, te As Long, nodeCap As Long, elemCap As Long
Private vertexMap As Object, midpointMap As Object
Private loops() As Variant, loopN As Long, outerN As Long, holeN As Long
Private stepSize As Double, modelSize As Double
Private radialUsed As Boolean
Private blocksUsed As Boolean

Private Sub TBFail(ByVal reason As String)
    Err.Raise vbObjectError + 3860, "FEMTerrainBlocks", "地形格子を適用できません：" & reason
End Sub

Public Sub MeshPrepareTerrainBlocks(ByRef outX() As Double, ByRef outY() As Double, ByRef elements() As Variant)
    Dim ws As Worksheet, counts() As Long, holes() As Long, ids() As Long, gx() As Double, gy() As Double
    Dim coords As Variant, record As Variant, points() As Double, i As Long, j As Long, row As Long, n As Long, k As Long
    Dim orthogonal As Boolean, method As String, materialRule As String, lowX As Double, highX As Double, lowY As Double, highY As Double
    Set ws = ThisWorkbook.Worksheets("図形定義")
    outerN = CLng(Val(ws.Cells(3, 4).Value2)): holeN = CLng(Val(ws.Cells(24, 4).Value2))
    If outerN < 1 Or outerN > 21 Or holeN < 0 Or holeN > 100 Then TBFail "外側境界1～21、内側境界0～100の範囲で入力してください。"
    loopN = outerN + holeN: n = CLng(Val(ws.Cells(1, 12).Value2))
    If n < 3 Or n > 60000 Then TBFail "座標数が不正です。形状入力から作成してください。"
    coords = ws.Range(ws.Cells(2, 2), ws.Cells(n + 1, 3)).Value2
    ReDim gx(1 To n): ReDim gy(1 To n): ReDim loops(1 To loopN)
    ReDim counts(1 To outerN): ReDim holes(1 To IIf(holeN > 0, holeN, 1)): ReDim ids(1 To loopN, 1 To 2000)
    For i = 1 To n
        gx(i) = CDbl(coords(i, 1)): gy(i) = CDbl(coords(i, 2))
        If i = 1 Then lowX = gx(i): highX = gx(i): lowY = gy(i): highY = gy(i)
        If gx(i) < lowX Then lowX = gx(i)
        If gx(i) > highX Then highX = gx(i)
        If gy(i) < lowY Then lowY = gy(i)
        If gy(i) > highY Then highY = gy(i)
    Next i
    modelSize = Application.Max(highX - lowX, highY - lowY)
    If modelSize <= 0# Then TBFail "領域の大きさが0です。"
    stepSize = 1E+100: orthogonal = True
    For i = 1 To loopN
        If i <= outerN Then row = i + 2 Else row = i - outerN + 23
        k = CLng(Val(ws.Cells(row, 5).Value2))
        If k < 3 Or k > 1999 Then TBFail "境界点は3～1999点で入力してください。"
        If i <= outerN Then counts(i) = k Else holes(i - outerN) = k
        If Not IsNumeric(ws.Cells(row, 6).Value2) Then TBFail "格子間隔が数値ではありません。"
        If CDbl(ws.Cells(row, 6).Value2) <= 0# Then TBFail "格子間隔には正数が必要です。"
        stepSize = Application.Min(stepSize, CDbl(ws.Cells(row, 6).Value2) / 2#)
        record = ws.Range(ws.Cells(row, 7), ws.Cells(row, 6 + k)).Value2
        ReDim points(1 To k, 1 To 2)
        For j = 1 To k
            ids(i, j) = CLng(record(1, j))
            If ids(i, j) < 1 Or ids(i, j) > n Then TBFail "境界の節点番号が座標表にありません。"
            points(j, 1) = gx(ids(i, j)): points(j, 2) = gy(ids(i, j))
        Next j
        For j = 1 To k
            If points(j, 1) <> points(j Mod k + 1, 1) And points(j, 2) <> points(j Mod k + 1, 2) Then orthogonal = False
        Next j
        loops(i) = points
    Next i
    P5ValidateBoundaryGeometry outerN, holeN, counts, holes, ids, gx, gy, n
    If modelSize / stepSize > 30000# Then TBFail "間隔が小さすぎます。節点・要素上限60000以内となるよう大きくしてください。"
    tn = 0: te = 0: nodeCap = 256: elemCap = 128
    ReDim tx(1 To nodeCap): ReDim ty(1 To nodeCap): ReDim tc(1 To elemCap, 1 To 8): ReDim tm(1 To elemCap)
    Set vertexMap = CreateObject("Scripting.Dictionary"): Set midpointMap = CreateObject("Scripting.Dictionary")
    radialUsed = False: blocksUsed = False
    If orthogonal Then
        Call TBOrthogonal
        method = "境界整合直交格子"
    ElseIf holeN = 0 And TBAllQuads() Then
        blocksUsed = True
        Call TBQuadBlocks
        method = "境界整合四辺形ブロック格子"
    Else
        radialUsed = True
        Call TBTunnel
        method = "トンネル放射ブロック格子"
    End If
    If te = 0 Then TBFail "有効な要素がありません。"
    ReDim Preserve tx(1 To tn): ReDim Preserve ty(1 To tn)
    Call TBValidate
    ReDim outX(1 To tn): ReDim outY(1 To tn): ReDim elements(1 To te, 1 To 10)
    For i = 1 To tn: outX(i) = tx(i): outY(i) = ty(i): Next i
    For i = 1 To te
        elements(i, 1) = i: elements(i, 10) = tm(i)
        For j = 1 To 8: elements(i, j + 1) = tc(i, j): Next j
    Next i
    If outerN = 1 Then
        materialRule = "材料=" & CStr(MeshReadMaterialNumber()) & "（メッシュ設定）"
    ElseIf blocksUsed Then
        materialRule = "材料=ブロック入力順（重なり不可）"
    Else
        materialRule = "材料=外側境界の入力順（重なりは先を優先）"
    End If
    MeshAdaptLastSummary = method & "：デローニ分割なし。" & materialRule & vbCrLf & _
        "全境界の最小入力間隔/2=" & CStr(stepSize) & " / 節点 " & CStr(tn) & " / 要素 " & CStr(te) & vbCrLf & P5QualitySummary()
    Set vertexMap = Nothing: Set midpointMap = Nothing
End Sub

Private Function TBInside(ByVal x As Double, ByVal y As Double, ByVal index As Long) As Boolean
    Dim p As Variant, i As Long, j As Long
    p = loops(index): j = UBound(p, 1)
    For i = 1 To UBound(p, 1)
        If (p(i, 2) > y) <> (p(j, 2) > y) Then
            If x < p(i, 1) + (p(j, 1) - p(i, 1)) * (y - p(i, 2)) / (p(j, 2) - p(i, 2)) Then TBInside = Not TBInside
        End If
        j = i
    Next i
End Function

Private Function TBAllQuads() As Boolean
    Dim i As Long, p As Variant
    For i = 1 To outerN
        p = loops(i)
        If UBound(p, 1) <> 4 Then Exit Function
    Next i
    TBAllQuads = True
End Function

Private Function TBRoot(ByRef parent() As Long, ByVal i As Long) As Long
    Do While parent(i) <> i: i = parent(i): Loop
    TBRoot = i
End Function

Private Function TBEdgeKey(ByVal a As Long, ByVal b As Long) As String
    If a < b Then TBEdgeKey = CStr(a) & ":" & CStr(b) Else TBEdgeKey = CStr(b) & ":" & CStr(a)
End Function

Private Function TBEdgePoint(ByVal a As Long, ByVal b As Long, ByVal slot As Long, ByVal count As Long) As Long
    Dim t As Double, k As Long
    If slot = 0 Then TBEdgePoint = a: Exit Function
    If slot = count Then TBEdgePoint = b: Exit Function
    If a > b Then k = a: a = b: b = k: slot = count - slot
    t = CDbl(slot) / count
    TBEdgePoint = TBVertex("E" & TBEdgeKey(a, b) & ":" & CStr(slot), tx(a) * (1# - t) + tx(b) * t, ty(a) * (1# - t) + ty(b) * t)
End Function

Private Function TBInteriorOnEdge(ByVal x As Double, ByVal y As Double, ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double) As Boolean
    Dim vx As Double, vy As Double, length As Double, t As Double
    vx = bx - ax: vy = by - ay: length = Sqr(vx * vx + vy * vy)
    If length <= 0# Then Exit Function
    If Abs(TBCross(x - ax, y - ay, vx, vy)) > modelSize * 0.0000000001 * length Then Exit Function
    t = ((x - ax) * vx + (y - ay) * vy) / (length * length)
    TBInteriorOnEdge = t > 0.0000000001 And t < 0.9999999999
End Function

Private Sub TBQuadBlocks()
    Dim i As Long, j As Long, k As Long, h As Long, other As Long, nEdge As Long, id As Long, a As Long, b As Long
    Dim p As Variant, q As Variant, z As Double, area As Double, cx As Double, cy As Double, length As Double
    Dim cids() As Long, edgeIds() As Long, parent() As Long, wanted() As Long, groups() As Long, edges As Object
    Dim key As String, nx As Long, ny As Long, grid() As Long, row As Long, col As Long, u As Double, v As Double, mat As Long
    Dim ca As Double, cb As Double, cc As Double, cd As Double
    Dim nodeNeed As Double, elementNeed As Double, band As Long, memory As Double
    ' Blocks must meet on complete sides: reject overlapping regions, crossing sides and T-junctions.
    For i = 1 To outerN
        p = loops(i): area = 0#
        For j = 1 To 4: area = area + TBCross(p(j, 1) - p(1, 1), p(j, 2) - p(1, 2), p(j Mod 4 + 1, 1) - p(1, 1), p(j Mod 4 + 1, 2) - p(1, 2)): Next j
        If area < 0# Then
            z = p(2, 1): p(2, 1) = p(4, 1): p(4, 1) = z
            z = p(2, 2): p(2, 2) = p(4, 2): p(4, 2) = z
        End If
        For j = 1 To 4
            k = j Mod 4 + 1: h = k Mod 4 + 1
            If TBCross(p(k, 1) - p(j, 1), p(k, 2) - p(j, 2), p(h, 1) - p(k, 1), p(h, 2) - p(k, 2)) <= 0# Then TBFail "四辺形ブロックは凸形状で入力してください。境界=" & CStr(i)
        Next j
        loops(i) = p
    Next i
    For i = 1 To outerN
        p = loops(i): cx = 0#: cy = 0#
        For j = 1 To 4: cx = cx + p(j, 1) / 4#: cy = cy + p(j, 2) / 4#: Next j
        For other = 1 To outerN
            If other <> i Then
                q = loops(other)
                If TBInside(cx, cy, other) Then TBFail "四辺形ブロックが重なっています。土・壁体・底盤を重ならない領域に分割してください。"
                For j = 1 To 4
                    k = j Mod 4 + 1
                    For h = 1 To 4
                        a = h Mod 4 + 1
                        If TBInteriorOnEdge(p(j, 1), p(j, 2), q(h, 1), q(h, 2), q(a, 1), q(a, 2)) Then TBFail "ブロック境界にT字接続があります。隣接ブロックも同じ端点で分割してください。"
                        ca = TBCross(p(k, 1) - p(j, 1), p(k, 2) - p(j, 2), q(h, 1) - p(j, 1), q(h, 2) - p(j, 2))
                        cb = TBCross(p(k, 1) - p(j, 1), p(k, 2) - p(j, 2), q(a, 1) - p(j, 1), q(a, 2) - p(j, 2))
                        cc = TBCross(q(a, 1) - q(h, 1), q(a, 2) - q(h, 2), p(j, 1) - q(h, 1), p(j, 2) - q(h, 2))
                        cd = TBCross(q(a, 1) - q(h, 1), q(a, 2) - q(h, 2), p(k, 1) - q(h, 1), p(k, 2) - q(h, 2))
                        If ca * cb < 0# And cc * cd < 0# Then TBFail "ブロックの辺が交差しています。交点で領域を分割してください。"
                    Next h
                Next j
            End If
        Next other
    Next i
    ReDim cids(1 To outerN, 1 To 4): ReDim edgeIds(1 To outerN, 1 To 4)
    ReDim parent(1 To 4 * outerN): ReDim wanted(1 To 4 * outerN): ReDim groups(1 To 4 * outerN)
    Set edges = CreateObject("Scripting.Dictionary")
    For i = 1 To outerN
        p = loops(i)
        For j = 1 To 4: cids(i, j) = TBVertex("P" & CStr(p(j, 1)) & ":" & CStr(p(j, 2)), p(j, 1), p(j, 2)): Next j
        For j = 1 To 4
            a = cids(i, j): b = cids(i, j Mod 4 + 1): key = TBEdgeKey(a, b)
            If Not edges.Exists(key) Then
                nEdge = nEdge + 1: edges.Add key, nEdge: parent(nEdge) = nEdge
                length = Sqr((tx(a) - tx(b)) ^ 2 + (ty(a) - ty(b)) ^ 2): wanted(nEdge) = CLng(-Int(-length / stepSize))
            End If
            edgeIds(i, j) = CLng(edges(key))
        Next j
    Next i
    For i = 1 To outerN
        parent(TBRoot(parent, edgeIds(i, 1))) = TBRoot(parent, edgeIds(i, 3))
        parent(TBRoot(parent, edgeIds(i, 2))) = TBRoot(parent, edgeIds(i, 4))
    Next i
    For i = 1 To nEdge: id = TBRoot(parent, i): groups(id) = Application.Max(groups(id), wanted(i)): Next i
    nodeNeed = tn
    For i = 1 To nEdge: nodeNeed = nodeNeed + 2# * groups(TBRoot(parent, i)) - 1#: Next i
    For i = 1 To outerN
        nx = groups(TBRoot(parent, edgeIds(i, 1))): ny = groups(TBRoot(parent, edgeIds(i, 2)))
        elementNeed = elementNeed + CDbl(nx) * ny
        nodeNeed = nodeNeed + 3# * nx * ny - 2# * nx - 2# * ny + 1#
    Next i
    If nodeNeed > 60000# Or elementNeed > 60000# Then TBFail "ブロック全体が節点・要素上限60000を超えます。間隔を大きくしてください。"
    band = CLng(8# * Sqr(nodeNeed)): If band > 2 * nodeNeed - 1 Then band = CLng(2 * nodeNeed - 1)
    memory = 4# * 8# * (2# * nodeNeed) * (band + 1#) + elementNeed * 16384#
    If memory > FEMReadSetting("SOLVER_MEMORY_LIMIT_MB", 512#) * 1048576# Then TBFail "ブロック全体の概算解析メモリが上限を超えます。間隔を大きくしてください。"
    For i = 1 To outerN
        nx = groups(TBRoot(parent, edgeIds(i, 1))): ny = groups(TBRoot(parent, edgeIds(i, 2)))
        If CDbl(nx) * ny > 19000# Or CDbl(nx + 1) * (ny + 1) > 60000# Then TBFail "ブロックの分割数が多すぎます。間隔を大きくしてください。"
        ReDim grid(0 To ny, 0 To nx): p = loops(i)
        For row = 0 To ny
            v = CDbl(row) / ny
            For col = 0 To nx
                u = CDbl(col) / nx
                If row = 0 Then
                    grid(row, col) = TBEdgePoint(cids(i, 1), cids(i, 2), col, nx)
                ElseIf row = ny Then
                    grid(row, col) = TBEdgePoint(cids(i, 4), cids(i, 3), col, nx)
                ElseIf col = 0 Then
                    grid(row, col) = TBEdgePoint(cids(i, 1), cids(i, 4), row, ny)
                ElseIf col = nx Then
                    grid(row, col) = TBEdgePoint(cids(i, 2), cids(i, 3), row, ny)
                Else
                    grid(row, col) = TBNode((1# - u) * (1# - v) * p(1, 1) + u * (1# - v) * p(2, 1) + u * v * p(3, 1) + (1# - u) * v * p(4, 1), _
                        (1# - u) * (1# - v) * p(1, 2) + u * (1# - v) * p(2, 2) + u * v * p(3, 2) + (1# - u) * v * p(4, 2))
                End If
            Next col
        Next row
        If outerN = 1 Then mat = MeshReadMaterialNumber() Else mat = i
        For row = 0 To ny - 1
            For col = 0 To nx - 1: TBQuad grid(row, col), grid(row, col + 1), grid(row + 1, col + 1), grid(row + 1, col), mat: Next col
        Next row
    Next i
End Sub

Private Function TBMaterial(ByVal x As Double, ByVal y As Double) As Long
    Dim i As Long
    For i = outerN + 1 To loopN
        If TBInside(x, y, i) Then Exit Function
    Next i
    For i = 1 To outerN
        If TBInside(x, y, i) Then
            If outerN = 1 Then TBMaterial = MeshReadMaterialNumber() Else TBMaterial = i
            Exit Function
        End If
    Next i
End Function

Private Sub TBSort(ByRef values() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Double, z As Double
    i = lo: j = hi: pivot = values((lo + hi) \ 2)
    Do While i <= j
        Do While values(i) < pivot: i = i + 1: Loop
        Do While values(j) > pivot: j = j - 1: Loop
        If i <= j Then z = values(i): values(i) = values(j): values(j) = z: i = i + 1: j = j - 1
    Loop
    If lo < j Then TBSort values, lo, j
    If i < hi Then TBSort values, i, hi
End Sub

Private Sub TBCuts(ByVal axis As Long, ByRef cuts() As Double, ByRef count As Long)
    Dim raw() As Double, p As Variant, i As Long, j As Long, k As Long, m As Long, num As Long, key As Variant, keys As Object
    Set keys = CreateObject("Scripting.Dictionary")
    For i = 1 To loopN
        p = loops(i)
        For j = 1 To UBound(p, 1): keys(p(j, axis)) = True: Next j
    Next i
    ReDim raw(1 To keys.Count): i = 0
    For Each key In keys.Keys: i = i + 1: raw(i) = CDbl(key): Next key
    TBSort raw, 1, UBound(raw)
    count = 1
    For i = 1 To UBound(raw) - 1
        count = count + CLng(-Int(-(raw(i + 1) - raw(i)) / stepSize))
        If count > 30001 Then TBFail "格子線が多すぎます。間隔を大きくしてください。"
    Next i
    ReDim cuts(1 To count): cuts(1) = raw(1): k = 1
    For i = 1 To UBound(raw) - 1
        num = CLng(-Int(-(raw(i + 1) - raw(i)) / stepSize))
        For m = 1 To num: k = k + 1: cuts(k) = raw(i) + (raw(i + 1) - raw(i)) * m / num: Next m
    Next i
End Sub

Private Sub TBOrthogonal()
    Dim xs() As Double, ys() As Double, nx As Long, ny As Long, i As Long, j As Long, mat As Long
    Dim a As Long, b As Long, c As Long, d As Long
    TBCuts 1, xs, nx
    TBCuts 2, ys, ny
    If CDbl(nx - 1) * (ny - 1) > 2000000# Then TBFail "走査格子数が200万を超えます。間隔を大きくしてください。"
    For j = 1 To ny - 1
        For i = 1 To nx - 1
            mat = TBMaterial((xs(i) + xs(i + 1)) / 2#, (ys(j) + ys(j + 1)) / 2#)
            If mat > 0 Then
                a = TBVertex("C" & CStr(i) & ":" & CStr(j), xs(i), ys(j))
                b = TBVertex("C" & CStr(i + 1) & ":" & CStr(j), xs(i + 1), ys(j))
                c = TBVertex("C" & CStr(i + 1) & ":" & CStr(j + 1), xs(i + 1), ys(j + 1))
                d = TBVertex("C" & CStr(i) & ":" & CStr(j + 1), xs(i), ys(j + 1))
                TBQuad a, b, c, d, mat
            End If
        Next i
    Next j
End Sub

Private Function TBNode(ByVal x As Double, ByVal y As Double) As Long
    If tn >= 60000 Then TBFail "節点上限60000を超えます。間隔を大きくしてください。"
    tn = tn + 1
    If tn > nodeCap Then
        nodeCap = Application.Min(60000, nodeCap * 2)
        ReDim Preserve tx(1 To nodeCap): ReDim Preserve ty(1 To nodeCap)
    End If
    tx(tn) = x: ty(tn) = y: TBNode = tn
End Function

Private Function TBVertex(ByVal key As String, ByVal x As Double, ByVal y As Double) As Long
    If Not vertexMap.Exists(key) Then vertexMap.Add key, TBNode(x, y)
    TBVertex = CLng(vertexMap(key))
    If tx(TBVertex) <> x Or ty(TBVertex) <> y Then TBFail "境界の座標キーが一致しません。近接する点・共有点の座標を確認してください。"
End Function

Private Function TBMid(ByVal a As Long, ByVal b As Long) As Long
    Dim key As String
    If a < b Then key = CStr(a) & ":" & CStr(b) Else key = CStr(b) & ":" & CStr(a)
    If Not midpointMap.Exists(key) Then midpointMap.Add key, TBNode((tx(a) + tx(b)) / 2#, (ty(a) + ty(b)) / 2#)
    TBMid = CLng(midpointMap(key))
End Function

Private Sub TBQuad(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal mat As Long)
    Dim grown() As Long, i As Long, j As Long
    If te >= 60000 Then TBFail "要素上限60000を超えます。間隔を大きくしてください。"
    te = te + 1
    If te > elemCap Then
        elemCap = Application.Min(60000, elemCap * 2): ReDim grown(1 To elemCap, 1 To 8)
        For i = 1 To te - 1
            For j = 1 To 8: grown(i, j) = tc(i, j): Next j
        Next i
        tc = grown: ReDim Preserve tm(1 To elemCap)
    End If
    tc(te, 1) = a: tc(te, 2) = b: tc(te, 3) = c: tc(te, 4) = d
    tc(te, 5) = TBMid(a, b): tc(te, 6) = TBMid(b, c): tc(te, 7) = TBMid(c, d): tc(te, 8) = TBMid(d, a): tm(te) = mat
End Sub

Private Function TBCross(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double) As Double
    TBCross = ax * by - ay * bx
End Function

Private Function TBAngle(ByVal x As Double, ByVal y As Double) As Double
    Dim pi As Double
    pi = 4# * Atn(1#)
    If x = 0# Then
        If y >= 0# Then TBAngle = pi / 2# Else TBAngle = 3# * pi / 2#
    Else
        TBAngle = Atn(y / x)
        If x < 0# Then TBAngle = TBAngle + pi
        If TBAngle < 0# Then TBAngle = TBAngle + 2# * pi
    End If
End Function

Private Function TBRay(ByVal index As Long, ByVal cx As Double, ByVal cy As Double, ByVal angle As Double) As Double
    Dim p As Variant, i As Long, j As Long, ex As Double, ey As Double, dx As Double, dy As Double, den As Double, t As Double, u As Double
    p = loops(index): dx = Cos(angle): dy = Sin(angle): TBRay = 1E+100
    For i = 1 To UBound(p, 1)
        j = i Mod UBound(p, 1) + 1: ex = p(j, 1) - p(i, 1): ey = p(j, 2) - p(i, 2)
        den = TBCross(dx, dy, ex, ey)
        If Abs(den) > 0.000000000001 * Sqr(ex * ex + ey * ey) Then
            t = TBCross(p(i, 1) - cx, p(i, 2) - cy, ex, ey) / den
            u = TBCross(p(i, 1) - cx, p(i, 2) - cy, dx, dy) / den
            If t > 0# And u >= -0.0000000001 And u <= 1.0000000001 Then
                If t < TBRay Then TBRay = t
            End If
        End If
    Next i
    If TBRay = 1E+100 Then TBFail "空洞中心から境界に到達できません。凸形状の入れ子に分割してください。"
End Function

Private Sub TBTunnel()
    Dim p As Variant, q As Variant, cx As Double, cy As Double, i As Long, j As Long, k As Long, n As Long, sign As Double, z As Double
    Dim angles() As Double, angleN As Long, layer() As Long, bx() As Double, by() As Double, steps() As Long, cols() As Long
    Dim count As Long, radius As Double, gap As Double, length As Double, f As Double, g As Double, sector As Long, col As Long
    Dim ratios() As Double, maxRadius As Double, innerRadius As Double, minChord As Double
    Dim row As Long, ring As Long, totalRows As Long, totalCols As Long, grid() As Long, a As Long, b As Long, c As Long, d As Long, mat As Long
    If holeN <> 1 Then TBFail "斜辺を含む複合形状は、凸多角形の空洞1つと、その周囲を囲む凸領域の入れ子に対応します。段差・擁壁は水平・鉛直の境界を使用してください。"
    p = loops(loopN)
    For i = 1 To UBound(p, 1): cx = cx + p(i, 1): cy = cy + p(i, 2): Next i
    cx = cx / UBound(p, 1): cy = cy / UBound(p, 1)
    ReDim layer(0 To outerN): layer(0) = loopN
    For i = 1 To outerN: layer(i) = i: Next i
    count = 0
    For i = 1 To loopN: p = loops(i): count = count + UBound(p, 1): Next i
    ReDim angles(1 To count)
    For i = 0 To outerN
        p = loops(layer(i)): n = UBound(p, 1): sign = 0#
        If Not TBInside(cx, cy, layer(i)) Then TBFail "空洞・覆工・外周が同じ中心を囲んでいません。"
        For j = 1 To n
            k = j Mod n + 1
            z = TBCross(p(k, 1) - p(j, 1), p(k, 2) - p(j, 2), p(k Mod n + 1, 1) - p(k, 1), p(k Mod n + 1, 2) - p(k, 2))
            If z <> 0# Then
                If sign * z < 0# Then TBFail "放射格子の境界は凸多角形で入力してください。"
                sign = z
            End If
            angleN = angleN + 1: angles(angleN) = TBAngle(p(j, 1) - cx, p(j, 2) - cy)
        Next j
    Next i
    TBSort angles, 1, angleN
    count = 1
    For i = 2 To angleN
        If angles(i) - angles(count) > 0.0000000001 Then count = count + 1: angles(count) = angles(i)
    Next i
    If count > 1 Then
        If angles(1) + 8# * Atn(1#) - angles(count) <= 0.0000000001 Then count = count - 1
    End If
    angleN = count: ReDim Preserve angles(1 To angleN + 1): angles(angleN + 1) = angles(1) + 8# * Atn(1#)
    ReDim bx(0 To outerN, 1 To angleN + 1): ReDim by(0 To outerN, 1 To angleN + 1)
    For i = 0 To outerN
        For j = 1 To angleN + 1
            radius = TBRay(layer(i), cx, cy, angles(j))
            bx(i, j) = cx + radius * Cos(angles(j)): by(i, j) = cy + radius * Sin(angles(j))
            If i > 0 Then
                gap = Sqr((bx(i, j) - bx(i - 1, j)) ^ 2 + (by(i, j) - by(i - 1, j)) ^ 2)
                If TBRay(layer(i - 1), cx, cy, angles(j)) >= radius - modelSize * 0.0000000001 Then TBFail "外側境界は内側の材料領域から外側へ、完全に入れ子の順で入力してください。"
            End If
        Next j
        bx(i, angleN + 1) = bx(i, 1): by(i, angleN + 1) = by(i, 1)
    Next i
    ReDim steps(1 To outerN): ReDim cols(1 To angleN): ReDim ratios(1 To outerN)
    For i = 1 To outerN
        maxRadius = 0#: ratios(i) = 1#
        For j = 1 To angleN
            radius = Sqr((bx(i, j) - cx) ^ 2 + (by(i, j) - cy) ^ 2)
            innerRadius = Sqr((bx(i - 1, j) - cx) ^ 2 + (by(i - 1, j) - cy) ^ 2)
            maxRadius = Application.Max(maxRadius, radius)
            ratios(i) = Application.Max(ratios(i), radius / innerRadius)
        Next j
        steps(i) = Application.Max(1, CLng(-Int(-Log(ratios(i)) / Log(1# + stepSize / maxRadius))))
        totalRows = totalRows + steps(i)
    Next i
    For j = 1 To angleN
        length = 0#
        For i = 0 To outerN: length = Application.Max(length, Sqr((bx(i, j + 1) - bx(i, j)) ^ 2 + (by(i, j + 1) - by(i, j)) ^ 2)): Next i
        cols(j) = CLng(-Int(-length / stepSize)): totalCols = totalCols + cols(j)
    Next j
    ' Polygon vertex density can be finer than the requested h (circle approximation).
    ' Refine radial layers as well, without weakening the aspect/angle acceptance criteria.
    totalRows = 0
    For i = 1 To outerN
        minChord = 1E+100: gap = 0#
        For j = 1 To angleN
            length = Sqr((bx(i - 1, j + 1) - bx(i - 1, j)) ^ 2 + (by(i - 1, j + 1) - by(i - 1, j)) ^ 2) / cols(j)
            minChord = Application.Min(minChord, length)
            gap = Application.Max(gap, Sqr((bx(i, j) - bx(i - 1, j)) ^ 2 + (by(i, j) - by(i - 1, j)) ^ 2))
        Next j
        If minChord <= 0# Then TBFail "放射格子に長さ0の辺があります。"
        steps(i) = Application.Max(steps(i), CLng(-Int(-gap / (3# * minChord))))
        totalRows = totalRows + steps(i)
    Next i
    If CDbl(totalRows + 1) * totalCols + 2# * totalRows * totalCols + totalCols > 60000# Then TBFail "放射格子が節点上限60000を超えます。間隔を大きくしてください。"
    ReDim grid(0 To totalRows, 0 To totalCols - 1)
    row = 0
    For ring = 1 To outerN
        For k = 0 To steps(ring)
            If ring = 1 Or k > 0 Then
                g = (ratios(ring) ^ (CDbl(k) / steps(ring)) - 1#) / (ratios(ring) - 1#): col = 0
                For sector = 1 To angleN
                    For j = 0 To cols(sector) - 1
                        f = CDbl(j) / cols(sector)
                        grid(row, col) = TBNode((1# - g) * (bx(ring - 1, sector) * (1# - f) + bx(ring - 1, sector + 1) * f) + g * (bx(ring, sector) * (1# - f) + bx(ring, sector + 1) * f), _
                            (1# - g) * (by(ring - 1, sector) * (1# - f) + by(ring - 1, sector + 1) * f) + g * (by(ring, sector) * (1# - f) + by(ring, sector + 1) * f))
                        col = col + 1
                    Next j
                Next sector
                row = row + 1
            End If
        Next k
    Next ring
    row = 0
    For ring = 1 To outerN
        If outerN = 1 Then mat = MeshReadMaterialNumber() Else mat = ring
        For k = 1 To steps(ring)
            For col = 0 To totalCols - 1
                a = grid(row, col): b = grid(row + 1, col): c = grid(row + 1, (col + 1) Mod totalCols): d = grid(row, (col + 1) Mod totalCols)
                TBQuad a, b, c, d, mat
            Next col
            row = row + 1
        Next k
    Next ring
End Sub

Private Function TBPolygonArea(ByVal index As Long) As Double
    Dim p As Variant, i As Long, j As Long
    p = loops(index)
    For i = 1 To UBound(p, 1)
        j = i Mod UBound(p, 1) + 1
        TBPolygonArea = TBPolygonArea + (p(i, 1) - p(1, 1)) * (p(j, 2) - p(1, 2)) - (p(j, 1) - p(1, 1)) * (p(i, 2) - p(1, 2))
    Next i
    TBPolygonArea = Abs(TBPolygonArea) / 2#
End Function

Private Sub TBValidate()
    Dim e As Long, i As Long, j As Long, k As Long, cn() As Long, du(1 To 4) As Double, dv(1 To 4) As Double
    Dim u As Double, v As Double, xu As Double, yu As Double, xv As Double, yv As Double, det As Double, memory As Double, band As Long
    Dim ax As Double, ay As Double, bx As Double, by As Double, cosine As Double, angle As Double
    Dim materialArea As Object, key As String, area As Double, expected As Double, mat As Long
    Dim edgeOwner As Object, elementParent() As Long, components As Long, owner As Long
    band = CLng(8# * Sqr(CDbl(tn))): If band > 2 * tn - 1 Then band = 2 * tn - 1
    memory = 4# * 8# * (2# * tn) * (band + 1#) + te * 16384#
    If memory > FEMReadSetting("SOLVER_MEMORY_LIMIT_MB", 512#) * 1048576# Then TBFail "概算の解析メモリが上限を超えます。間隔を大きくしてください。"
    ReDim cn(1 To te, 1 To 8)
    Set materialArea = CreateObject("Scripting.Dictionary")
    For e = 1 To te
        For k = 1 To 8: cn(e, k) = tc(e, k): Next k
        area = 0#
        For k = 1 To 4
            area = area + (tx(tc(e, k)) - tx(tc(e, 1))) * (ty(tc(e, k Mod 4 + 1)) - ty(tc(e, 1))) - (tx(tc(e, k Mod 4 + 1)) - tx(tc(e, 1))) * (ty(tc(e, k)) - ty(tc(e, 1)))
        Next k
        key = CStr(tm(e))
        If Not materialArea.Exists(key) Then materialArea.Add key, 0#
        materialArea(key) = materialArea(key) + area / 2#
        For k = 1 To 4
            ax = tx(tc(e, k Mod 4 + 1)) - tx(tc(e, k)): ay = ty(tc(e, k Mod 4 + 1)) - ty(tc(e, k))
            bx = tx(tc(e, (k + 2) Mod 4 + 1)) - tx(tc(e, k)): by = ty(tc(e, (k + 2) Mod 4 + 1)) - ty(tc(e, k))
            If (ax * ax + ay * ay) * (bx * bx + by * by) <= 0# Then TBFail "要素辺の長さが0です。"
            cosine = (ax * bx + ay * by) / Sqr((ax * ax + ay * ay) * (bx * bx + by * by))
            cosine = Application.Max(-1#, Application.Min(1#, cosine))
            angle = Application.Acos(cosine) * 45# / Atn(1#)
            If angle < 29.9999999 Or angle > 150.0000001 Then TBFail "内角30～150度を満たしません。要素=" & CStr(e)
        Next k
        For i = 0 To 4
            u = -1# + CDbl(i) / 2#
            For j = 0 To 4
                v = -1# + CDbl(j) / 2#
                du(1) = -(1# - v) / 4#: du(2) = (1# - v) / 4#: du(3) = (1# + v) / 4#: du(4) = -(1# + v) / 4#
                dv(1) = -(1# - u) / 4#: dv(2) = -(1# + u) / 4#: dv(3) = (1# + u) / 4#: dv(4) = (1# - u) / 4#
                xu = 0#: yu = 0#: xv = 0#: yv = 0#
                For k = 1 To 4
                    xu = xu + du(k) * tx(tc(e, k)): yu = yu + du(k) * ty(tc(e, k))
                    xv = xv + dv(k) * tx(tc(e, k)): yv = yv + dv(k) * ty(tc(e, k))
                Next k
                det = xu * yv - xv * yu
                If det <= modelSize * modelSize * 0.0000000000001 Then TBFail "Q8の25点Jacobianが不足します。要素=" & CStr(e)
            Next j
        Next i
    Next e
    For i = 1 To outerN
        If outerN = 1 Then mat = MeshReadMaterialNumber() Else mat = i
        key = CStr(mat)
        If Not materialArea.Exists(key) Then TBFail "材料領域が空です。外側境界の順序・重なり・穴を確認してください。材料=" & key
        If radialUsed Or blocksUsed Then
            expected = TBPolygonArea(i)
            If radialUsed Then
                If i = 1 Then expected = expected - TBPolygonArea(loopN) Else expected = expected - TBPolygonArea(i - 1)
            End If
            If Abs(materialArea(key) - expected) > 0.00000001 * (1# + Abs(expected)) Then TBFail "ブロック格子の材料面積が入力多角形と一致しません。材料=" & key
        End If
    Next i
    ' Point-only contact is not a connected continuum block mesh.
    Set edgeOwner = CreateObject("Scripting.Dictionary"): ReDim elementParent(1 To te)
    For e = 1 To te: elementParent(e) = e: Next e
    For e = 1 To te
        For k = 1 To 4
            key = TBEdgeKey(tc(e, k), tc(e, k Mod 4 + 1))
            If edgeOwner.Exists(key) Then
                owner = CLng(edgeOwner(key))
                elementParent(TBRoot(elementParent, e)) = TBRoot(elementParent, owner)
            Else
                edgeOwner.Add key, e
            End If
        Next k
    Next e
    For e = 1 To te: If TBRoot(elementParent, e) = e Then components = components + 1
    Next e
    If components <> 1 Then TBFail "要素が辺全体で連結していません。点だけの接続・離れた領域を見直してください。"
    P5ValidateMeshQuality tn, te, cn, tx, ty
    If Not P5QualityMeetsSettings() And FEMReadSetting("MESH_T6_FALLBACK", 0#) = 0# Then TBFail "品質目標を満たしません。細い領域・近接した境界点・放射方向の分割を見直してください。" & P5QualitySummary()
End Sub
