Option Explicit
Public MeshT6LastFlipCount As Long

' Shape targets are separate from validity. A constrained sharp boundary may
' miss a target; it must be reported, never accepted by relaxing the target.
Public Function MeshShapeWithinTargets(ByVal triangle As Boolean, ByVal minAngle As Double, ByVal maxAngle As Double, ByVal aspect As Double) As Boolean
    Dim lo As Double, hi As Double, ratio As Double
    If triangle Then
        lo = FEMReadSetting("MESH_T6_MIN_ANGLE_DEG", 20#)
        hi = FEMReadSetting("MESH_T6_MAX_ANGLE_DEG", 140#)
        ratio = FEMReadSetting("MESH_T6_MAX_ASPECT_RATIO", 5#)
    Else
        lo = FEMReadSetting("MESH_QUALITY_MIN_ANGLE_DEG", 45#)
        hi = FEMReadSetting("MESH_QUALITY_MAX_ANGLE_DEG", 135#)
        ratio = FEMReadSetting("MESH_QUALITY_MAX_ASPECT_RATIO", 5#)
    End If
    If lo <= 0# Or hi >= 180# Or hi < lo Or ratio < 1# Then Err.Raise vbObjectError + 3980, , "メッシュ形状の品質目標が不正です。"
    If triangle And (lo > 60# Or hi < 60#) Then Err.Raise vbObjectError + 3980, , "T6の角度目標は最小60度以下・最大60度以上にしてください。"
    If Not triangle And (lo > 90# Or hi < 90#) Then Err.Raise vbObjectError + 3980, , "Q8の角度目標は最小90度以下・最大90度以上にしてください。"
    MeshShapeWithinTargets = (minAngle >= lo - 0.0000001 And maxAngle <= hi + 0.0000001 And aspect <= ratio + 0.0000001)
End Function

Public Function MeshQ8JacobianWithinTarget(ByVal area As Double, ByVal minJacobian As Double) As Boolean
    Dim target As Double
    target = FEMReadSetting("MESH_Q8_MIN_JACOBIAN_AREA_RATIO", 0.1)
    If target <= 0# Or target > 1# Then Err.Raise vbObjectError + 3980, , "Q8の正規化Jacobian目標は0より大きく1以下にしてください。"
    If area > 0# Then MeshQ8JacobianWithinTarget = (4# * minJacobian / area >= target - 0.0000001)
End Function

' Mean-ratio quality: 1 for an equilateral triangle, 0 at degeneration.
Public Function MeshTriangleScore(ByRef x() As Double, ByRef y() As Double, ByVal a As Long, ByVal b As Long, ByVal c As Long) As Double
    Dim cross As Double, sum As Double
    cross = (x(b) - x(a)) * (y(c) - y(a)) - (y(b) - y(a)) * (x(c) - x(a))
    If cross <= 0# Then Exit Function
    sum = (x(b) - x(a)) ^ 2 + (y(b) - y(a)) ^ 2 + (x(c) - x(b)) ^ 2 + (y(c) - y(b)) ^ 2 + (x(a) - x(c)) ^ 2 + (y(a) - y(c)) ^ 2
    If sum > 0# Then MeshTriangleScore = 2# * Sqr(3#) * cross / sum
End Function

Public Function MeshTriangleWithinTargets(ByRef x() As Double, ByRef y() As Double, ByVal a As Long, ByVal b As Long, ByVal c As Long) As Boolean
    Dim ids As Variant, k As Long, i As Long, j As Long, h As Long
    Dim ux As Double, uy As Double, vx As Double, vy As Double, cs As Double, angle As Double
    Dim lo As Double, hi As Double, shortest As Double, longest As Double, ln As Double
    ids = Array(a, b, c): lo = 180#: shortest = 1E+308
    If MeshTriangleScore(x, y, a, b, c) <= 0# Then Exit Function
    For k = 0 To 2
        i = ids(k): j = ids((k + 1) Mod 3): h = ids((k + 2) Mod 3)
        ux = x(j) - x(i): uy = y(j) - y(i): vx = x(h) - x(i): vy = y(h) - y(i)
        ln = Sqr(ux * ux + uy * uy)
        If ln < shortest Then shortest = ln
        If ln > longest Then longest = ln
        cs = (ux * vx + uy * vy) / (ln * Sqr(vx * vx + vy * vy))
        cs = Application.Max(-1#, Application.Min(1#, cs))
        angle = Application.Acos(cs) * 180# / 3.14159265358979
        If angle < lo Then lo = angle
        If angle > hi Then hi = angle
    Next k
    MeshTriangleWithinTargets = MeshShapeWithinTargets(True, lo, hi, longest / shortest)
End Function

Public Sub MeshQualityEnsureSettings()
    Dim keys As Variant, vals As Variant, labels As Variant, units As Variant
    Dim ws As Worksheet, page As Worksheet, i As Long, r As Long, last As Long, found As Boolean, pr As Long
    keys = Array("MESH_T6_MIN_ANGLE_DEG", "MESH_T6_MAX_ANGLE_DEG", "MESH_T6_MAX_ASPECT_RATIO", "MESH_Q8_MIN_JACOBIAN_AREA_RATIO", "MESH_T6_SMOOTH_ITERATIONS", "MESH_QUALITY_REFINE", "MESH_QUALITY_REFINE_PASSES", "MESH_QUALITY_REFINE_MAX_RATIO", "MESH_QUALITY_REFINE_FRACTION", "MESH_T6_PATCH_ENABLE", "MESH_T6_PATCH_ITERATIONS", "MESH_T6_PATCH_REBUILD")
    vals = Array(20#, 140#, 5#, 0.1, 10#, 0#, 1#, 1.12, 0.03, 1#, 40#, 1#)
    labels = Array("T6 最小角度目標", "T6 最大角度目標", "T6 最大辺長比", "Q8 正規化Jacobian目標", "T6 平滑化の最大反復", "品質による局所細分化", "局所細分化の最大回数", "局所細分化の要素数上限倍率", "局所細分化の対象割合", "T6周辺の局所節点改善", "T6局所改善の最大反復", "T6局所パッチ再分割")
    units = Array("deg。境界制約で未達の場合は警告", "deg。Q8の角度目標と独立", "最長辺/最短辺", "4×最小Jacobian/面積。長方形=1", "0=省略 / 1～50。境界・材料境界・Q8を保持", "0=OFF / 1=ON。デローニのQ8修正後・T6置換前に比較", "1～3。改善しない場合は中止", "元のQ8要素数に対する倍率。1.01～2", "0より大きく1以下。悪い要素から候補選定", "0=OFF / 1=ON。Q8の品質と材料境界を保持", "0～100。境界上は直線に沿う移動のみ", "0=OFF / 1=ON。自由内点の整理で改善した候補だけ採用")
    Set ws = ThisWorkbook.Worksheets("設定"): Set page = ThisWorkbook.Worksheets("メッシュ設定")
    For i = 0 To UBound(keys)
        found = False: last = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
        For r = 5 To last
            If CStr(ws.Cells(r, 5).Value2) = keys(i) Then found = True: Exit For
        Next r
        If Not found Then
            r = last + 1
            ws.Cells(r, 1).Value2 = "メッシュ品質": ws.Cells(r, 2).Value2 = labels(i)
            ws.Cells(r, 3).Value2 = vals(i): ws.Cells(r, 4).Value2 = units(i): ws.Cells(r, 5).Value2 = keys(i)
        End If
        found = False
        For pr = 8 To page.Cells(page.Rows.Count, 6).End(xlUp).Row
            If CStr(page.Cells(pr, 6).Value2) = keys(i) Then found = True: Exit For
        Next pr
        If Not found Then
            pr = page.Cells(page.Rows.Count, 6).End(xlUp).Row + 1
            If pr < 50 Then pr = 50
            page.Range("B30:G30").Copy Destination:=page.Cells(pr, 2)
            page.Rows(pr).RowHeight = 48
            page.Cells(pr, 2).Value2 = labels(i): page.Cells(pr, 3).Value2 = ws.Cells(r, 3).Value2
            page.Cells(pr, 4).ClearContents: page.Cells(pr, 5).Value2 = units(i): page.Cells(pr, 6).Value2 = keys(i)
            page.Cells(pr, 7).Value2 = vals(i)
            page.Cells(pr, 3).Validation.Delete
        End If
    Next i
    Application.CutCopyMode = False
    FEMInvalidateSettingCache
End Sub

' Only T6 interior vertices move. Q8 vertices, exterior/material interfaces,
' curved edges and nodes carrying conditions are held fixed. A local move is
' accepted only if the worst mean-ratio improves and target misses do not grow.
Public Sub MeshImproveT6(ByRef x() As Double, ByRef y() As Double, ByRef rows() As Variant, ByRef conditions As Variant, ByRef moved As Long, ByRef sweeps As Long, Optional ByVal flipEdges As Boolean = True)
    Dim n As Long, ne As Long, e As Long, j As Long, a As Long, b As Long, m As Long, k As Long, t As Long, limit As Long
    Dim tri() As Boolean, fixed() As Boolean, touched() As Boolean, adj() As Collection, edges As Object, key As Variant, info As Variant, item As Variant
    Dim ox As Double, oy As Double, tx As Double, ty As Double, score As Double, best As Double, q As Double, stepSize As Double, count As Long, misses As Long, oldMisses As Long, changed As Long
    n = UBound(x): ne = UBound(rows, 1): moved = 0: sweeps = 0
    limit = CLng(FEMReadSetting("MESH_T6_SMOOTH_ITERATIONS", 10#)): If limit < 0 Or limit > 50 Then Err.Raise vbObjectError + 3980, , "T6平滑化反復は0～50です。"
    MeshT6LastFlipCount = 0
    If limit = 0 Then Exit Sub
    If flipEdges Then MeshFlipT6Edges x, y, rows, conditions, Application.Min(limit, 5), MeshT6LastFlipCount
    ReDim tri(1 To ne): ReDim fixed(1 To n): ReDim touched(1 To n): ReDim adj(1 To n)
    For a = 1 To n: Set adj(a) = New Collection: Next a
    Set edges = CreateObject("Scripting.Dictionary")
    For e = 1 To ne
        tri(e) = (CStr(rows(e, 11)) = "T6")
        For j = 1 To IIf(tri(e), 3, 4)
            a = CLng(rows(e, j + 1)): b = CLng(rows(e, (j Mod IIf(tri(e), 3, 4)) + 2))
            m = CLng(rows(e, IIf(tri(e) And j = 3, 9, j + 5)))
            If Not tri(e) Then fixed(a) = True Else adj(a).Add e
            If a < b Then key = CStr(a) & ":" & CStr(b) Else key = CStr(b) & ":" & CStr(a)
            If edges.Exists(key) Then
                info = edges(key)
                If info(0) <> 1 Or CLng(info(4)) <> m Then Err.Raise vbObjectError + 3976, , "T6平滑化の辺接続が不正です。"
                If info(3) <> rows(e, 10) Then fixed(a) = True: fixed(b) = True
                info(0) = 2: edges(key) = info
            Else
                edges.Add key, Array(1, a, b, rows(e, 10), m)
            End If
            If Abs(x(m) - (x(a) + x(b)) / 2#) + Abs(y(m) - (y(a) + y(b)) / 2#) > 0.000000001 * (1# + Abs(x(a) - x(b)) + Abs(y(a) - y(b))) Then fixed(a) = True: fixed(b) = True
            If a <= UBound(conditions, 1) Then
                For k = 4 To UBound(conditions, 2)
                    If Val(CStr(conditions(a, k))) <> 0# Then fixed(a) = True
                Next k
            End If
            If m <= UBound(conditions, 1) Then
                For k = 4 To UBound(conditions, 2)
                    If Val(CStr(conditions(m, k))) <> 0# Then fixed(a) = True: fixed(b) = True
                Next k
            End If
        Next j
    Next e
    For Each key In edges.Keys
        info = edges(key): If info(0) = 1 Then fixed(info(1)) = True: fixed(info(2)) = True
    Next key
    For t = 1 To limit
        changed = 0
        For a = 1 To n
            If fixed(a) Or adj(a).Count = 0 Then GoTo NextNode
            tx = 0#: ty = 0#: count = 0: best = 1#: oldMisses = 0
            For Each item In adj(a)
                e = CLng(item): q = MeshTriangleScore(x, y, CLng(rows(e, 2)), CLng(rows(e, 3)), CLng(rows(e, 4)))
                If q < best Then best = q
                If Not MeshTriangleWithinTargets(x, y, CLng(rows(e, 2)), CLng(rows(e, 3)), CLng(rows(e, 4))) Then oldMisses = oldMisses + 1
                For j = 2 To 4
                    b = CLng(rows(e, j)): If b <> a Then tx = tx + x(b): ty = ty + y(b): count = count + 1
                Next j
            Next item
            ox = x(a): oy = y(a): tx = tx / count: ty = ty / count: stepSize = 1#
            For k = 1 To 8
                x(a) = ox + stepSize * (tx - ox): y(a) = oy + stepSize * (ty - oy): score = 1#: misses = 0
                For Each item In adj(a)
                    e = CLng(item): q = MeshTriangleScore(x, y, CLng(rows(e, 2)), CLng(rows(e, 3)), CLng(rows(e, 4)))
                    If q < score Then score = q
                    If Not MeshTriangleWithinTargets(x, y, CLng(rows(e, 2)), CLng(rows(e, 3)), CLng(rows(e, 4))) Then misses = misses + 1
                Next item
                If score > best + 0.00000001 And misses <= oldMisses Then changed = changed + 1: touched(a) = True: Exit For
                x(a) = ox: y(a) = oy: stepSize = stepSize / 2#
            Next k
NextNode:
        Next a
        sweeps = t: If changed = 0 Then Exit For
    Next t
    For Each key In edges.Keys
        info = edges(key): a = info(1): b = info(2): m = info(4)
        If touched(a) Or touched(b) Then x(m) = (x(a) + x(b)) / 2#: y(m) = (y(a) + y(b)) / 2#
    Next key
    For a = 1 To n: If touched(a) Then moved = moved + 1
    Next a
End Sub

' Flip only interior edges of two T6 of the same material. The midpoint is
' reused only if it carries no condition. Exterior edges and Q8 stay fixed.
Public Sub MeshFlipT6Edges(ByRef x() As Double, ByRef y() As Double, ByRef rows() As Variant, ByRef conditions As Variant, ByVal limit As Long, ByRef flips As Long)
    Dim edges As Object, mids As Object, key As Variant, info As Variant, used() As Boolean
    Dim pass As Long, e As Long, j As Long, f As Long, a As Long, b As Long, c As Long, d As Long, m As Long, k As Long, changed As Long
    Dim score0 As Double, score1 As Double, misses0 As Long, misses1 As Long, safe As Boolean
    Dim midBC As Long, midCA As Long, midAD As Long, midDB As Long
    flips = 0
    For pass = 1 To limit
        Set edges = CreateObject("Scripting.Dictionary"): Set mids = CreateObject("Scripting.Dictionary")
        ReDim used(1 To UBound(rows, 1)): changed = 0
        For e = 1 To UBound(rows, 1)
            For j = 1 To IIf(CStr(rows(e, 11)) = "T6", 3, 4)
                a = rows(e, j + 1): b = rows(e, (j Mod IIf(CStr(rows(e, 11)) = "T6", 3, 4)) + 2)
                m = rows(e, IIf(CStr(rows(e, 11)) = "T6" And j = 3, 9, j + 5))
                key = MeshQualityEdgeKey(a, b): mids(key) = m
                If edges.Exists(key) Then
                    info = edges(key): info(1) = e: edges(key) = info
                Else
                    edges.Add key, Array(e, 0, a, b, m)
                End If
            Next j
        Next e
        For Each key In edges.Keys
            info = edges(key): e = info(0): f = info(1)
            If f = 0 Then GoTo NextEdge
            If used(e) Or used(f) Then GoTo NextEdge
            If CStr(rows(e, 11)) <> "T6" Or CStr(rows(f, 11)) <> "T6" Or rows(e, 10) <> rows(f, 10) Then GoTo NextEdge
            a = info(2): b = info(3): m = info(4): c = 0: d = 0
            For j = 2 To 4
                If rows(e, j) <> a And rows(e, j) <> b Then c = rows(e, j)
                If rows(f, j) <> a And rows(f, j) <> b Then d = rows(f, j)
            Next j
            If c = 0 Or d = 0 Or c = d Then GoTo NextEdge
            If edges.Exists(MeshQualityEdgeKey(c, d)) Then GoTo NextEdge
            safe = True
            If m <= UBound(conditions, 1) Then
                For k = 4 To UBound(conditions, 2)
                    If Val(CStr(conditions(m, k))) <> 0# Then safe = False
                Next k
            End If
            If Not safe Then GoTo NextEdge
            score0 = Application.Min(MeshTriangleScore(x, y, a, b, c), MeshTriangleScore(x, y, b, a, d))
            score1 = Application.Min(MeshTriangleScore(x, y, c, d, b), MeshTriangleScore(x, y, d, c, a))
            If score1 <= score0 + 0.00000001 Then GoTo NextEdge
            misses0 = 0: misses1 = 0
            If Not MeshTriangleWithinTargets(x, y, a, b, c) Then misses0 = misses0 + 1
            If Not MeshTriangleWithinTargets(x, y, b, a, d) Then misses0 = misses0 + 1
            If Not MeshTriangleWithinTargets(x, y, c, d, b) Then misses1 = misses1 + 1
            If Not MeshTriangleWithinTargets(x, y, d, c, a) Then misses1 = misses1 + 1
            If misses1 > misses0 Then GoTo NextEdge
            midBC = mids(MeshQualityEdgeKey(b, c)): midCA = mids(MeshQualityEdgeKey(c, a))
            midAD = mids(MeshQualityEdgeKey(a, d)): midDB = mids(MeshQualityEdgeKey(d, b))
            MeshQualityWriteT6 rows, e, c, d, b, m, midDB, midBC
            MeshQualityWriteT6 rows, f, d, c, a, m, midCA, midAD
            x(m) = (x(c) + x(d)) / 2#: y(m) = (y(c) + y(d)) / 2#
            used(e) = True: used(f) = True: changed = changed + 1: flips = flips + 1
NextEdge:
        Next key
        If changed = 0 Then Exit For
    Next pass
End Sub

Private Function MeshQualityEdgeKey(ByVal a As Long, ByVal b As Long) As String
    If a < b Then MeshQualityEdgeKey = CStr(a) & ":" & CStr(b) Else MeshQualityEdgeKey = CStr(b) & ":" & CStr(a)
End Function
Private Sub MeshQualityWriteT6(ByRef rows() As Variant, ByVal e As Long, ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal ab As Long, ByVal bc As Long, ByVal ca As Long)
    rows(e, 2) = a: rows(e, 3) = b: rows(e, 4) = c: rows(e, 5) = c
    rows(e, 6) = ab: rows(e, 7) = bc: rows(e, 8) = c: rows(e, 9) = ca
End Sub

Public Sub MeshImproveExistingT6Sheet()
    Dim wn As Worksheet, we As Worksheet, ns As Variant, rows() As Variant, x() As Double, y() As Double, cn() As Long, coords() As Variant
    Dim n As Long, ne As Long, e As Long, j As Long, moved As Long, sweeps As Long
    If AnalysisRunning Then Err.Raise 5, , "解析中です。"
    Set wn = ThisWorkbook.Worksheets("節点データ"): Set we = ThisWorkbook.Worksheets("要素データ")
    n = wn.Cells(wn.Rows.Count, 1).End(xlUp).Row - 1: ne = we.Cells(we.Rows.Count, 1).End(xlUp).Row - 1
    ns = wn.Range(wn.Cells(2, 1), wn.Cells(n + 1, 14)).Value2: rows = we.Range(we.Cells(2, 1), we.Cells(ne + 1, 11)).Value2
    ReDim x(1 To n): ReDim y(1 To n): ReDim cn(1 To ne, 1 To 8): ReDim coords(1 To n, 1 To 2)
    For j = 1 To n: x(j) = CDbl(ns(j, 2)): y(j) = CDbl(ns(j, 3)): Next j
    MeshImproveT6 x, y, rows, ns, moved, sweeps
    For e = 1 To ne: For j = 1 To 8: cn(e, j) = CLng(rows(e, j + 1)): Next j: Next e
    P5ValidateMeshQuality n, ne, cn, x, y
    For j = 1 To n: coords(j, 1) = x(j): coords(j, 2) = y(j): Next j
    wn.Range(wn.Cells(2, 2), wn.Cells(n + 1, 3)).Value2 = coords
    we.Range(we.Cells(2, 1), we.Cells(ne + 1, 11)).Value2 = rows
    FEMInvalidateRemeshedModel
    FEMWriteTextSetting "MESH_GEN_LAST_T6_DETAIL", FEMReadTextSetting("MESH_GEN_LAST_T6_DETAIL", "") & " / T6追加辺交換=" & CStr(MeshT6LastFlipCount) & " / 追加移動節点=" & CStr(moved)
    P5EvaluateCurrentSheet "T6の辺交換・平滑化後"
End Sub
