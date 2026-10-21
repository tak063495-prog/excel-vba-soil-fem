Option Explicit

' 増分ひずみのSVGオーバーレイ。入力はpixel座標、pixel Yは下向き。
Public Function FEMIncrementOverlaySvg(ByVal cells As Object, ByVal shearOn As Boolean, _
    ByVal crossOn As Boolean, ByVal contourMode As Boolean, ByVal vectorScale As Double, _
    ByRef legend As String) As String
    Dim s As String, key As Variant, a As Variant, maxG As Double, maxP As Double, refHalf As Double, em As Double
    Dim nodes As Object, n As Variant, v As Variant, i As Long, x As Double, y As Double, edgeCount As Long
    Set nodes = CreateObject("Scripting.Dictionary")
    For Each key In cells.Keys
        a = cells(key)
        If IsArray(a) Then
            If IsNumeric(a(8)) Then If Abs(CDbl(a(8))) > maxG Then maxG = Abs(CDbl(a(8)))
            FEMIncrementNodeAdd nodes, a(0), a(1), a(8)
            FEMIncrementNodeAdd nodes, a(2), a(3), a(8)
            FEMIncrementNodeAdd nodes, a(4), a(5), a(8)
            FEMIncrementNodeAdd nodes, a(6), a(7), a(8)
            If IsNumeric(a(9)) And IsNumeric(a(10)) And IsNumeric(a(11)) Then
                Dim q1 As Double, q2 As Double, qt As Double
                FEMStrainPrincipal2D CDbl(a(9)), CDbl(a(10)), CDbl(a(11)), q1, q2, qt
                If Abs(q1) > maxP Then maxP = Abs(q1)
                If Abs(q2) > maxP Then maxP = Abs(q2)
            End If
            em = FEMIncrementEdgeLength(a, 0, 1): If FEMIncrementEdgeLength(a, 1, 2) < em Then em = FEMIncrementEdgeLength(a, 1, 2)
            If FEMIncrementEdgeLength(a, 2, 3) < em Then em = FEMIncrementEdgeLength(a, 2, 3)
            If FEMIncrementEdgeLength(a, 3, 0) < em Then em = FEMIncrementEdgeLength(a, 3, 0)
            ' One very short edge must not shrink every cross to subpixel size.
            ' Use the mean of each element's shortest positive edge, with one common scale.
            If em > 0# Then refHalf = refHalf + em: edgeCount = edgeCount + 1
        End If
    Next key
    If edgeCount > 0 Then refHalf = refHalf / edgeCount
    s = "<g id='increment-overlay'>"
    If shearOn And contourMode And maxG > 0# Then
        s = s & "<g id='shear-increment-contours'>"
        For i = 1 To 3
            For Each key In cells.Keys
                a = cells(key)
                If IsArray(a) Then s = s & FEMIncrementCellContours(a, nodes, maxG * i / 4#, i)
            Next key
        Next i
        s = s & "</g>"
    Else
        s = s & "<g id='shear-increment-contours'/>"
    End If
    If crossOn And maxP > 0# Then
        s = s & "<g id='plastic-principal-crosses'>"
        For Each key In cells.Keys
            a = cells(key)
            If IsArray(a) Then s = s & FEMIncrementCross(a, maxP, refHalf, vectorScale, CStr(key))
        Next key
        s = s & "</g>"
    Else
        s = s & "<g id='plastic-principal-crosses'/>"
    End If
    s = s & "</g>"
    legend = "最大せん断ひずみ増分=" & FEMIncrementSci(maxG)
    If crossOn And maxP > 0# Then
        legend = legend & " / 塑性主ひずみ基準=" & FEMIncrementSci(maxP) & "（赤=引張、青=圧縮）"
    ElseIf crossOn Then
        legend = legend & " / 塑性主ひずみクロス=0（表示なし）"
    Else
        legend = legend & " / 塑性主ひずみクロスOFF"
    End If
    FEMIncrementOverlaySvg = s
End Function

Private Sub FEMIncrementNodeAdd(ByVal d As Object, ByVal x As Variant, ByVal y As Variant, ByVal g As Variant)
    Dim k As String, q As Variant
    If Not IsNumeric(x) Or Not IsNumeric(y) Or Not IsNumeric(g) Then Exit Sub
    k = FEMIncrementNum(CDbl(x), True) & ":" & FEMIncrementNum(CDbl(y), True)
    If d.Exists(k) Then
        q = d(k): q(2) = q(2) + CDbl(g): q(3) = q(3) + 1#: d(k) = q
    Else
        d.Add k, Array(CDbl(x), CDbl(y), CDbl(g), 1#)
    End If
End Sub

Private Function FEMIncrementNodeValue(ByVal d As Object, ByVal x As Double, ByVal y As Double) As Double
    Dim k As String, q As Variant
    k = FEMIncrementNum(x, True) & ":" & FEMIncrementNum(y, True)
    If d.Exists(k) Then q = d(k): FEMIncrementNodeValue = q(2) / q(3)
End Function

Private Function FEMIncrementCellContours(ByVal a As Variant, ByVal nodes As Object, ByVal level As Double, ByVal band As Long) As String
    Dim p(0 To 3, 0 To 1) As Double, z(0 To 3) As Double, out As String, t As Long
    p(0, 0) = a(0): p(0, 1) = a(1): p(1, 0) = a(2): p(1, 1) = a(3)
    p(2, 0) = a(4): p(2, 1) = a(5): p(3, 0) = a(6): p(3, 1) = a(7)
    For t = 0 To 3: z(t) = FEMIncrementNodeValue(nodes, p(t, 0), p(t, 1)): Next t
    out = FEMIncrementTriangleContour(p, z, 0, 1, 2, level, band) & FEMIncrementTriangleContour(p, z, 0, 2, 3, level, band)
    FEMIncrementCellContours = out
End Function

Private Function FEMIncrementTriangleContour(ByRef p() As Double, ByRef z() As Double, ByVal i0 As Long, ByVal i1 As Long, ByVal i2 As Long, ByVal lev As Double, ByVal band As Long) As String
    Dim ex(0 To 1) As Double, ey(0 To 1) As Double, n As Long, q As String
    FEMIncrementEdge p, z, i0, i1, lev, ex, ey, n
    FEMIncrementEdge p, z, i1, i2, lev, ex, ey, n
    FEMIncrementEdge p, z, i2, i0, lev, ex, ey, n
    If n = 2 And (Abs(ex(0) - ex(1)) > 0.000001 Or Abs(ey(0) - ey(1)) > 0.000001) Then q = "<path d='M " & FEMIncrementNum(ex(0)) & " " & FEMIncrementNum(ey(0)) & " L " & FEMIncrementNum(ex(1)) & " " & FEMIncrementNum(ey(1)) & "' stroke='" & Choose(band, "#f28e2b", "#e15759", "#b07aa1") & "' stroke-width='1.4' fill='none' data-level='" & FEMIncrementSci(lev) & "' data-band='" & band & "'/>"
    FEMIncrementTriangleContour = q
End Function

Private Sub FEMIncrementEdge(ByRef p() As Double, ByRef z() As Double, ByVal i As Long, ByVal j As Long, ByVal lev As Double, ByRef ex() As Double, ByRef ey() As Double, ByRef n As Long)
    Dim di As Boolean, dj As Boolean, f As Double
    di = (z(i) > lev): dj = (z(j) > lev)
    If di = dj Or z(i) = z(j) Then Exit Sub
    If n >= 2 Then Exit Sub
    f = (lev - z(i)) / (z(j) - z(i))
    ex(n) = p(i, 0) + f * (p(j, 0) - p(i, 0)): ey(n) = p(i, 1) + f * (p(j, 1) - p(i, 1)): n = n + 1
End Sub

Private Function FEMIncrementCross(ByVal a As Variant, ByVal maxP As Double, ByVal refHalf As Double, ByVal vectorMultiplier As Double, ByVal eid As String) As String
    Dim ex As Double, ey As Double, g As Double, p1 As Double, p2 As Double, th As Double, c As Double, sn As Double, h As Double
    ex = a(9): ey = a(10): g = a(11): FEMStrainPrincipal2D ex, ey, g, p1, p2, th
    c = Cos(th): sn = Sin(th): h = refHalf * 0.3
    If h > 12# Then h = 12#
    h = h * vectorMultiplier: If h <= 0# Then Exit Function
    FEMIncrementCross = FEMIncrementAxis(a, c, -sn, h * Abs(p1) / maxP, p1, eid, "max") & FEMIncrementAxis(a, -sn, -c, h * Abs(p2) / maxP, p2, eid, "min")
End Function

Private Function FEMIncrementAxis(ByVal a As Variant, ByVal dx As Double, ByVal dy As Double, ByVal h As Double, ByVal val As Double, ByVal eid As String, ByVal principal As String) As String
    Dim cx As Double, cy As Double, col As String, sign As Long, x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim head As Double, ux As Double, uy As Double, bX As Double, bY As Double, fragment As String
    If val = 0# Or h <= 0.000001 Then Exit Function
    cx = (a(0) + a(2) + a(4) + a(6)) / 4#: cy = (a(1) + a(3) + a(5) + a(7)) / 4#
    If val > 0# Then col = "#1f77b4" Else col = "#d62728"
    head = h * 0.35: If head > 2.4 Then head = 2.4
    For sign = -1 To 1 Step 2
        x1 = cx: y1 = cy: x2 = cx + sign * dx * h: y2 = cy + sign * dy * h
        ux = sign * dx: uy = sign * dy
        If val > 0# Then x1 = x2: y1 = y2: x2 = cx: y2 = cy: ux = -ux: uy = -uy
        bX = x2 - ux * head: bY = y2 - uy * head
        fragment = fragment & "<line x1='" & FEMIncrementNum(x1) & "' y1='" & FEMIncrementNum(y1) & "' x2='" & FEMIncrementNum(x2) & "' y2='" & FEMIncrementNum(y2) & "' stroke='" & col & "' stroke-width='0.85' data-element='" & FEMIncrementXml(eid) & "' data-principal='" & principal & "' data-strain='" & FEMIncrementSci(val) & "'/>"
        fragment = fragment & "<path d='M " & FEMIncrementNum(x2) & " " & FEMIncrementNum(y2) & " L " & FEMIncrementNum(bX - uy * head * 0.4) & " " & FEMIncrementNum(bY + ux * head * 0.4) & " L " & FEMIncrementNum(bX + uy * head * 0.4) & " " & FEMIncrementNum(bY - ux * head * 0.4) & " Z' fill='" & col & "'/>"
    Next sign
    FEMIncrementAxis = fragment
End Function

Private Function FEMIncrementEdgeLength(ByVal a As Variant, ByVal i As Long, ByVal j As Long) As Double
    Dim x As Double, y As Double
    x = a(j * 2) - a(i * 2): y = a(j * 2 + 1) - a(i * 2 + 1)
    FEMIncrementEdgeLength = Sqr(x * x + y * y)
End Function

Private Function FEMIncrementNum(ByVal x As Double, Optional ByVal keyMode As Boolean = False) As String
    Dim s As String
    If keyMode Then s = Format$(x, "0.000000") Else s = Format$(x, "0.######")
    s = Replace$(s, ",", ".")
    FEMIncrementNum = s
End Function
Private Function FEMIncrementXml(ByVal s As String) As String
    FEMIncrementXml = Replace$(Replace$(s, "&", "&amp;"), "'", "&apos;")
End Function

Private Function FEMIncrementSci(ByVal x As Double) As String
    FEMIncrementSci = Replace$(Format$(x, "0.000E+00"), Application.International(xlDecimalSeparator), ".")
End Function
