Option Explicit

' Array-only L-BFGS. Boundary/material/load locks are supplied by MeshAdapt.
Public Function MeshAngleOptimize(ByRef x() As Double, ByRef y() As Double, ByRef cn() As Long, ByVal nn As Long, ByVal ne As Long, ByRef fixed() As Boolean, ByVal limit As Long, ByRef iterations As Long) As Boolean
    MeshPerfEnter "MeshAngle.MeshAngleOptimize"
    Dim map() As Long, nodes() As Long, count As Long, i As Long, e As Long, j As Long, n As Long, nv As Long
    Dim xx() As Double, yy() As Double, g() As Double, ng() As Double, vec() As Double, trial() As Double, z() As Double
    Dim hs() As Double, hy() As Double, rho(1 To 12) As Double, alpha(1 To 12) As Double
    Dim slot As Long, head As Long, h As Long, hcount As Long, f As Double, nf As Double, dot As Double, beta As Double, gamma As Double, slope As Double, stepSize As Double
    Dim ys As Double, ysq As Double, gnorm As Double, Accepted As Boolean, used() As Boolean, initialArea() As Double
    Dim workGX() As Double, workGY() As Double, workArea() As Double, workGA() As Double
    Dim edge As Object, key As String, ad1() As Long, ad2() As Long, nad As Long, a As Long, b As Long, relative As Double
    ReDim used(1 To nn): ReDim map(1 To nn): ReDim nodes(1 To nn)
    ReDim initialArea(1 To ne): Set edge = CreateObject("Scripting.Dictionary")
    ReDim ad1(1 To ne * 4): ReDim ad2(1 To ne * 4)
    For e = 1 To ne
        For j = 1 To 4
            a = cn(e, j): b = cn(e, j Mod 4 + 1): used(a) = True
            initialArea(e) = initialArea(e) + x(a) * y(b) - x(b) * y(a)
            If a < b Then key = CStr(a) & ":" & CStr(b) Else key = CStr(b) & ":" & CStr(a)
            If edge.Exists(key) Then
                nad = nad + 1: ad1(nad) = CLng(edge(key)): ad2(nad) = e
            Else
                edge.Add key, e
            End If
        Next j
        initialArea(e) = initialArea(e) / 2#
    Next e
    For i = 1 To nn
        If used(i) And Not fixed(i) Then count = count + 1: map(i) = count: nodes(count) = i
    Next i
    MeshAngleOptimize = True: iterations = 0
    If count = 0 Then MeshPerfLeave "MeshAngle.MeshAngleOptimize": Exit Function
    nv = count * 2: ReDim vec(1 To nv): ReDim trial(1 To nv): ReDim z(1 To nv)
    ReDim hs(1 To nv, 1 To 12): ReDim hy(1 To nv, 1 To 12)
    For i = 1 To count: vec(2 * i - 1) = x(nodes(i)): vec(2 * i) = y(nodes(i)): Next i
    ReDim xx(1 To nn): ReDim yy(1 To nn)
    For i = 1 To nn: xx(i) = x(i): yy(i) = y(i): Next i
    ReDim g(1 To nv): ReDim ng(1 To nv)
    ReDim workGX(1 To nn): ReDim workGY(1 To nn): ReDim workArea(1 To ne): ReDim workGA(1 To ne)
    f = AngleFG(vec, xx, yy, cn, ne, nodes, count, map, initialArea, ad1, ad2, nad, g, workGX, workGY, workArea, workGA)
    For iterations = 1 To limit
        gnorm = 0#: For i = 1 To nv: z(i) = g(i): gnorm = gnorm + g(i) * g(i): Next i
        If gnorm < 0.00000000000001 Then Exit For
        For h = hcount To 1 Step -1
            slot = (head + h - 1) Mod 12 + 1
            dot = 0#: For i = 1 To nv: dot = dot + hs(i, slot) * z(i): Next i
            alpha(slot) = rho(slot) * dot
            For i = 1 To nv: z(i) = z(i) - alpha(slot) * hy(i, slot): Next i
        Next h
        gamma = 1#
        If hcount > 0 Then
            slot = (head + hcount - 1) Mod 12 + 1
            ys = 0#: ysq = 0#
            For i = 1 To nv: ys = ys + hs(i, slot) * hy(i, slot): ysq = ysq + hy(i, slot) ^ 2: Next i
            If ysq > 1E-30 Then gamma = ys / ysq
        End If
        For i = 1 To nv: z(i) = z(i) * gamma: Next i
        For h = 1 To hcount
            slot = (head + h - 1) Mod 12 + 1
            dot = 0#: For i = 1 To nv: dot = dot + hy(i, slot) * z(i): Next i
            beta = rho(slot) * dot
            For i = 1 To nv: z(i) = z(i) + hs(i, slot) * (alpha(slot) - beta): Next i
        Next h
        slope = 0#: For i = 1 To nv: z(i) = -z(i): slope = slope + g(i) * z(i): Next i
        If slope >= 0# Then
            hcount = 0: head = 0: slope = -gnorm
            For i = 1 To nv: z(i) = -g(i): Next i
        End If
        stepSize = 1#: If hcount = 0 And gnorm > 1# Then stepSize = 1# / Sqr(gnorm)
        Accepted = False
        For j = 1 To 40
            For i = 1 To nv: trial(i) = vec(i) + stepSize * z(i): Next i
            nf = AngleFG(trial, xx, yy, cn, ne, nodes, count, map, initialArea, ad1, ad2, nad, ng, workGX, workGY, workArea, workGA)
            If nf < 1E+90 And nf <= f + 0.0001 * stepSize * slope Then Accepted = True: Exit For
            stepSize = stepSize / 2#
        Next j
        If Not Accepted Then Exit For
        ys = 0#: For i = 1 To nv: ys = ys + (trial(i) - vec(i)) * (ng(i) - g(i)): Next i
        If ys > 1E-20 Then
            If hcount = 12 Then
                head = (head + 1) Mod 12
            Else
                hcount = hcount + 1
            End If
            slot = (head + hcount - 1) Mod 12 + 1: rho(slot) = 1# / ys
            For i = 1 To nv: hs(i, slot) = trial(i) - vec(i): hy(i, slot) = ng(i) - g(i): Next i
        End If
        relative = Abs(nf - f) / (1# + Abs(f)): vec = trial: g = ng: f = nf
        If relative < 0.000000000001 Then Exit For
    Next iterations
    If iterations > limit Then iterations = limit
    For i = 1 To count: x(nodes(i)) = vec(2 * i - 1): y(nodes(i)) = vec(2 * i): Next i
    MeshPerfLeave "MeshAngle.MeshAngleOptimize"
End Function
Private Function AngleFG(ByRef v() As Double, ByRef x() As Double, ByRef y() As Double, ByRef cn() As Long, ByVal ne As Long, ByRef nodes() As Long, ByVal count As Long, ByRef map() As Long, ByRef baseArea() As Double, ByRef ad1() As Long, ByRef ad2() As Long, ByVal nad As Long, ByRef g() As Double, ByRef gx() As Double, ByRef gy() As Double, ByRef ar() As Double, ByRef ga() As Double) As Double
    MeshPerfEnter "MeshAngle.AngleFG"
    Dim i As Long, j As Long, e As Long, a As Long, b As Long, c As Long, k As Long
    Dim ux As Double, uy As Double, vx As Double, vy As Double, lu As Double, lv As Double, cs As Double, cr As Double
    Dim penalty As Double, dc As Double, gux As Double, guy As Double, gvx As Double, gvy As Double, lr As Double, low As Double, dj As Double
    Dim ratio As Double, value As Double, tol As Double
    For i = 1 To UBound(gx): gx(i) = 0#: gy(i) = 0#: Next i
    For i = 1 To ne: ar(i) = 0#: ga(i) = 0#: Next i
    For i = 1 To count: x(nodes(i)) = v(2 * i - 1): y(nodes(i)) = v(2 * i): Next i
    tol = 0.694658370458997 ' cos(46 degrees)
    For e = 1 To ne
        For j = 1 To 4
            a = cn(e, j): b = cn(e, j Mod 4 + 1): c = cn(e, (j + 2) Mod 4 + 1)
            ux = x(b) - x(a): uy = y(b) - y(a): vx = x(c) - x(a): vy = y(c) - y(a)
            cr = ux * vy - uy * vx
            If cr <= baseArea(e) * 0.0000000001 Then AngleFG = 1E+100: MeshPerfLeave "MeshAngle.AngleFG": Exit Function
            lu = Sqr(ux * ux + uy * uy): lv = Sqr(vx * vx + vy * vy)
            cs = (ux * vx + uy * vy) / (lu * lv): penalty = Abs(cs) - tol: If penalty < 0# Then penalty = 0#
            value = value + 100# * penalty * penalty + 0.01 * cs * cs
            dc = 200# * penalty * Sgn(cs) + 0.02 * cs
            gux = dc * (vx / (lu * lv) - cs * ux / (lu * lu)): guy = dc * (vy / (lu * lv) - cs * uy / (lu * lu))
            gvx = dc * (ux / (lu * lv) - cs * vx / (lv * lv)): gvy = dc * (uy / (lu * lv) - cs * vy / (lv * lv))
            lr = Log(lu / lv): value = value + 0.01 * lr * lr
            gux = gux + 0.02 * lr * ux / (lu * lu): guy = guy + 0.02 * lr * uy / (lu * lu)
            gvx = gvx - 0.02 * lr * vx / (lv * lv): gvy = gvy - 0.02 * lr * vy / (lv * lv)
            low = 0.08 - cr / baseArea(e): If low < 0# Then low = 0#
            value = value + 1000# * low * low: dj = -2000# * low / baseArea(e)
            gux = gux + dj * vy: guy = guy - dj * vx: gvx = gvx - dj * uy: gvy = gvy + dj * ux
            gx(a) = gx(a) - gux - gvx: gy(a) = gy(a) - guy - gvy
            gx(b) = gx(b) + gux: gy(b) = gy(b) + guy: gx(c) = gx(c) + gvx: gy(c) = gy(c) + gvy
            ar(e) = ar(e) + x(a) * y(b) - x(b) * y(a)
        Next j
        ar(e) = ar(e) / 2#
    Next e
    For i = 1 To nad
        a = ad1(i): b = ad2(i): ratio = Log(ar(a) / ar(b)): value = value + 0.02 * ratio * ratio
        ga(a) = ga(a) + 0.04 * ratio / ar(a): ga(b) = ga(b) - 0.04 * ratio / ar(b)
    Next i
    For e = 1 To ne
        For j = 1 To 4
            a = cn(e, j): b = cn(e, j Mod 4 + 1): c = cn(e, (j + 2) Mod 4 + 1)
            gx(a) = gx(a) + 0.5 * ga(e) * (y(b) - y(c)): gy(a) = gy(a) + 0.5 * ga(e) * (x(c) - x(b))
        Next j
    Next e
    For i = 1 To count: g(2 * i - 1) = gx(nodes(i)): g(2 * i) = gy(nodes(i)): Next i
    AngleFG = value
    MeshPerfLeave "MeshAngle.AngleFG"
End Function