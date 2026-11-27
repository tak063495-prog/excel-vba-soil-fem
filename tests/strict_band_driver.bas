' Append to FEMSolver in a TEST COPY, so the original private kernels are exercised.
Public Function QA_KernelEquality() As Variant
    Dim sizes As Variant, spec As Variant, mode As Long, n As Long, bw As Long
    Dim band() As Double, pV() As Double, pC() As Double, off() As Long
    Dim ak() As Double, ai() As Double, rhsV() As Double, rhsC() As Double, xV() As Double, xC() As Double
    Dim i As Long, j As Long, which As Long, nv As Boolean, nc As Long, fv As Long, fc As Long
    Dim ok As Boolean, factorEqual As Boolean, rhsEqual As Boolean, xEqual As Boolean, count As Long
    Dim tol As Double, results() As Variant, caseCount As Long, row As Long
    sizes = Array(Array(1, 0), Array(1, 2), Array(2, 1), Array(17, 0), Array(17, 1), Array(17, 3), Array(17, 16), Array(80, 19), Array(128, 31), Array(512, 80))
    ReDim results(0 To 69, 0 To 8)
    tol = 0.000000000001
    For Each spec In sizes
        n = spec(0): bw = spec(1): count = n * (bw + 1)
        For mode = 0 To 6
            ReDim band(0 To n + 2, 0 To bw)
            For i = 0 To n - 1
                band(i, 0) = 10# + bw
                If mode = 1 And i Mod 2 = 0 Then band(i, 0) = -band(i, 0)
                For j = 1 To bw
                    If i + j < n Then band(i, j) = Sin(13# * i + 7# * j + n) * 0.025
                Next j
            Next i
            If mode = 2 Or mode = 3 Or mode = 4 Then
                For i = 0 To n - 1
                    For j = 1 To bw: band(i, j) = 0#: Next j
                Next i
                which = 0
                If mode = 3 Then which = n - 1
                If mode = 4 Then which = n \ 2
                band(which, 0) = 0#
                If mode = 4 Then band(which, 0) = tol
            End If
            ReDim pV(0 To count - 1): ReDim pC(0 To count - 1)
            ReDim off(0 To n - 1): ReDim ak(0 To bw): ReDim ai(0 To bw)
            P6BandBuildRowOff off, n, bw + 1
            P6PackBand band, pV, off, n, bw
            nv = P6BandFactor1D(pV, off, ak, ai, n, bw, tol)
            fv = P6BandFactorFailK
            nc = FEMBandFastFactor(band, pC, ak, n, bw, tol, fc)
            factorEqual = (QA_MemCmp(VarPtr(pV(0)), VarPtr(pC(0)), CLngPtr(count) * 8) = 0)
            ok = (nv = (nc = 0)) And fv = fc And factorEqual
            rhsEqual = True: xEqual = True
            If nv Then
                If mode = 5 Then pV(0) = 0#: pC(0) = 0#
                If mode = 6 Then pV(off(n - 1)) = 0#: pC(off(n - 1)) = 0#
                ReDim rhsV(0 To n - 1): ReDim rhsC(0 To n - 1)
                ReDim xV(0 To n - 1): ReDim xC(0 To n - 1)
                For i = 0 To n - 1
                    rhsV(i) = Cos(5# * i + n) + 0.1: rhsC(i) = rhsV(i)
                    xV(i) = -i: xC(i) = xV(i)
                Next i
                nv = P6BandSolve1D(pV, off, rhsV, xV, n, bw, tol): fv = P6BandSolveFailK
                nc = FEMBandFastSolve(pC, rhsC, xC, n, bw, tol, fc)
                rhsEqual = (QA_MemCmp(VarPtr(rhsV(0)), VarPtr(rhsC(0)), CLngPtr(n) * 8) = 0)
                xEqual = (QA_MemCmp(VarPtr(xV(0)), VarPtr(xC(0)), CLngPtr(n) * 8) = 0)
                ok = ok And (nv = (nc = 0)) And fv = fc And rhsEqual And xEqual
            End If
            results(row, 0) = n: results(row, 1) = bw: results(row, 2) = mode
            results(row, 3) = ok: results(row, 4) = factorEqual
            results(row, 5) = rhsEqual: results(row, 6) = xEqual
            results(row, 7) = fv: results(row, 8) = fc
            row = row + 1
        Next mode
    Next spec
    QA_KernelEquality = results
End Function

Public QA_ActualCompared As Boolean
Public QA_ActualRecord As Variant
Private Sub QA_CompareActualBand(ByVal n As Long, ByVal bw As Long, ByVal tol As Double)
    Dim pv() As Double, off() As Long, ak() As Double, ai() As Double
    Dim rv() As Double, rc() As Double, xv() As Double, xc() As Double
    Dim i As Long, fv As Long, fc As Long, ok As Boolean, nv As Boolean, nc As Long
    Dim packedEqual As Boolean, rhsEqual As Boolean, xEqual As Boolean, started As Double, elapsed As Double
    ReDim pv(0 To n * (bw + 1) - 1): ReDim off(0 To n - 1)
    ReDim ak(0 To bw): ReDim ai(0 To bw)
    P6BandBuildRowOff off, n, bw + 1
    P6PackBand P6FactoredBand, pv, off, n, bw
    started = Timer
    nv = P6BandFactor1D(pv, off, ak, ai, n, bw, tol): fv = P6BandFactorFailK
    elapsed = P6ElapsedMs(started) / 1000#
    packedEqual = (QA_MemCmp(VarPtr(pv(0)), VarPtr(P6Packed(0)), CLngPtr(n) * (bw + 1) * 8) = 0)
    rv = force: rc = force
    ReDim xv(0 To n - 1): ReDim xc(0 To n - 1)
    If nv Then
        nv = P6BandSolve1D(pv, off, rv, xv, n, bw, tol): fv = P6BandSolveFailK
        nc = FEMBandFastSolve(P6Packed, rc, xc, n, bw, tol, fc)
        rhsEqual = (QA_MemCmp(VarPtr(rv(0)), VarPtr(rc(0)), CLngPtr(n) * 8) = 0)
        xEqual = (QA_MemCmp(VarPtr(xv(0)), VarPtr(xc(0)), CLngPtr(n) * 8) = 0)
        ok = (nv = (nc = 0)) And fv = fc And packedEqual And rhsEqual And xEqual
    End If
    QA_ActualRecord = Array(ok, n, bw, packedEqual, rhsEqual, xEqual, fv, fc, elapsed)
    QA_ActualCompared = True
    If ok Then Err.Raise vbObjectError + 7300, "QA", "QA_ACTUAL_MATRIX_BITWISE_PASS"
    Err.Raise vbObjectError + 7301, "QA", "QA_ACTUAL_MATRIX_MISMATCH"
End Sub
