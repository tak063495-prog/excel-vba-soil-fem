Option Explicit

Public Function TIncrementMath() As String
    On Error GoTo Failed
    Dim a As Double, b As Double, t As Double, m As Double
    Dim pi As Double, exr As Double, eyr As Double, gxr As Double
    pi = 4 * Atn(1)
    FEMStrainPrincipal2D 0.02, 0, 0, a, b, t
    TIncAssert Abs(a - 0.02) < 1E-12 And Abs(b) < 1E-12 And Abs(t) < 1E-12, "uniaxial"
    TIncAssert Abs(FEMStrainMaxShear(0.02, 0, 0) - 0.02) < 1E-12, "uniaxial shear"
    FEMStrainPrincipal2D 0, 0, 0.04, a, b, t
    TIncAssert Abs(a - 0.02) < 1E-12 And Abs(b + 0.02) < 1E-12 And Abs(t - pi / 4) < 1E-12, "pure shear"
    FEMStrainPrincipal2D 0, 0, -0.04, a, b, t
    TIncAssert Abs(t - 3 * pi / 4) < 1E-12, "negative shear"
    TIncAssert Abs(FEMStrainMaxShear(0.02, 0.02, 0, 0) - 0.02) < 1E-12, "biaxial"
    FEMStrainPrincipal2D -0.01, -0.03, 0, a, b, t
    TIncAssert Abs(a + 0.01) < 1E-12 And Abs(b + 0.03) < 1E-12, "compression"
    TIncAssert Abs(FEMStrainMaxShear(-0.01, -0.03, 0, 0) - 0.03) < 1E-12, "compression 3D"
    FEMStrainPrincipal2D 0.01, 0.03, 0.02, a, b, t
    TIncAssert Abs(a + b - 0.04) < 1E-12, "rotated trace"
    exr = a * Cos(t) ^ 2 + b * Sin(t) ^ 2
    eyr = a * Sin(t) ^ 2 + b * Cos(t) ^ 2
    gxr = 2 * (a - b) * Sin(t) * Cos(t)
    TIncAssert Abs(exr - 0.01) < 1E-12 And Abs(eyr - 0.03) < 1E-12 And Abs(gxr - 0.02) < 1E-12, "rotated reconstruction"
    FEMStrainPrincipal2D 1E-200, -1E-200, 1E-200, a, b, t
    TIncAssert a >= b And t >= 0 And t < pi, "tiny"
    TIncAssert Abs((a - b) - Sqr(5) * 1E-200) < 1E-212, "tiny magnitude"
    FEMStrainPrincipal2D 1E-300, 1E100, 1E-300, a, b, t
    TIncAssert a >= b And t >= 0 And t < pi, "mixed scale"
    TIncrementMath = "PASS"
    Exit Function
Failed:
    TIncrementMath = "FAIL: " & Err.Description
End Function

Public Function TIncrementSvgCases(ByVal kind As String) As String
    Dim d As Object, before As String, legend As String
    Set d = TIncCells(kind)
    before = TIncCellSignature(d)
    TIncrementSvgCases = FEMIncrementOverlaySvg(d, True, True, (kind = "constant" Or kind = "linear"), 1, legend)
    If before <> TIncCellSignature(d) Then Err.Raise vbObjectError + 4911, "IncrementDisplayHarness", "入力cellsが変更されました。"
    If InStr(1, TIncrementSvgCases, "increment-overlay", vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 4912, "IncrementDisplayHarness", "increment overlay layer がありません。"
    If kind = "axes" Then
        If InStr(1, TIncrementSvgCases, "plastic-principal-crosses", vbTextCompare) = 0 Then _
            Err.Raise vbObjectError + 4913, "IncrementDisplayHarness", "principal cross layer がありません。"
        If InStr(1, TIncrementSvgCases, "data-principal", vbTextCompare) = 0 Then _
            Err.Raise vbObjectError + 4914, "IncrementDisplayHarness", "principal cross data tag がありません。"
    End If
End Function

Public Function TIncrementSvgFixture(ByVal kind As String) As String
    TIncrementSvgFixture = TIncrementSvgCases(kind)
End Function

Public Function TOptionError(ByVal which As String) As String
    On Error GoTo Failed
    If which = "mode" Then TOptionError = CStr(FEMOverlayModeA()): Exit Function
    If which = "scale" Then TOptionError = CStr(FEMPlasticCrossScale()): Exit Function
    TOptionError = "NO_ERROR"
    Exit Function
Failed:
    TOptionError = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function TIncCells(ByVal kind As String) As Object
    Dim d As Object, x As Long, y As Long, id As Long, gx As Double
    Set d = CreateObject("Scripting.Dictionary")
    For y = 0 To 1
        For x = 0 To 1
            id = id + 1: gx = 0
            If kind = "constant" Then gx = 0.01
            If kind = "linear" Then gx = (x + y + 1) * 0.01
            If kind = "shear" Or kind = "negative_shear" Then gx = 0.04
            If kind = "axes" Then d.Add CStr(id), Array(x, y, x + 1, y, x + 1, y + 1, x, y + 1, 0, 0.02, -0.01, 0) Else d.Add CStr(id), Array(x, y, x + 1, y, x + 1, y + 1, x, y + 1, gx, 0, 0, IIf(kind = "shear", 0.04, IIf(kind = "negative_shear", -0.04, 0)))
        Next x
    Next y
    If kind = "thin" Then
        d("1") = Array(0, 0, 0.001, 0, 0.001, 1, 0, 1, 0, 0.02, -0.01, 0)
        d("2") = Array(1, 0, 2, 0, 2, 1, 1, 1, 0, 0.02, -0.01, 0)
        d("3") = Array(0, 1, 1, 1, 1, 2, 0, 2, 0, 0.02, -0.01, 0)
        d("4") = Array(1, 1, 2, 1, 2, 2, 1, 2, 0, 0.02, -0.01, 0)
    End If
    Set TIncCells = d
End Function

Private Function TIncCellSignature(ByVal d As Object) As String
    Dim k As Variant, a As Variant, s As String
    For Each k In d.Keys
        a = d(k): s = s & CStr(k) & ":" & CStr(a(8)) & ":" & CStr(a(9)) & ":" & CStr(a(10)) & ":" & CStr(a(11)) & ";"
    Next k
    TIncCellSignature = s
End Function

Private Sub TIncAssert(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise vbObjectError + 4910, "IncrementDisplayHarness", label & " failed"
End Sub
