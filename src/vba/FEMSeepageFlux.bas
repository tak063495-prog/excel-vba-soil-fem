Option Explicit
' Equilibrated face-integrated flows. GP Darcy flux remains explicitly RAW_CG.
Public Sub GWRecoverFaceFlows(ByRef conductivity() As Double, ByRef heads() As Double, ByRef fixedHead() As Boolean, ByRef faceFlow() As Double, ByRef localError As Double, ByRef correctionRatio As Double)
  Dim owners As Object, k As String, owner As Variant, e As Long, s As Long, n As Long, a As Long, b As Long, i As Long
  Dim neighbor() As Long, adjustable() As Boolean, diagonal() As Double, residual() As Double
  Dim lambda() As Double, direction() As Double, product() As Double, z() As Double
  Dim original() As Double, rho As Double, rhoNew As Double, den As Double, alpha As Double
  Dim value As Double, raw As Double, balance As Double, flowScale As Double, norm0 As Double, localDefect As Double, corr As Double
  n = NumberOfElement
  ReDim faceFlow(0 To 3, 0 To n - 1): ReDim neighbor(0 To 3, 0 To n - 1)
  ReDim adjustable(0 To 3, 0 To n - 1): ReDim original(0 To 3, 0 To n - 1)
  ReDim diagonal(n - 1): ReDim residual(n - 1): ReDim lambda(n - 1)
  ReDim direction(n - 1): ReDim product(n - 1): ReDim z(n - 1)
  Set owners = CreateObject("Scripting.Dictionary")
  For e = 0 To n - 1
    For s = 0 To 3
      neighbor(s, e) = -1
      If conductivity(e) > 0# Then
        a = GWFlowNode(e, s): b = GWFlowNode(e, (s + 1) Mod 4)
        If a < b Then k = a & ":" & b Else k = b & ":" & a
        k = k & ":" & GWFlowNode(e, s + 4)
        raw = GWRawFaceFlow(e, s, conductivity(e), heads)
        If owners.Exists(k) Then
          owner = owners(k)
          neighbor(s, e) = CLng(owner(0)): neighbor(CLng(owner(1)), CLng(owner(0))) = e
          value = 0.5 * (raw - faceFlow(CLng(owner(1)), CLng(owner(0))))
          faceFlow(s, e) = value: faceFlow(CLng(owner(1)), CLng(owner(0))) = -value
          adjustable(s, e) = True: adjustable(CLng(owner(1)), CLng(owner(0))) = True
        Else
          owners.Add k, Array(e, s)
          faceFlow(s, e) = raw
        End If
      End If
    Next s
  Next e
  For e = 0 To n - 1
    If conductivity(e) > 0# Then
      For s = 0 To 3
        If neighbor(s, e) < 0 Then
          adjustable(s, e) = fixedHead(GWFlowNode(e, s)) And fixedHead(GWFlowNode(e, (s + 1) Mod 4)) And fixedHead(GWFlowNode(e, s + 4))
          If Not adjustable(s, e) Then faceFlow(s, e) = 0#
        End If
        If adjustable(s, e) Then diagonal(e) = diagonal(e) + 1#
        original(s, e) = faceFlow(s, e)
        residual(e) = residual(e) - faceFlow(s, e)
      Next s
      If diagonal(e) <= 0# Then Err.Raise vbObjectError + 7960, , "保存流束を再構成できる境界/接続がありません。"
      z(e) = residual(e) / diagonal(e): direction(e) = z(e)
    End If
  Next e
  rho = GWFluxDot(residual, z): norm0 = Sqr(GWFluxDot(residual, residual))
  If norm0 > 0# Then
    For i = 1 To 4 * n + 20
      For e = 0 To n - 1
        product(e) = diagonal(e) * direction(e)
        For s = 0 To 3
          If neighbor(s, e) >= 0 Then product(e) = product(e) - direction(neighbor(s, e))
        Next s
      Next e
      den = GWFluxDot(direction, product)
      If den <= 0# Then Err.Raise vbObjectError + 7960, , "保存流束再構成の系が特異です。"
      alpha = rho / den
      For e = 0 To n - 1
        lambda(e) = lambda(e) + alpha * direction(e)
        residual(e) = residual(e) - alpha * product(e)
      Next e
      If Sqr(GWFluxDot(residual, residual)) <= norm0 * 0.000000000001 Then Exit For
      For e = 0 To n - 1
        If diagonal(e) > 0# Then z(e) = residual(e) / diagonal(e)
      Next e
      rhoNew = GWFluxDot(residual, z)
      For e = 0 To n - 1
        direction(e) = z(e) + rhoNew / rho * direction(e)
      Next e
      rho = rhoNew
    Next i
  End If
  localError = 0#: correctionRatio = 0#: raw = 0#: corr = 0#
  For e = 0 To n - 1
    If conductivity(e) > 0# Then
      balance = 0#: flowScale = 0#
      For s = 0 To 3
        value = 0#
        If adjustable(s, e) Then
          value = lambda(e)
          If neighbor(s, e) >= 0 Then value = value - lambda(neighbor(s, e))
        End If
        faceFlow(s, e) = original(s, e) + value
        balance = balance + faceFlow(s, e): flowScale = flowScale + Abs(faceFlow(s, e))
        corr = corr + value * value: raw = raw + original(s, e) * original(s, e)
      Next s
      ' Conductivity * thickness is a documented unit-gradient flow flowScale.
      flowScale = flowScale + conductivity(e) * Material(Elem(e).MatNo).thickness
      localDefect = Abs(balance) / flowScale
      If localDefect > localError Then localError = localDefect
    End If
  Next e
  If raw > 0# Then correctionRatio = Sqr(corr / raw)
  If localError > 0.00000001 Then Err.Raise vbObjectError + 7960, , "再構成流束の局所水収支が許容値を超えています。"
End Sub

Private Function GWRawFaceFlow(ByVal e As Long, ByVal side As Long, ByVal conductivity As Double, ByRef heads() As Double) As Double
  Dim q As Long, j As Long, xi As Double, eta As Double, t As Double, dx As Double, dy As Double
  Dim j11 As Double, j12 As Double, j21 As Double, j22 As Double, det As Double, hx As Double, hy As Double
  Dim ns(0 To 7) As Double, dxi(0 To 7) As Double, deta(0 To 7) As Double, gp As Variant, weights As Variant
  gp = Array(-Sqr(0.6), 0#, Sqr(0.6)): weights = Array(5# / 9#, 8# / 9#, 5# / 9#)
  For q = 0 To 2
    t = gp(q)
    Select Case side
      Case 0: xi = t: eta = -1#
      Case 1: xi = 1#: eta = t
      Case 2: xi = -t: eta = 1#
      Case 3: xi = -1#: eta = -t
    End Select
    P6Q8Shape xi, eta, ns, dxi, deta
    j11 = 0#: j12 = 0#: j21 = 0#: j22 = 0#: hx = 0#: hy = 0#
    For j = 0 To 7
      j11 = j11 + dxi(j) * Elem(e).x(j): j12 = j12 + deta(j) * Elem(e).x(j)
      j21 = j21 + dxi(j) * Elem(e).y(j): j22 = j22 + deta(j) * Elem(e).y(j)
    Next j
    det = j11 * j22 - j12 * j21
    If det <= 0# Then Err.Raise vbObjectError + 7960, , "保存流束の境界Jacobianが不正です。"
    For j = 0 To 7
      hx = hx + (j22 * dxi(j) - j21 * deta(j)) / det * heads(GWFlowNode(e, j))
      hy = hy + (-j12 * dxi(j) + j11 * deta(j)) / det * heads(GWFlowNode(e, j))
    Next j
    Select Case side
      Case 0: dx = j11: dy = j21
      Case 1: dx = j12: dy = j22
      Case 2: dx = -j11: dy = -j21
      Case 3: dx = -j12: dy = -j22
    End Select
    GWRawFaceFlow = GWRawFaceFlow + conductivity * (-hx * dy + hy * dx) * Material(Elem(e).MatNo).thickness * weights(q)
  Next q
End Function

Private Function GWFluxDot(ByRef a() As Double, ByRef b() As Double) As Double
  Dim i As Long
  For i = 0 To UBound(a)
    GWFluxDot = GWFluxDot + a(i) * b(i)
  Next i
End Function
