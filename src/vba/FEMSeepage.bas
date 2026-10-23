Option Explicit

Public Sub GWSeepSolve(ByRef conductivity() As Double, ByRef fixedHead() As Boolean, ByRef prescribedHead() As Double, _
 ByRef heads() As Double, ByRef fluxX() As Double, ByRef fluxY() As Double, ByRef reactionFlow() As Double, _
 ByRef relativeResidual As Double, ByRef balanceError As Double)
  Dim locked() As Boolean
  Dim ne As Long, nn As Long, e As Long, i As Long, j As Long, g As Long, it As Long
  Dim lm() As Double, rhs() As Double, rr() As Double, zz() As Double
  Dim pp() As Double, aa() As Double, diag() As Double
  Dim parent() As Long, activeNode() As Boolean, hasBC() As Boolean
  Dim q As Long, ni As Long, nj As Long, Root As Long, col As Long
  Dim rho As Double, rhoNew As Double, denom As Double, alpha As Double, beta As Double
  Dim normR As Double, normB As Double, value As Double
  Dim gp(0 To 2) As Double, gw(0 To 2) As Double, Nshape(0 To 7) As Double
  Dim dxi(0 To 7) As Double, deta(0 To 7) As Double
  Dim dNx(0 To 7) As Double, dNy(0 To 7) As Double
  Dim j11 As Double, j12 As Double, j21 As Double, j22 As Double, detJ As Double
  Dim thick As Double, kValue As Double, gx As Double, gy As Double
  Dim sumReaction As Double, sumAbsReaction As Double, foundBC As Boolean
  On Error GoTo Failed
  ne = NumberOfElement
  nn = NumberOfFreeNode
  If ne < 1 Or nn < 1 Then GWFail "No seepage mesh."
  If UBound(conductivity) < ne - 1 Then GWFail "Conductivity array is too small."
  If UBound(fixedHead) < nn - 1 Then GWFail "Fixed head array is too small."
  If UBound(prescribedHead) < nn - 1 Then GWFail "Prescribed head array is too small."
  ReDim heads(0 To nn - 1)
  ReDim reactionFlow(0 To nn - 1)
  ReDim fluxX(0 To 3, 0 To ne - 1)
  ReDim fluxY(0 To 3, 0 To ne - 1)
  ReDim lm(0 To 7, 0 To 7, 0 To ne - 1)
  ReDim rhs(0 To nn - 1)
  ReDim rr(0 To nn - 1)
  ReDim zz(0 To nn - 1)
  ReDim pp(0 To nn - 1)
  ReDim aa(0 To nn - 1)
  ReDim diag(0 To nn - 1)
  ReDim parent(0 To nn - 1)
  ReDim activeNode(0 To nn - 1)
  ReDim hasBC(0 To nn - 1)
  For i = 0 To nn - 1
    parent(i) = i
    If fixedHead(i) Then
      hasBC(i) = True
      If Not P2IsFinite(prescribedHead(i)) Then GWFail "Nonfinite prescribed head."
    End If
  Next i
  gp(0) = -Sqr(0.6)
  gp(1) = 0#
  gp(2) = Sqr(0.6)
  gw(0) = 5# / 9#
  gw(1) = 8# / 9#
  gw(2) = 5# / 9#
  For e = 0 To ne - 1
    kValue = conductivity(e)
    If Not P2IsFinite(kValue) Or kValue < 0# Then GWFail "Invalid conductivity."
    If kValue <= 0# Then GoTo NextElement
    For i = 0 To 7
      ni = GWFlowNode(e, i)
      If ni < 0 Or ni >= nn Then GWFail "Element node out of range."
      activeNode(ni) = True
    Next i
    thick = Material(Elem(e).MatNo).thickness
    If thick <= 0# Then thick = 1#
    For g = 0 To 8
      q = g Mod 3
      Root = g \ 3
      FEMShapeForElement e, gp(q), gp(Root), Nshape, dxi, deta
      j11 = 0#: j12 = 0#: j21 = 0#: j22 = 0#
      For i = 0 To 7
        j11 = j11 + dxi(i) * Elem(e).x(i)
        j12 = j12 + deta(i) * Elem(e).x(i)
        j21 = j21 + dxi(i) * Elem(e).y(i)
        j22 = j22 + deta(i) * Elem(e).y(i)
      Next i
      detJ = j11 * j22 - j12 * j21
      If detJ <= 1E-20 Then GWFail "Inverted or singular element geometry."
      For i = 0 To 7
        dNx(i) = (j22 * dxi(i) - j21 * deta(i)) / detJ
        dNy(i) = (-j12 * dxi(i) + j11 * deta(i)) / detJ
      Next i
      value = gw(q) * gw(Root) * detJ * thick * kValue
      For i = 0 To 7
        For j = 0 To 7
          lm(i, j, e) = lm(i, j, e) + value * (dNx(i) * dNx(j) + dNy(i) * dNy(j))
        Next j
      Next i
    Next g
    For i = 1 To 7
      GWUnion parent, GWFlowNode(e, 0), GWFlowNode(e, i)
    Next i
NextElement:
  Next e
  locked = fixedHead
  For i = 0 To nn - 1
    If Not activeNode(i) Then
      If fixedHead(i) Then GWFail "Prescribed head is not connected to a flow element."
      locked(i) = True
    End If
  Next i
  For i = 0 To nn - 1
    If activeNode(i) Then
      Root = GWFind(parent, i)
      foundBC = False
      For j = 0 To nn - 1
        If activeNode(j) Then
          If GWFind(parent, j) = Root Then
            If hasBC(j) Then
              foundBC = True
              Exit For
            End If
          End If
        End If
      Next j
      If Not foundBC Then GWFail "Flow component has no prescribed head."
    End If
  Next i
  For e = 0 To ne - 1
    If conductivity(e) > 0# Then
      For i = 0 To 7
        ni = GWFlowNode(e, i)
        For j = 0 To 7
          nj = GWFlowNode(e, j)
          If ni = nj Then diag(ni) = diag(ni) + lm(i, j, e)
          If fixedHead(nj) Then rhs(ni) = rhs(ni) - lm(i, j, e) * prescribedHead(nj)
        Next j
      Next i
    End If
  Next e
  For i = 0 To nn - 1
    If activeNode(i) And Not locked(i) Then
      If diag(i) <= 1E-30 Then GWFail "Singular seepage system."
    End If
    If locked(i) Then heads(i) = prescribedHead(i)
  Next i
  For i = 0 To nn - 1
    rr(i) = rhs(i)
    If Not locked(i) Then
      zz(i) = rr(i) / diag(i)
      pp(i) = zz(i)
    End If
  Next i
  rho = GWDot(rr, zz, locked, nn)
  normB = Sqr(GWDot(rhs, rhs, locked, nn))
  If normB < 1E-30 Then normB = 1#
  normR = Sqr(GWDot(rr, rr, locked, nn)) / normB
  If normR > 0.0000000001 Then
    For it = 1 To nn * 4 + 20
      GWMatVec lm, pp, aa, locked, conductivity, ne, nn
      denom = GWDot(pp, aa, locked, nn)
      If denom <= 0# Or Not P2IsFinite(denom) Then GWFail "Singular CG system."
      alpha = rho / denom
      For i = 0 To nn - 1
        If Not locked(i) Then
          heads(i) = heads(i) + alpha * pp(i)
          rr(i) = rr(i) - alpha * aa(i)
        End If
      Next i
      normR = Sqr(GWDot(rr, rr, locked, nn)) / normB
      If normR <= 0.0000000001 Then Exit For
      For i = 0 To nn - 1
        If Not locked(i) Then zz(i) = rr(i) / diag(i)
      Next i
      rhoNew = GWDot(rr, zz, locked, nn)
      beta = rhoNew / rho
      For i = 0 To nn - 1
        If Not locked(i) Then pp(i) = zz(i) + beta * pp(i)
      Next i
      rho = rhoNew
    Next it
  End If
  relativeResidual = normR
  If normR > 0.0000000001 Then GWFail "CG did not converge."
  GWComputeResidual lm, heads, rr, locked, conductivity, ne, nn
  relativeResidual = Sqr(GWDot(rr, rr, locked, nn)) / normB
  For i = 0 To nn - 1
    reactionFlow(i) = 0#
  Next i
  For e = 0 To ne - 1
    If conductivity(e) > 0# Then
      For i = 0 To 7
        ni = GWFlowNode(e, i)
        If locked(ni) Then
          For j = 0 To 7
            nj = GWFlowNode(e, j)
            reactionFlow(ni) = reactionFlow(ni) + lm(i, j, e) * heads(nj)
          Next j
        End If
      Next i
    End If
  Next e
  sumReaction = 0#: sumAbsReaction = 0#
  For i = 0 To nn - 1
    If fixedHead(i) Then
      sumReaction = sumReaction + reactionFlow(i)
      sumAbsReaction = sumAbsReaction + Abs(reactionFlow(i))
    End If
  Next i
  If sumAbsReaction < normB * 0.00000001 Then sumAbsReaction = normB
  If sumAbsReaction < 1E-30 Then sumAbsReaction = 1#
  balanceError = Abs(sumReaction) / sumAbsReaction
  GWComputeFlux conductivity, heads, fluxX, fluxY, ne
  Exit Sub
Failed:
  If Err.number = vbObjectError + 7901 Then Err.Raise Err.number, "GWSeepSolve", Err.description
  Err.Raise vbObjectError + 7901, "GWSeepSolve", Err.description
End Sub

Private Sub GWFail(ByVal messageText As String)
  Err.Raise vbObjectError + 7901, "GWSeepSolve", messageText
End Sub

Private Function GWDot(ByRef a() As Double, ByRef b() As Double, ByRef fixedHead() As Boolean, ByVal nn As Long) As Double
  Dim i As Long
  For i = 0 To nn - 1
    If Not fixedHead(i) Then GWDot = GWDot + a(i) * b(i)
  Next i
End Function

Private Sub GWMatVec(ByRef lm() As Double, ByRef x() As Double, ByRef y() As Double, ByRef fixedHead() As Boolean, ByRef conductivity() As Double, ByVal ne As Long, ByVal nn As Long)
  Dim e As Long, i As Long, j As Long, ni As Long, nj As Long
  For i = 0 To nn - 1
    y(i) = 0#
  Next i
  For e = 0 To ne - 1
    If conductivity(e) > 0# Then
      For i = 0 To 7
        ni = GWFlowNode(e, i)
        If Not fixedHead(ni) Then
          For j = 0 To 7
            nj = GWFlowNode(e, j)
            If Not fixedHead(nj) Then y(ni) = y(ni) + lm(i, j, e) * x(nj)
          Next j
        End If
      Next i
    End If
  Next e
End Sub

Private Sub GWComputeResidual(ByRef lm() As Double, ByRef h() As Double, ByRef rr() As Double, ByRef fixedHead() As Boolean, ByRef conductivity() As Double, ByVal ne As Long, ByVal nn As Long)
  Dim i As Long, j As Long, e As Long, ni As Long, nj As Long
  For i = 0 To nn - 1
    rr(i) = 0#
  Next i
  For e = 0 To ne - 1
    If conductivity(e) > 0# Then
      For i = 0 To 7
        ni = GWFlowNode(e, i)
        If Not fixedHead(ni) Then
          For j = 0 To 7
            nj = GWFlowNode(e, j)
            rr(ni) = rr(ni) + lm(i, j, e) * h(nj)
          Next j
        End If
      Next i
    End If
  Next e
  For i = 0 To nn - 1
    If Not fixedHead(i) Then rr(i) = rr(i) - 0#
  Next i
End Sub

Private Sub GWComputeFlux(ByRef conductivity() As Double, ByRef h() As Double, ByRef fx() As Double, ByRef fy() As Double, ByVal ne As Long)
  Dim gp As Double, e As Long, g As Long, i As Long, q As Long, r As Long
  Dim xis(0 To 3) As Double, etas(0 To 3) As Double
  Dim Nshape(0 To 7) As Double, dxi(0 To 7) As Double, deta(0 To 7) As Double
  Dim j11 As Double, j12 As Double, j21 As Double, j22 As Double, detJ As Double
  Dim dx As Double, dy As Double, gx As Double, gy As Double
  gp = 0.577350269189626
  xis(0) = -gp: etas(0) = -gp
  xis(1) = gp: etas(1) = -gp
  xis(2) = gp: etas(2) = gp
  xis(3) = -gp: etas(3) = gp
  For e = 0 To ne - 1
    If conductivity(e) > 0# Then
      For g = 0 To 3
        FEMShapeForElement e, xis(g), etas(g), Nshape, dxi, deta
        j11 = 0#: j12 = 0#: j21 = 0#: j22 = 0#
        For i = 0 To 7
          j11 = j11 + dxi(i) * Elem(e).x(i): j12 = j12 + deta(i) * Elem(e).x(i)
          j21 = j21 + dxi(i) * Elem(e).y(i): j22 = j22 + deta(i) * Elem(e).y(i)
        Next i
        detJ = j11 * j22 - j12 * j21
        gx = 0#: gy = 0#
        For i = 0 To 7
          dx = (j22 * dxi(i) - j21 * deta(i)) / detJ
          dy = (-j12 * dxi(i) + j11 * deta(i)) / detJ
          gx = gx + dx * h(GWFlowNode(e, i)): gy = gy + dy * h(GWFlowNode(e, i))
        Next i
        fx(g, e) = -conductivity(e) * gx: fy(g, e) = -conductivity(e) * gy
      Next g
    End If
  Next e
End Sub

Private Sub GWUnion(ByRef parent() As Long, ByVal a As Long, ByVal b As Long)
  Dim ra As Long, rb As Long
  ra = GWFind(parent, a): rb = GWFind(parent, b)
  If ra <> rb Then parent(rb) = ra
End Sub

Private Function GWFind(ByRef parent() As Long, ByVal a As Long) As Long
  Dim r As Long, t As Long
  r = a
  Do While parent(r) <> r
    r = parent(r)
  Loop
  Do While parent(a) <> a
    t = parent(a)
    parent(a) = r
    a = t
  Loop
  GWFind = r
End Function
