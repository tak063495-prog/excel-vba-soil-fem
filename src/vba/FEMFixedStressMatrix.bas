Option Explicit

' Pressure Q4/P1, displacement Q8/T6. Q uses tension-positive gradients.
' M: 1/kPa * m3; H: m3/(kPa s); Q: m2; gravity: m3/s.
Public Sub FSPressureShape(ByVal triangle As Boolean, ByVal xi As Double, ByVal eta As Double, ByRef ns() As Double, ByRef dx() As Double, ByRef dy() As Double)
  Dim r As Double, s As Double
  If triangle Then
    r = (1# + xi) * (1# - eta) / 4#: s = (1# + eta) / 2#
    ns(0) = 1# - r - s: ns(1) = r: ns(2) = s: ns(3) = 0#
    dx(0) = -(1# - eta) / 4#: dx(1) = -dx(0): dx(2) = 0#: dx(3) = 0#
    dy(1) = -(1# + xi) / 4#: dy(2) = 0.5: dy(0) = -dy(1) - dy(2): dy(3) = 0#
  Else
    ns(0) = (1# - xi) * (1# - eta) / 4#: ns(1) = (1# + xi) * (1# - eta) / 4#
    ns(2) = (1# + xi) * (1# + eta) / 4#: ns(3) = (1# - xi) * (1# + eta) / 4#
    dx(0) = -(1# - eta) / 4#: dx(1) = -dx(0): dx(2) = (1# + eta) / 4#: dx(3) = -dx(2)
    dy(0) = -(1# - xi) / 4#: dy(1) = -(1# + xi) / 4#: dy(2) = -dy(1): dy(3) = -dy(0)
  End If
End Sub

Public Sub FSLocalMatrices(ByVal e As Long, ByVal biot As Double, ByVal storage As Double, ByVal kx As Double, ByVal ky As Double, ByVal gamma As Double, ByRef mass() As Double, ByRef flow() As Double, ByRef stab() As Double, ByRef coupling() As Double, ByRef gravity() As Double)
  Dim a As Long, b As Long, i As Long, ix As Long, iy As Long, mat As Long
  Dim ns(0 To 7) As Double, dx(0 To 7) As Double, dy(0 To 7) As Double
  Dim pn(0 To 3) As Double, px(0 To 3) As Double, py(0 To 3) As Double
  Dim z(0 To 2) As Double, w(0 To 2) As Double, j11 As Double, j12 As Double, j21 As Double, j22 As Double
  Dim det As Double, weight As Double, gx As Double, gy As Double, bulk As Double, thick As Double
  If e < 0 Or e >= NumberOfElement Then Err.Raise vbObjectError + 8101, , "FS: element out of range"
  If Not P2IsFinite(biot) Or Not P2IsFinite(storage) Or Not P2IsFinite(kx) Or Not P2IsFinite(ky) Or Not P2IsFinite(gamma) Then Err.Raise vbObjectError + 8101, , "FS: nonfinite coefficient"
  If biot <= 0# Or biot > 1# Or storage < 0# Or kx <= 0# Or ky <= 0# Or gamma <= 0# Then Err.Raise vbObjectError + 8101, , "FS: invalid coefficient"
  mat = Elem(e).MatNo: thick = Material(mat).thickness
  If thick <= 0# Or Material(mat).young <= 0# Or Material(mat).Poisson <= -1# Or Material(mat).Poisson >= 0.5 Then Err.Raise vbObjectError + 8101, , "FS: invalid drained modulus/thickness"
  bulk = Material(mat).young / (3# * (1# - 2# * Material(mat).Poisson))
  ReDim mass(0 To 3, 0 To 3): ReDim flow(0 To 3, 0 To 3): ReDim stab(0 To 3, 0 To 3)
  ReDim coupling(0 To 15, 0 To 3): ReDim gravity(0 To 3)
  z(0) = -Sqr(0.6): z(1) = 0#: z(2) = Sqr(0.6)
  w(0) = 5# / 9#: w(1) = 8# / 9#: w(2) = 5# / 9#
  For iy = 0 To 2
    For ix = 0 To 2
      FEMShapeForElement e, z(ix), z(iy), ns, dx, dy
      FSPressureShape FEMElementIsT6(e), z(ix), z(iy), pn, px, py
      j11 = 0#: j12 = 0#: j21 = 0#: j22 = 0#
      For i = 0 To 7
        j11 = j11 + dx(i) * Elem(e).x(i): j12 = j12 + dy(i) * Elem(e).x(i)
        j21 = j21 + dx(i) * Elem(e).y(i): j22 = j22 + dy(i) * Elem(e).y(i)
      Next i
      det = j11 * j22 - j12 * j21
      If det <= 0# Or Not P2IsFinite(det) Then Err.Raise vbObjectError + 8102, , "FS: invalid Jacobian"
      weight = det * w(ix) * w(iy) * thick
      For i = 0 To 7
        gx = (j22 * dx(i) - j21 * dy(i)) / det: gy = (-j12 * dx(i) + j11 * dy(i)) / det
        dx(i) = gx: dy(i) = gy
      Next i
      For a = 0 To 3
        gx = (j22 * px(a) - j21 * py(a)) / det: gy = (-j12 * px(a) + j11 * py(a)) / det
        px(a) = gx: py(a) = gy
      Next a
      For a = 0 To 3
        gravity(a) = gravity(a) + ky * py(a) * weight
        For b = 0 To 3
          mass(a, b) = mass(a, b) + storage * pn(a) * pn(b) * weight
          flow(a, b) = flow(a, b) + (kx * px(a) * px(b) + ky * py(a) * py(b)) * weight / gamma
          stab(a, b) = stab(a, b) + biot * biot / bulk * pn(a) * pn(b) * weight
        Next b
        For i = 0 To 7
          coupling(2 * i, a) = coupling(2 * i, a) + biot * dx(i) * pn(a) * weight
          coupling(2 * i + 1, a) = coupling(2 * i + 1, a) + biot * dy(i) * pn(a) * weight
        Next i
      Next a
    Next ix
  Next iy
End Sub

Public Sub FSFlowMatVec(ByRef blocks() As Double, ByRef mapping() As Long, ByVal ne As Long, ByRef xv() As Double, ByRef yv() As Double)
  Dim e As Long, a As Long, b As Long, ia As Long, ib As Long, i As Long
  For i = 0 To UBound(xv): yv(i) = 0#: Next i
  For e = 0 To ne - 1
    For a = 0 To 3
      ia = mapping(a, e)
      If ia >= 0 Then
        For b = 0 To 3
          ib = mapping(b, e)
          If ib >= 0 Then yv(ia) = yv(ia) + blocks(a, b, e) * xv(ib)
        Next b
      End If
    Next a
  Next e
End Sub

' Element-wise PCG, including original-equation residual and boundary reaction.
' No global dense matrix, no diagonal floors. tol is the relative true residual.
Public Function FSFlowSolve(ByRef blocks() As Double, ByRef mapping() As Long, ByVal ne As Long, ByRef rhs() As Double, ByRef locked() As Boolean, ByRef prescribed() As Double, ByRef guess() As Double, ByVal tol As Double, ByVal maxit As Long, ByRef solution() As Double, ByRef residual() As Double, ByRef reaction() As Double) As Boolean
  Dim n As Long, i As Long, e As Long, a As Long, b As Long, ia As Long, ib As Long, it As Long
  Dim diag() As Double, rr() As Double, zz() As Double, direction() As Double, product() As Double, values() As Double
  Dim rho As Double, nextRho As Double, curvature As Double, rate As Double, norm As Double, residualScale As Double
  n = UBound(rhs) + 1
  If n < 1 Or tol <= 0# Or maxit < 1 Then Err.Raise vbObjectError + 8103, , "FS: invalid flow dimensions/tolerance"
  ReDim solution(0 To n - 1): ReDim residual(0 To n - 1): ReDim reaction(0 To n - 1)
  ReDim diag(0 To n - 1): ReDim rr(0 To n - 1): ReDim zz(0 To n - 1)
  ReDim direction(0 To n - 1): ReDim product(0 To n - 1): ReDim values(0 To n - 1)
  For e = 0 To ne - 1
    For a = 0 To 3
      ia = mapping(a, e)
      If ia >= 0 Then
        For b = 0 To 3
          ib = mapping(b, e)
          If ib = ia Then diag(ia) = diag(ia) + blocks(a, b, e)
        Next b
      End If
    Next a
  Next e
  For i = 0 To n - 1
    If locked(i) Then
      solution(i) = prescribed(i): values(i) = prescribed(i)
    Else
      If diag(i) <= 0# Or Not P2IsFinite(diag(i)) Then Err.Raise vbObjectError + 8103, , "FS: singular flow diagonal"
      solution(i) = guess(i)
    End If
  Next i
  FSFlowMatVec blocks, mapping, ne, values, product
  For i = 0 To n - 1
    If Not locked(i) Then residualScale = residualScale + (rhs(i) - product(i)) ^ 2
  Next i
  residualScale = Sqr(residualScale): If residualScale < 1E-30 Then residualScale = 1E-30
  For it = 0 To maxit
    FSFlowMatVec blocks, mapping, ne, solution, product
    norm = 0#: nextRho = 0#
    For i = 0 To n - 1
      residual(i) = rhs(i) - product(i): reaction(i) = 0#: rr(i) = 0#: zz(i) = 0#
      If locked(i) Then
        reaction(i) = -residual(i): residual(i) = 0#
      Else
        rr(i) = residual(i): zz(i) = rr(i) / diag(i)
        norm = norm + rr(i) * rr(i): nextRho = nextRho + rr(i) * zz(i)
      End If
    Next i
    If Sqr(norm) <= tol * residualScale Then FSFlowSolve = True: Exit Function
    If it = maxit Then Exit Function
    If it = 0 Then
      direction = zz
    Else
      If rho <= 0# Then Err.Raise vbObjectError + 8103, , "FS: invalid PCG residual"
      For i = 0 To n - 1: direction(i) = zz(i) + nextRho / rho * direction(i): Next i
    End If
    rho = nextRho
    FSFlowMatVec blocks, mapping, ne, direction, product
    curvature = 0#
    For i = 0 To n - 1
      If Not locked(i) Then curvature = curvature + direction(i) * product(i)
    Next i
    If curvature <= 0# Or Not P2IsFinite(curvature) Then Err.Raise vbObjectError + 8103, , "FS: flow matrix is not SPD"
    rate = rho / curvature
    For i = 0 To n - 1
      If Not locked(i) Then solution(i) = solution(i) + rate * direction(i)
    Next i
  Next it
End Function
