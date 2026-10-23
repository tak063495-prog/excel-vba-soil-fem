Option Explicit

' T6 uses the existing eight-slot/native ABI without adding fictitious DOFs.
' Slots: A,B,C,C,AB,BC,C,CA. Slots 3 and 6 have identically zero shape.
Public Function FEMIsT6Data(ByRef el As Element_Data) As Boolean
  If el.IsJoint Then Exit Function
  FEMIsT6Data = (el.node(2) = el.node(3) And el.node(2) = el.node(6) And el.node(0) <> el.node(1) And el.node(1) <> el.node(2))
End Function

Public Function FEMElementIsT6(ByVal e As Long) As Boolean
  If e < 0 Or e >= NumberOfElement Then Exit Function
  FEMElementIsT6 = FEMIsT6Data(Elem(e))
End Function

Public Function FEMT6Dummy(ByVal e As Long, ByVal slot As Long) As Boolean
  If slot = 3 Or slot = 6 Then FEMT6Dummy = FEMElementIsT6(e)
End Function

Public Function FEMMeshSheetHasT6() As Boolean
  Dim ws As Worksheet, data As Variant, r As Long, last As Long
  Set ws = ThisWorkbook.Worksheets("要素データ")
  last = ws.cells(ws.rows.count, 1).End(xlUp).row
  If last < 2 Then Exit Function
  If last = 2 Then
    FEMMeshSheetHasT6 = (UCase$(Trim$(CStr(ws.cells(2, 11).value2))) = "T6")
  Else
    data = ws.range(ws.cells(2, 11), ws.cells(last, 11)).value2
    For r = 1 To UBound(data, 1)
      If UCase$(Trim$(CStr(data(r, 1)))) = "T6" Then FEMMeshSheetHasT6 = True: Exit Function
    Next r
  End If
End Function

Public Function FEMElementSideValid(ByVal e As Long, ByVal side As Long) As Boolean
  FEMElementSideValid = Not (FEMElementIsT6(e) And side = 2)
End Function

Public Function FEMGaussFraction(ByVal e As Long, ByVal gp As Long, Optional ByVal fromSavedTable As Boolean = False) As Double
  Dim triangle As Boolean, gx As Double, gy As Double, ws As Worksheet, dataRow As Variant
  If fromSavedTable Then
    Set ws = ThisWorkbook.Worksheets("要素データ")
    dataRow = e + 2
    If ws.cells(CLng(dataRow), 1).value2 <> e + 1 Then dataRow = Application.Match(e + 1, ws.columns(1), 0)
    If IsError(dataRow) Then Err.Raise vbObjectError + 3977, , "増分図の要素形式が見つかりません。要素=" & CStr(e + 1)
    triangle = (UCase$(Trim$(CStr(ws.cells(CLng(dataRow), 11).value2))) = "T6")
  Else
    triangle = FEMElementIsT6(e)
  End If
  FEMGaussFraction = 0.25
  If triangle Then
    FEMGaussNatural gp, gx, gy
    FEMGaussFraction = (1# - gy) / 4#
  End If
End Function

Public Sub FEMShapeForElement(ByVal e As Long, ByVal xi As Double, ByVal eta As Double, ByRef ns() As Double, ByRef dx() As Double, ByRef dy() As Double)
  If FEMElementIsT6(e) Then
    FEMT6Shape xi, eta, ns, dx, dy
  Else
    P6Q8Shape xi, eta, ns, dx, dy
  End If
End Sub

Public Sub FEMT6Shape(ByVal xi As Double, ByVal eta As Double, ByRef ns() As Double, ByRef dx() As Double, ByRef dy() As Double)
  Dim r As Double, s As Double, a As Double, drx As Double, dry As Double
  Dim nr(0 To 7) As Double, nt(0 To 7) As Double, j As Long
  ' Duffy transform: square top edge collapses to C. Four positive 2x2
  ' quadrature weights integrate the straight T6 elastic stiffness exactly.
  r = (1# + xi) * (1# - eta) / 4#: s = (1# + eta) / 2#: a = 1# - r - s
  drx = (1# - eta) / 4#: dry = -(1# + xi) / 4#
  For j = 0 To 7: ns(j) = 0#: dx(j) = 0#: dy(j) = 0#: Next j
  ns(0) = a * (2# * a - 1#): nr(0) = 1# - 4# * a: nt(0) = nr(0)
  ns(1) = r * (2# * r - 1#): nr(1) = 4# * r - 1#
  ns(2) = s * (2# * s - 1#): nt(2) = 4# * s - 1#
  ns(4) = 4# * a * r: nr(4) = 4# * (a - r): nt(4) = -4# * r
  ns(5) = 4# * r * s: nr(5) = 4# * s: nt(5) = 4# * r
  ns(7) = 4# * s * a: nr(7) = -4# * s: nt(7) = 4# * (a - s)
  For j = 0 To 7
    dx(j) = nr(j) * drx
    dy(j) = nr(j) * dry + nt(j) / 2#
  Next j
End Sub

Public Function FEMGaussShapeValue(ByVal e As Long, ByVal gp As Long, ByVal slot As Long) As Double
  Dim ns(0 To 7) As Double, dx(0 To 7) As Double, dy(0 To 7) As Double
  Dim gx As Double, gy As Double
  If Not FEMElementIsT6(e) Then
    P6EnsureGaussCache
    FEMGaussShapeValue = P6GaussN(gp, slot)
  Else
    FEMGaussNatural gp, gx, gy
    FEMT6Shape gx, gy, ns, dx, dy
    FEMGaussShapeValue = ns(slot)
  End If
End Function

Public Sub FEMGaussNatural(ByVal gp As Long, ByRef gx As Double, ByRef gy As Double)
  Dim a As Double
  a = 1# / Sqr(3#)
  gx = -a: gy = -a
  If gp = 1 Or gp = 2 Then gx = a
  If gp = 2 Or gp = 3 Then gy = a
End Sub

Public Function FEMT6GeometryValid(ByRef el As Element_Data) As Boolean
  Dim ns(0 To 7) As Double, dx(0 To 7) As Double, dy(0 To 7) As Double
  Dim i As Long, g As Long, gx As Double, gy As Double, j11 As Double, j12 As Double, j21 As Double, j22 As Double
  Dim det As Double, geomScale As Double
  geomScale = Abs(el.x(1) - el.x(0)) + Abs(el.y(1) - el.y(0)) + Abs(el.x(2) - el.x(0)) + Abs(el.y(2) - el.y(0))
  If geomScale <= 0# Then Exit Function
  ' Initially support straight-sided T6 only; reject unverified curved maps.
  If Abs(el.x(4) - (el.x(0) + el.x(1)) / 2#) + Abs(el.y(4) - (el.y(0) + el.y(1)) / 2#) > geomScale * 0.000000001 Then Exit Function
  If Abs(el.x(5) - (el.x(1) + el.x(2)) / 2#) + Abs(el.y(5) - (el.y(1) + el.y(2)) / 2#) > geomScale * 0.000000001 Then Exit Function
  If Abs(el.x(7) - (el.x(2) + el.x(0)) / 2#) + Abs(el.y(7) - (el.y(2) + el.y(0)) / 2#) > geomScale * 0.000000001 Then Exit Function
  For g = 0 To 3
    FEMGaussNatural g, gx, gy
    FEMT6Shape gx, gy, ns, dx, dy
    j11 = 0#: j12 = 0#: j21 = 0#: j22 = 0#
    For i = 0 To 7
      j11 = j11 + dx(i) * el.x(i): j12 = j12 + dy(i) * el.x(i)
      j21 = j21 + dx(i) * el.y(i): j22 = j22 + dy(i) * el.y(i)
    Next i
    det = j11 * j22 - j12 * j21
    If det <= geomScale * geomScale * 0.000000000001 Then Exit Function
  Next g
  FEMT6GeometryValid = True
End Function

Public Sub FEMValidateT6Input(ByVal e As Long)
  Dim area As Double, jac As Double, lo As Double, hi As Double, minAngle As Double, maxAngle As Double
  FEMT6MeshMetrics e, P6InputElementNode, P6InputNodeX, P6InputNodeY, area, jac, lo, hi, minAngle, maxAngle
  If area < P5MinElementArea Then P5MinElementArea = area
  If jac < P5MinJacobian Then P5MinJacobian = jac
  If hi / lo > P5MaxAspectRatio Then P5MaxAspectRatio = hi / lo
  If minAngle < P5MinAngleDeg Then P5MinAngleDeg = minAngle
  If maxAngle > P5MaxAngleDeg Then P5MaxAngleDeg = maxAngle
End Sub

Public Sub FEMExtrapolateMixed(ByVal e As Long, ByRef gv() As Double, ByRef cv() As Double, ByRef q8Extrap() As Double)
  Dim a(0 To 2, 0 To 2) As Double, rhs(0 To 2) As Double, basis(0 To 2) As Double
  Dim g As Long, i As Long, j As Long, gx As Double, gy As Double, r As Double, s As Double
  Dim coef(0 To 2) As Double
  If Not FEMElementIsT6(e) Then
    ExtrapolateElementResult gv, cv, q8Extrap
    Exit Sub
  End If
  ' Linear least-squares recovery at the actual triangle integration points.
  For g = 0 To 3
    FEMGaussNatural g, gx, gy
    r = (1# + gx) * (1# - gy) / 4#: s = (1# + gy) / 2#
    basis(0) = 1#: basis(1) = r: basis(2) = s
    For i = 0 To 2
      rhs(i) = rhs(i) + basis(i) * gv(g)
      For j = 0 To 2: a(i, j) = a(i, j) + basis(i) * basis(j): Next j
    Next i
  Next g
  If Not P6InvertDense(a, 3) Then Err.Raise vbObjectError + 3973, , "T6応力回復行列が特異です。"
  For i = 0 To 2
    For j = 0 To 2: coef(i) = coef(i) + a(i, j) * rhs(j): Next j
  Next i
  cv(0) = coef(0): cv(1) = coef(0) + coef(1): cv(2) = coef(0) + coef(2): cv(3) = cv(2)
End Sub

Public Sub FEMT6MeshMetrics(ByVal e As Long, ByRef conn() As Long, ByRef xx() As Double, ByRef yy() As Double, ByRef area As Double, ByRef jac As Double, ByRef minEdge As Double, ByRef maxEdge As Double, ByRef minAngle As Double, ByRef maxAngle As Double)
  Dim el As Element_Data, j As Long, k As Long, prev As Long, nxt As Long, n As Long
  Dim dx1 As Double, dy1 As Double, dx2 As Double, dy2 As Double, angle As Double, length As Double, slots As Variant
  slots = Array(0, 1, 2, 4, 5, 7)
  For j = 0 To 7
    n = conn(e, j + 1)
    If n < LBound(xx) Or n > UBound(xx) Then Err.Raise vbObjectError + 3972, , "T6節点番号が範囲外です。"
    el.node(j) = n: el.x(j) = xx(n): el.y(j) = yy(n)
  Next j
  For j = 0 To 5
    For k = j + 1 To 5
      If el.node(slots(j)) = el.node(slots(k)) Then Err.Raise vbObjectError + 3972, , "T6の実節点が重複しています。"
    Next k
  Next j
  If Not FEMT6GeometryValid(el) Then Err.Raise vbObjectError + 3972, , "T6のJacobian・辺中点が不正です。要素=" & CStr(e)
  area = ((el.x(1) - el.x(0)) * (el.y(2) - el.y(0)) - (el.x(2) - el.x(0)) * (el.y(1) - el.y(0))) / 2#
  jac = area / 2#: minEdge = 1E+308: maxEdge = 0#: minAngle = 180#: maxAngle = 0#
  For j = 0 To 2
    prev = (j + 2) Mod 3: nxt = (j + 1) Mod 3
    dx1 = el.x(prev) - el.x(j): dy1 = el.y(prev) - el.y(j)
    dx2 = el.x(nxt) - el.x(j): dy2 = el.y(nxt) - el.y(j)
    length = Sqr(dx2 * dx2 + dy2 * dy2)
    If length < minEdge Then minEdge = length
    If length > maxEdge Then maxEdge = length
    angle = P1Atan2(Abs(dx1 * dy2 - dy1 * dx2), dx1 * dx2 + dy1 * dy2) * 180# / 3.14159265358979
    If angle < minAngle Then minAngle = angle
    If angle > maxAngle Then maxAngle = angle
  Next j
End Sub
