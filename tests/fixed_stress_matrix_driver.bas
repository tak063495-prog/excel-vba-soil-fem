Option Explicit

Public Function FSMatrixDump(ByVal triangle As Boolean) As Variant
  Dim i As Long, a As Long, b As Long, count As Long
  Dim cx As Variant, cy As Variant, mass() As Double, flow() As Double, stab() As Double, coupling() As Double, gravity() As Double, output() As Double
  NumberOfElement = 1: NumberOfMaterial = 1
  ReDim Elem(0): ReDim Material(0)
  Material(0).young = 12000#: Material(0).Poisson = 0.27: Material(0).thickness = 1.7
  If triangle Then
    cx = Array(0#, 1#, 0#, 0#, 0.5, 0.5, 0#, 0#)
    cy = Array(0#, 0#, 1#, 1#, 0#, 0.5, 1#, 0.5)
  Else
    cx = Array(0#, 1#, 1#, 0#, 0.5, 1#, 0.5, 0#)
    cy = Array(0#, 0#, 1#, 1#, 0#, 0.5, 1#, 0.5)
  End If
  For i = 0 To 7
    Elem(0).node(i) = i
    Elem(0).x(i) = 2# * cx(i) + 0.3 * cy(i) + 1#
    Elem(0).y(i) = 0.2 * cx(i) + 1.5 * cy(i) - 2#
  Next i
  If triangle Then Elem(0).node(3) = 2: Elem(0).node(6) = 2
  FSLocalMatrices 0, 0.8, 0.000002, 0.00000003, 0.00000001, 9.80665, mass, flow, stab, coupling, gravity
  ReDim output(0 To 115)
  For a = 0 To 3
    For b = 0 To 3
      output(count) = mass(a, b): count = count + 1
    Next b
  Next a
  For a = 0 To 3
    For b = 0 To 3
      output(count) = flow(a, b): count = count + 1
    Next b
  Next a
  For a = 0 To 3
    For b = 0 To 3
      output(count) = stab(a, b): count = count + 1
    Next b
  Next a
  For a = 0 To 15
    For b = 0 To 3
      output(count) = coupling(a, b): count = count + 1
    Next b
  Next a
  For a = 0 To 3: output(count) = gravity(a): count = count + 1: Next a
  FSMatrixDump = output
End Function

Public Function FSFlowDump(ByVal withBoundary As Boolean) As Variant
  Dim blocks(0 To 3, 0 To 3, 0 To 1) As Double, mapping(0 To 3, 0 To 1) As Long
  Dim rhs(0 To 2) As Double, fixed(0 To 2) As Boolean, prescribed(0 To 2) As Double, guess(0 To 2) As Double
  Dim solution() As Double, residual() As Double, reactions() As Double, output(0 To 9) As Double, i As Long, e As Long, a As Long
  For e = 0 To 1
    For a = 0 To 3: mapping(a, e) = -1: Next a
    mapping(0, e) = e: mapping(1, e) = e + 1
    blocks(0, 0, e) = 2#: blocks(1, 1, e) = 3#: blocks(0, 1, e) = -1#: blocks(1, 0, e) = -1#
  Next e
  rhs(0) = 1#: rhs(1) = 2#: rhs(2) = 3#
  fixed(0) = withBoundary: prescribed(0) = 0.7
  If FSFlowSolve(blocks, mapping, 2, rhs, fixed, prescribed, guess, 0.000000000001, 100, solution, residual, reactions) Then output(9) = 1#
  For i = 0 To 2
    output(i) = solution(i): output(i + 3) = residual(i): output(i + 6) = reactions(i)
  Next i
  FSFlowDump = output
End Function
