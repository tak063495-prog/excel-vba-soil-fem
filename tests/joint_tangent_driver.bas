Option Explicit

' Test-only driver: central difference of the actual production JOINT force routine.
Public Function JTCheck() As Variant
  Dim saved() As Double, savedForce() As Double, patch() As Double
  Dim e As Long, i As Long, j As Long, a As Long, mode As Long, dof As Long, expected As Long
  Dim k0(0 To 11, 0 To 11) As Double, fp(0 To 11) As Double, fm(0 To 11) As Double
  Dim tx As Double, ty As Double, nx As Double, ny As Double, length As Double
  Dim gap As Double, slip As Double, h As Double, difference As Double, maxErr As Double, maxK As Double, asym As Double
  Dim expectedNormal As Double, expectedShear As Double, cap As Double, stateKey As Long
  Dim positiveKey As Long, negativeKey As Long, report(0 To 3) As String, modeName As Variant
  Dim number As Long, description As String
  On Error GoTo Failed
  saved = TDisp: savedForce = iNForce
  e = -1
  For i = 0 To NumberOfElement - 1
    If Elem(i).IsJoint And P3IsElementActive(i) Then e = i: Exit For
  Next i
  If e < 0 Then Err.Raise vbObjectError + 7980, , "No active JOINT in tangent fixture."
  length = Sqr((Elem(e).x(1) - Elem(e).x(0)) ^ 2 + (Elem(e).y(1) - Elem(e).y(0)) ^ 2)
  tx = (Elem(e).x(1) - Elem(e).x(0)) / length: ty = (Elem(e).y(1) - Elem(e).y(0)) / length
  nx = ty: ny = -tx: h = 0.0000000001
  modeName = Array("STICK", "SLIP_POSITIVE", "SLIP_NEGATIVE", "OPEN")
  For mode = 0 To 3
    gap = -0.00001: slip = 0.00000001: expected = 0
    If mode = 1 Then slip = 0.001: expected = 2
    If mode = 2 Then slip = -0.001: expected = 2
    If mode = 3 Then gap = 0.00001: slip = 0.001: expected = 1
    ReDim patch(lastDof)
    For a = 0 To 2
      dof = 2 * Elem(e).node(a + 3)
      patch(dof) = gap * nx + slip * tx: patch(dof + 1) = gap * ny + slip * ty
    Next a
    TDisp = patch
    JTForce e
    For i = 0 To 11
      For j = 0 To 11: k0(i, j) = Elem(e).kmat(i, j): Next j
    Next i
    expectedNormal = -Material(Elem(e).MatNo).kn * gap
    cap = Material(Elem(e).MatNo).cohesion + WorksheetFunction.Max(0#, expectedNormal) * Tan(Material(Elem(e).MatNo).fai * 3.14159265358979 / 180#)
    expectedShear = Material(Elem(e).MatNo).ks * slip
    If Abs(expectedShear) > cap Then expectedShear = Sgn(slip) * cap
    If mode = 3 Then expectedNormal = 0#: expectedShear = 0#
    For a = 0 To 2
      If Elem(e).JointContact(a) <> expected Then Err.Raise vbObjectError + 7980, , "Wrong contact branch."
      If Abs(Elem(e).Stmat(0, a) - expectedNormal) > 0.000001 Then Err.Raise vbObjectError + 7980, , "Normal law changed."
      If Abs(Elem(e).Stmat(1, a) - expectedShear) > 0.000001 Then Err.Raise vbObjectError + 7980, , "Shear law changed."
      If Abs(Elem(e).Stmat(16, a) - expectedNormal - GWJointPressure(e, a)) > 0.000001 Then Err.Raise vbObjectError + 7980, , "Water added twice or missing."
    Next a
    stateKey = P3JointTangentState(e, 0)
    If mode = 1 Then positiveKey = stateKey
    If mode = 2 Then negativeKey = stateKey
    maxErr = 0#: maxK = 0#: asym = 0#
    For j = 0 To 11
      dof = Elem(e).ElNode(j)
      TDisp = patch: TDisp(dof) = TDisp(dof) + h: JTForce e
      For i = 0 To 11: fp(i) = iNForce(Elem(e).ElNode(i)): Next i
      TDisp = patch: TDisp(dof) = TDisp(dof) - h: JTForce e
      For i = 0 To 11
        fm(i) = iNForce(Elem(e).ElNode(i))
        difference = Abs((fp(i) - fm(i)) / (2# * h) - k0(i, j))
        If difference > maxErr Then maxErr = difference
        If Abs(k0(i, j)) > maxK Then maxK = Abs(k0(i, j))
        If Abs(k0(i, j) - k0(j, i)) > asym Then asym = Abs(k0(i, j) - k0(j, i))
      Next i
    Next j
    If maxK < 1# Then maxK = 1#
    If maxErr / maxK > 0.0000001 Then Err.Raise vbObjectError + 7980, , "Force/tangent central difference mismatch."
    If expected = 2 And asym <= 0# Then Err.Raise vbObjectError + 7980, , "Sliding tangent was incorrectly symmetrized."
    report(mode) = CStr(modeName(mode)) & "|" & CStr(maxErr / maxK) & "|" & CStr(asym) & "|" & CStr(expected) & "|" & CStr(stateKey)
  Next mode
  If positiveKey = negativeKey Then Err.Raise vbObjectError + 7980, , "Factor key misses reversed friction."
  TDisp = saved: P3JointEval e, False: iNForce = savedForce
  JTCheck = report
  Exit Function
Failed:
  number = Err.Number: description = Err.Description
  On Error Resume Next
  TDisp = saved: iNForce = savedForce
  On Error GoTo 0
  JTCheck = Array("ERROR|" & CStr(number) & "|" & description)
End Function

Private Sub JTForce(ByVal e As Long)
  Dim i As Long
  For i = 0 To lastDof: iNForce(i) = 0#: Next i
  P3JointEval e, True
End Sub
