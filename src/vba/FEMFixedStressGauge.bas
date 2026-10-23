Option Explicit
Private Function FSRoot(ByRef parent() As Long, ByVal node As Long) As Long
  Dim root As Long, nextNode As Long
  root = node
  Do While parent(root) <> root: root = parent(root): Loop
  Do While parent(node) <> node
    nextNode = parent(node): parent(node) = root: node = nextNode
  Loop
  FSRoot = root
End Function

Public Sub FSCheckPressureGauge(ByRef map() As Long, ByRef mass() As Double, ByRef coupling() As Double, ByRef locked() As Boolean)
  Dim parent() As Long, used() As Boolean, anchored() As Boolean, storage() As Double, magnitude() As Double
  Dim values As Object, e As Long, a As Long, b As Long, n As Long, first As Long, root As Long, other As Long, d As Long, dof As Long
  Dim key As String, entry As Variant, value As Double, i As Long
  ReDim parent(NumberOfFreeNode - 1): ReDim used(NumberOfFreeNode - 1): ReDim anchored(NumberOfFreeNode - 1)
  ReDim storage(NumberOfFreeNode - 1): ReDim magnitude(NumberOfFreeNode - 1)
  For n = 0 To NumberOfFreeNode - 1: parent(n) = n: Next n
  For e = 0 To NumberOfElement - 1
    first = -1
    For a = 0 To 3
      n = map(a, e)
      If n >= 0 Then
        used(n) = True
        If first < 0 Then
          first = n
        Else
          root = FSRoot(parent, first): other = FSRoot(parent, n): parent(other) = root
        End If
      End If
    Next a
  Next e
  For n = 0 To NumberOfFreeNode - 1
    If used(n) Then
      root = FSRoot(parent, n)
      If locked(n) Then anchored(root) = True
    End If
  Next n
  Set values = CreateObject("Scripting.Dictionary")
  For e = 0 To NumberOfElement - 1
    first = map(0, e)
    If first >= 0 Then
      root = FSRoot(parent, first)
      For a = 0 To 3
        For b = 0 To 3: storage(root) = storage(root) + mass(a, b, e): Next b
      Next a
      For d = 0 To 15
        dof = Elem(e).elnode(d)
        If dof >= 0 Then
          If NodeCond(dof) = 0 Then
            value = 0#
            For a = 0 To 3: value = value + coupling(d, a, e): Next a
            key = CStr(root) & ":" & CStr(dof)
            If Not values.Exists(key) Then values.Add key, 0#
            values(key) = CDbl(values(key)) + value
            magnitude(root) = magnitude(root) + Abs(value)
          End If
        End If
      Next d
    End If
  Next e
  For Each entry In values.keys
    i = CLng(Split(CStr(entry), ":")(0))
    If Abs(CDbl(values(entry))) > 0.000000000001 * magnitude(i) Then anchored(i) = True
  Next entry
  For n = 0 To NumberOfFreeNode - 1
    If used(n) And parent(n) = n Then
      If storage(n) = 0# And Not anchored(n) Then FSCoupledInputError "非排水・貯留0で体積変形も拘束され、水圧の基準が定まりません。有限の貯留または水圧境界が必要です。"
    End If
  Next n
End Sub
