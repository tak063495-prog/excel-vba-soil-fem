Option Explicit
' Mechanical displacement pairs remain independent; only hydraulic/topological aliases merge.
Public GJReady As Boolean
Private jMode() As String, jProfile() As String
Private jOwner() As Long, jSide() As Long, jActive() As Boolean
Private topo() As Long, flow() As Long

Private Sub GJFail(ByVal message As String)
  Err.Raise vbObjectError + 7970, "FEMJointWater", message
End Sub

Private Function Root(ByRef parent() As Long, ByVal n As Long) As Long
  Do While parent(n) <> n
    parent(n) = parent(parent(n)): n = parent(n)
  Loop
  Root = n
End Function

Private Sub Merge(ByRef parent() As Long, ByVal a As Long, ByVal b As Long)
  a = Root(parent, a): b = Root(parent, b)
  If a < b Then parent(b) = a Else parent(a) = b
End Sub

Public Function GJTopoNode(ByVal n As Long) As Long
  GJTopoNode = n
  If GWActive And GJReady Then GJTopoNode = topo(n)
End Function

Public Function GWFlowNode(ByVal e As Long, ByVal localNode As Long) As Long
  GWFlowNode = Elem(e).node(localNode)
  If GWActive And GJReady Then GWFlowNode = flow(GWFlowNode)
End Function

Public Function GJParticipates(ByVal e As Long) As Boolean
  GJParticipates = True
  If GWActive And GJReady Then
    If Elem(e).IsJoint Then GJParticipates = jActive(e)
  End If
End Function

Public Function GJSource(ByVal e As Long) As String
  If GWActive And GJReady Then GJSource = jMode(e) & ":" & jProfile(e)
End Function

' Locate each side by all three node IDs. Coordinates alone must not connect water domains.
Public Sub GJPrepare()
  Dim e As Long, q As Long, side As Long, k As Long, b As Long, i As Long, r As Long
  Dim ids As Variant, ws As Worksheet, found As Long, aMat As Long, bMat As Long
  Dim mode As String, profile As String, scaleValue As Double
  GJReady = False
  ReDim topo(NumberOfFreeNode - 1): ReDim flow(NumberOfFreeNode - 1)
  ReDim jMode(NumberOfElement - 1): ReDim jProfile(NumberOfElement - 1)
  ReDim jOwner(0 To 1, 0 To NumberOfElement - 1): ReDim jSide(0 To 1, 0 To NumberOfElement - 1)
  ReDim jActive(NumberOfElement - 1)
  For i = 0 To NumberOfFreeNode - 1: topo(i) = i: flow(i) = i: Next i
  Set ws = ThisWorkbook.Worksheets("接合")
  For e = 0 To NumberOfElement - 1
    If Not Elem(e).IsJoint Then GoTo NextJoint
    For b = 0 To 1
      jOwner(b, e) = -1
      For q = 0 To NumberOfElement - 1
        If Not Elem(q).IsJoint Then
          For side = 0 To 3
            ids = Array(Elem(q).node(side), Elem(q).node((side + 1) Mod 4), Elem(q).node(side + 4))
            If ((ids(0) = Elem(e).node(3 * b) And ids(1) = Elem(e).node(3 * b + 1)) Or (ids(1) = Elem(e).node(3 * b) And ids(0) = Elem(e).node(3 * b + 1))) And ids(2) = Elem(e).node(3 * b + 2) Then
              If jOwner(b, e) >= 0 Then GJFail "JOINT面の隣接要素が一意ではありません。要素=" & CStr(e + 1)
              jOwner(b, e) = q: jSide(b, e) = side
            End If
          Next side
        End If
      Next q
      If jOwner(b, e) < 0 Then GJFail "JOINTの各面にQ8要素が必要です。要素=" & CStr(e + 1)
    Next b
    If jOwner(0, e) = jOwner(1, e) Then GJFail "JOINT両面が同じ要素です。"
    scaleValue = 1# + Abs(Elem(e).x(1) - Elem(e).x(0)) + Abs(Elem(e).y(1) - Elem(e).y(0))
    For i = 0 To 2
      If Abs(Elem(e).x(i) - Elem(e).x(i + 3)) + Abs(Elem(e).y(i) - Elem(e).y(i + 3)) > 0.000000001 * scaleValue Then GJFail "水圧JOINTは初期に同位置の節点対が必要です。"
    Next i
    If Abs(Elem(e).x(2) - 0.5 * (Elem(e).x(0) + Elem(e).x(1))) + Abs(Elem(e).y(2) - 0.5 * (Elem(e).y(0) + Elem(e).y(1))) > 0.000000001 * scaleValue Then GJFail "曲線JOINTの水圧積分は未対応です。"
    If Elem(jOwner(0, e)).node(jSide(0, e)) <> Elem(e).node(0) Or Elem(jOwner(1, e)).node(jSide(1, e)) <> Elem(e).node(4) Then GJFail "JOINTの面方向は主面の外向き法線へ統一してください。"
    For b = 0 To 1
      If Abs(Material(Elem(jOwner(b, e)).MatNo).thickness - Material(Elem(e).MatNo).thickness) > 0.000000001 Then GJFail "水圧JOINTと両側要素の厚さを一致させてください。"
    Next b
    jActive(e) = P3ElementActive(e) And P3ElementActive(jOwner(0, e)) And P3ElementActive(jOwner(1, e))
    If Not jActive(e) Then GoTo NextJoint
    aMat = Elem(jOwner(0, e)).MatNo + 1: bMat = Elem(jOwner(1, e)).MatNo + 1: found = 0
    For r = 2 To ws.cells(ws.rows.count, 1).End(xlUp).row
      If val(CStr(ws.cells(r, 4).value2)) <> 0# And val(CStr(ws.cells(r, 3).value2)) = Elem(e).MatNo + 1 Then
        If (val(CStr(ws.cells(r, 1).value2)) = aMat And val(CStr(ws.cells(r, 2).value2)) = bMat) Or (val(CStr(ws.cells(r, 1).value2)) = bMat And val(CStr(ws.cells(r, 2).value2)) = aMat) Then
          found = found + 1: mode = UCase$(Trim$(CStr(ws.cells(r, 6).value2))): profile = Trim$(CStr(ws.cells(r, 7).value2))
        End If
      End If
    Next r
    If found <> 1 Then GJFail "有効JOINTに対応する接合行を1行指定してください。要素=" & CStr(e + 1)
    If mode <> "SOIL" And mode <> "CONTINUOUS" And mode <> "PRESCRIBED" Then GJFail "接合F列へSOIL/CONTINUOUS/PRESCRIBEDを指定してください。"
    If mode = "PRESCRIBED" And Len(profile) = 0 Then GJFail "指定水圧には接合G列の水位線IDが必要です。"
    If mode <> "PRESCRIBED" And Len(profile) > 0 Then GJFail "接合G列はPRESCRIBEDだけに指定します。"
    jMode(e) = mode: jProfile(e) = profile
    For i = 0 To 2: Merge topo, Elem(e).node(i), Elem(e).node(i + 3): Next i
    If mode = "CONTINUOUS" Then
      If Material(aMat - 1).kind <> "SOIL" Or Material(bMat - 1).kind <> "SOIL" Then GJFail "CONTINUOUSには両側の透水地盤が必要です。"
      For i = 0 To 2: Merge flow, Elem(e).node(i), Elem(e).node(i + 3): Next i
    ElseIf mode = "SOIL" Then
      If (Material(aMat - 1).kind = "SOIL") = (Material(bMat - 1).kind = "SOIL") Then GJFail "SOILは地盤と構造体の接合に使用します。両側地盤はCONTINUOUS、指定水域はPRESCRIBEDを選んでください。"
    End If
NextJoint:
  Next e
  For i = 0 To NumberOfFreeNode - 1: topo(i) = Root(topo, i): flow(i) = Root(flow, i): Next i
  GJReady = True
End Sub

Public Sub GJApplyHeads(ByRef fixed() As Boolean, ByRef heads() As Double)
  Dim e As Long, b As Long, a As Long, n As Long, q As Long, value As Double, profile As String
  Dim ids As Variant
  For e = 0 To NumberOfElement - 1
    If Elem(e).IsJoint And jActive(e) And jMode(e) = "PRESCRIBED" Then
      profile = jProfile(e)
      For b = 0 To 1
        q = jOwner(b, e)
        If Material(Elem(q).MatNo).kind = "SOIL" Then
          ids = Array(jSide(b, e), (jSide(b, e) + 1) Mod 4, jSide(b, e) + 4)
          For a = 0 To 2
            n = Elem(q).node(ids(a)): value = GWProfileLevel(profile, Elem(q).x(ids(a)))
            If value < Elem(q).y(ids(a)) - 0.000000001 Then GJFail "STEADYの指定水圧JOINTは全面飽和が必要です。"
            If fixed(n) And Abs(heads(n) - value) > 0.000000001 Then GJFail "JOINT指定水位と固定水頭が競合しています。"
            fixed(n) = True: heads(n) = value
          Next a
        End If
      Next b
    End If
  Next e
  For n = 0 To NumberOfFreeNode - 1
    q = flow(n)
    If fixed(n) And q <> n Then
      If fixed(q) And Abs(heads(q) - heads(n)) > 0.000000001 Then GJFail "水頭連続JOINTの固定水頭が競合しています。"
      fixed(q) = True: heads(q) = heads(n): fixed(n) = False
    End If
  Next n
End Sub

Public Sub GJExpandHeads(ByRef heads() As Double)
  Dim n As Long
  For n = 0 To NumberOfFreeNode - 1: heads(n) = heads(flow(n)): Next n
End Sub

Public Sub GJSetPressures(ByRef target() As Double, ByVal gamma As Double, ByVal steady As Boolean)
  Dim e As Long, gp As Long, b As Long, q As Long, a As Long, count As Long
  Dim xi As Double, value As Double, p As Double, px As Double, py As Double, profile As String
  Dim shape(0 To 2) As Double
  For e = 0 To NumberOfElement - 1
    If Elem(e).IsJoint And jActive(e) Then
      For gp = 0 To 2
        xi = (gp - 1) * Sqr(0.6): shape(0) = 0.5 * xi * (xi - 1#): shape(1) = 0.5 * xi * (xi + 1#): shape(2) = 1# - xi * xi
        px = 0#: py = 0#: count = 0: p = 0#
        For a = 0 To 2: px = px + shape(a) * Elem(e).x(a): py = py + shape(a) * Elem(e).y(a): Next a
        If jMode(e) = "PRESCRIBED" Then
          p = gamma * (GWProfileLevel(jProfile(e), px) - py): If p < 0# Then p = 0#
        End If
        For b = 0 To 1
          q = jOwner(b, e)
          If Material(Elem(q).MatNo).kind = "SOIL" Then
            value = GWFacePressure(q, jSide(b, e), xi, Elem(e).node(0))
            If count > 0 Or jMode(e) = "PRESCRIBED" Then
              If Abs(value - p) > 0.0000001 * (1# + Abs(p)) Then GJFail "JOINT両面/指定水位の水圧が一致しません。単純平均は行いません。"
            End If
            p = value: count = count + 1
          End If
        Next b
        target(gp, e) = p
      Next gp
    End If
  Next e
End Sub

Public Sub GJAddForces(ByRef target() As Double, ByRef force() As Double, ByVal steady As Boolean)
  Dim e As Long, b As Long, gp As Long, a As Long, n As Long, q As Long
  Dim profile As String, xi As Double, jac As Double, nx As Double, ny As Double, weight As Double, value As Double
  Dim shape(0 To 2) As Double
  For e = 0 To NumberOfElement - 1
    If Elem(e).IsJoint And jActive(e) Then
      profile = jProfile(e)
      If Not steady Or jMode(e) = "PRESCRIBED" Then
        If jMode(e) <> "PRESCRIBED" Then
          q = jOwner(0, e): If Material(Elem(q).MatNo).kind <> "SOIL" Then q = jOwner(1, e)
          profile = GWMaterialProfile(Elem(q).MatNo)
        End If
        For b = 0 To 1: GWAddProfileFaceLoad jOwner(b, e), jSide(b, e), profile, force: Next b
      Else
        jac = 0.5 * Sqr((Elem(e).x(1) - Elem(e).x(0)) ^ 2 + (Elem(e).y(1) - Elem(e).y(0)) ^ 2)
        nx = (Elem(e).y(1) - Elem(e).y(0)) / (2# * jac): ny = -(Elem(e).x(1) - Elem(e).x(0)) / (2# * jac)
        For gp = 0 To 2
          xi = (gp - 1) * Sqr(0.6): weight = 5# / 9#: If gp = 1 Then weight = 8# / 9#
          shape(0) = 0.5 * xi * (xi - 1#): shape(1) = 0.5 * xi * (xi + 1#): shape(2) = 1# - xi * xi
          For a = 0 To 2
            value = target(gp, e) * shape(a) * weight * jac * Material(Elem(e).MatNo).thickness
            n = Elem(e).node(a): force(2 * n) = force(2 * n) - value * nx: force(2 * n + 1) = force(2 * n + 1) - value * ny
            n = Elem(e).node(a + 3): force(2 * n) = force(2 * n) + value * nx: force(2 * n + 1) = force(2 * n + 1) + value * ny
          Next a
        Next gp
      End If
    End If
  Next e
End Sub
