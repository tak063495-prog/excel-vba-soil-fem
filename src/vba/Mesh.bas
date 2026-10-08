
Option Explicit
Private P5ModelScale As Double
Private P5GeometryTolerance As Double
Private P5AreaTolerance As Double
Private P5MinElementArea As Double
Private P5MinJacobian As Double
Private P5MaxAspectRatio As Double
Private P5MinAngleDeg As Double
Private P5MaxAngleDeg As Double
Private P5BoundaryEdgeCount As Long
Private P5SharedEdgeCount As Long
Private P5IsolatedNodeCount As Long
Private P5ConnectedComponentCount As Long
Private P5BoundarySelfIntersectionCount As Long
Private P5LocateCallCount As Long
Private P5LocateTotalIterations As Long
Private P5LocateMaxIterations As Long
Private P5LocateLastIterations As Long
Private P5LocateLastElement As Long
Private P5LocateLastFailure As String
Private locateStamp() As Long, locateCapacity As Long, locateEpoch As Long

    ' 旧「図形細分化」経路の配列は現行Q8経路では使用しない。
    ' 互換手続きの宣言だけ残し、固定確保は行わない。
    Public ibex() As Long
    Public ibin() As Long
    Public mibex() As Long
    Public mdelx() As Long
    Public ibno() As Long
    Public mibno() As Long
    
    Public mindex() As Long
    Public mtj() As Long
    Public delx() As Double
    Public NEX As Long
    Public NIN As Long
    Public NOB As Long
    Public NIB As Long
    Public node As Long
    Public NELM As Long
    Public ntp As Long
    Public ia As Long
    Public ib As Long
    Public ic As Long
    Private meshXmin As Double
    Private meshXmax As Double
    Private meshYmin As Double
    Private meshYmax As Double
    Public rax As Double
    Public ray As Double
    Public dMax As Double
    Public list() As Long
    Public istack() As Long
    Public ibr() As Long
    Public iadres() As Long
    Public px() As Double
    Public py() As Double
    Public sbex() As Long
    Public sx() As Double
    Public sy() As Double
    Public mpx() As Double
    Public mpy() As Double
    Public xlok() As Long
    Public NNEX As Long
    Public NNIN As Long
           
    Public ifix() As Long
    Public jac() As Long
    Public idm() As Long
    Public mmtj() As Long
    Public smtj() As Long
    Public pmtj() As Long
    
    Public kv() As Long
    Public id() As Long
    Public map() As Long
    Public ihen() As Long
    Public jhen() As Long
    Public khen() As Long
    Public jnb() As Long
    Public nei() As Long

Private Function P5ReadSetting(ByVal keyName As String, ByVal defaultValue As Double) As Double
    P5ReadSetting = FEMReadSetting(keyName, defaultValue)
End Function

Private Sub P5WriteNumericSetting(ByVal keyName As String, ByVal value As Double)
    FEMWriteNumericSetting keyName, value
End Sub

Private Sub P5WriteTextSetting(ByVal keyName As String, ByVal value As String)
    FEMWriteTextSetting keyName, value
End Sub

Public Sub ValidateMeshInputCapacity(ByVal externalBoundaryCount As Long, ByVal internalBoundaryCount As Long, _
                                     ByVal boundaryNodeCount As Long, ByVal interiorNodeCount As Long)
    Dim totalBoundary As Long
    totalBoundary = externalBoundaryCount + internalBoundaryCount
    If externalBoundaryCount < 0 Or externalBoundaryCount > UBound(ibex) Then Err.Raise vbObjectError + 3401, "Mesh.ValidateMeshInputCapacity", "外部境界数が固定配列容量を超えています。"
    If internalBoundaryCount < 0 Or internalBoundaryCount > UBound(ibin) Then Err.Raise vbObjectError + 3402, "Mesh.ValidateMeshInputCapacity", "内部境界数が固定配列容量を超えています。"
    If totalBoundary > UBound(mibno, 1) Then Err.Raise vbObjectError + 3403, "Mesh.ValidateMeshInputCapacity", "境界配列の第1次元容量を超えています。"
    If totalBoundary > UBound(delx) Then Err.Raise vbObjectError + 3404, "Mesh.ValidateMeshInputCapacity", "境界間隔配列の容量を超えています。"
    If externalBoundaryCount > UBound(sbex) Then Err.Raise vbObjectError + 3405, "Mesh.ValidateMeshInputCapacity", "外部境界補助配列の容量を超えています。"
    If boundaryNodeCount < 0 Or interiorNodeCount < 0 Then Err.Raise vbObjectError + 3406, "Mesh.ValidateMeshInputCapacity", "節点数が負です。"
    If boundaryNodeCount + interiorNodeCount > UBound(px) Then Err.Raise vbObjectError + 3407, "Mesh.ValidateMeshInputCapacity", "メッシュ節点配列の容量を超えています。"
    If boundaryNodeCount + interiorNodeCount > UBound(mtj, 1) Then Err.Raise vbObjectError + 3408, "Mesh.ValidateMeshInputCapacity", "メッシュ接続配列の容量を超えています。"
End Sub

Public Sub AssertMeshNodeCapacity(ByVal nextNode As Long)
    If nextNode < 1 Then Err.Raise vbObjectError + 3411, "Mesh.AssertMeshNodeCapacity", "生成節点番号が不正です。"
    If nextNode > UBound(px) Or nextNode > UBound(py) Or nextNode > UBound(ifix) Then _
        Err.Raise vbObjectError + 3412, "Mesh.AssertMeshNodeCapacity", "メッシュ生成中に節点固定配列の容量を超えました。"
End Sub

Public Sub AssertMeshElementCapacity(ByVal NextElement As Long)
    If NextElement < 1 Then Err.Raise vbObjectError + 3413, "Mesh.AssertMeshElementCapacity", "生成要素番号が不正です。"
    If NextElement > UBound(mtj, 1) Or NextElement > UBound(jac, 1) Or NextElement > UBound(idm) Then _
        Err.Raise vbObjectError + 3414, "Mesh.AssertMeshElementCapacity", "メッシュ生成中に要素固定配列の容量を超えました。"
End Sub

' P5現行Q8経路だけが使う作業領域。必要数だけ確保し、処理終了時に解放する。
Private Sub P5EnsureQ8RepairWorkspace(ByVal nodeCount As Long, ByVal elementCount As Long)
    Dim edgeCapacity As Long
    If nodeCount < 1 Or elementCount < 1 Then Err.Raise vbObjectError + 3500, "Mesh.P5EnsureQ8RepairWorkspace", "メッシュ数が0です。"
    edgeCapacity = elementCount * 4
    If edgeCapacity < 1 Then edgeCapacity = 1
    ReDim px(1 To nodeCount): ReDim py(1 To nodeCount): ReDim ifix(1 To nodeCount)
    ReDim mpx(1 To nodeCount): ReDim mpy(1 To nodeCount): ReDim xlok(1 To nodeCount)
    ReDim mtj(1 To elementCount, 1 To 9): ReDim jac(1 To elementCount, 1 To 4): ReDim idm(1 To elementCount)
    ReDim ihen(1 To edgeCapacity, 1 To 2): ReDim jhen(1 To edgeCapacity): ReDim khen(1 To edgeCapacity)
    ReDim jnb(1 To nodeCount): ReDim nei(1 To nodeCount, 1 To 100)
End Sub

' 生成済みの節点・要素配列を、容量が足りている限り再確保しない。
' デローニ経路の大容量作業配列をQ8品質検査で縮小すると、座標が消える。
Private Function P5Q8DataWorkspaceSufficient(ByVal nodeCount As Long, ByVal elementCount As Long) As Boolean
    On Error GoTo NotReady
    If UBound(px) < nodeCount Or UBound(py) < nodeCount Or UBound(ifix) < nodeCount Then Exit Function
    If UBound(mtj, 1) < elementCount Or UBound(mtj, 2) < 9 Then Exit Function
    If UBound(idm) < elementCount Then Exit Function
    P5Q8DataWorkspaceSufficient = True
    Exit Function
NotReady:
    P5Q8DataWorkspaceSufficient = False
End Function

Private Sub P5EnsureQ8AuxiliaryWorkspace(ByVal nodeCount As Long, ByVal elementCount As Long)
    Dim edgeCapacity As Long, needsReset As Boolean
    edgeCapacity = elementCount * 4
    If edgeCapacity < 1 Then edgeCapacity = 1
    On Error GoTo ResetAuxiliary
    If UBound(ihen, 1) < edgeCapacity Or UBound(ihen, 2) < 2 Then needsReset = True
    If UBound(jhen) < edgeCapacity Or UBound(khen) < edgeCapacity Then needsReset = True
    If UBound(jnb) < nodeCount Or UBound(nei, 1) < nodeCount Or UBound(nei, 2) < 100 Then needsReset = True
    If Not needsReset Then Exit Sub
ResetAuxiliary:
    ReDim ihen(1 To edgeCapacity, 1 To 2): ReDim jhen(1 To edgeCapacity): ReDim khen(1 To edgeCapacity)
    ReDim jnb(1 To nodeCount): ReDim nei(1 To nodeCount, 1 To 100)
End Sub

Private Sub P5ReleaseQ8RepairWorkspace()
    Erase px: Erase py: Erase ifix: Erase mpx: Erase mpy: Erase xlok
    Erase mtj: Erase jac: Erase idm: Erase mmtj: Erase smtj: Erase pmtj
    Erase ihen: Erase jhen: Erase khen: Erase jnb: Erase nei
End Sub

' 固定配列時のEraseによるゼロ化を、動的配列を解放せずに再現する。
Private Sub P5ZeroLongArray(ByRef values() As Long)
    Dim i As Long
    On Error GoTo Done
    For i = LBound(values) To UBound(values)
        values(i) = 0
    Next i
Done:
End Sub

' 要素作成②（図形定義→デローニ分割）専用の一時作業領域。
' 現行の①/解析経路の配列容量を兼用せず、実行時だけ確保して終了時に解放する。
Public Sub P5PrepareGeometryDefinitionWorkspace()
    Dim ws As Worksheet
    Dim headerValues As Variant, boundaryValues As Variant
    Dim i As Long, boundaryCount As Long, maxBoundaryNodes As Long
    Dim externalCount As Long, internalCount As Long, coordinateNodeCount As Long
    Dim nodeCapacity As Long, elementCapacity As Long, adjacencyNodeCapacity As Long
    Dim candidateCount As Long

    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Worksheets("図形定義")
    headerValues = ws.range(ws.Cells(1, 1), ws.Cells(24, 14)).value2
    externalCount = CLng(Val(headerValues(3, 4)))
    internalCount = CLng(Val(headerValues(24, 4)))
    If externalCount < 1 Then Err.Raise vbObjectError + 3470, "Mesh.P5PrepareGeometryDefinitionWorkspace", "図形定義の外部境界数が1未満です。"
    If internalCount < 0 Then Err.Raise vbObjectError + 3471, "Mesh.P5PrepareGeometryDefinitionWorkspace", "図形定義の内部境界数が負です。"
    boundaryCount = externalCount + internalCount
    coordinateNodeCount = CLng(Val(headerValues(1, 8))) + CLng(Val(headerValues(1, 10)))
    If coordinateNodeCount < 3 Then Err.Raise vbObjectError + 3472, "Mesh.P5PrepareGeometryDefinitionWorkspace", "図形定義の節点数が3未満です。"
    If coordinateNodeCount > 60000 Then Err.Raise vbObjectError + 3473, "Mesh.P5PrepareGeometryDefinitionWorkspace", "図形定義の節点数が容量を超えています。"

    ' 図形定義の境界節点リストはE列から最大2000節点分を一括で読む。
    boundaryValues = ws.range(ws.Cells(1, 5), ws.Cells(23 + internalCount, 5)).value2
    For i = 1 To boundaryCount
        If i <= externalCount Then
            candidateCount = CLng(Val(boundaryValues(2 + i, 1)))
        Else
            candidateCount = CLng(Val(boundaryValues(23 + i - externalCount, 1)))
        End If
        If candidateCount > maxBoundaryNodes Then maxBoundaryNodes = candidateCount
    Next i
    If maxBoundaryNodes < 3 Then Err.Raise vbObjectError + 3474, "Mesh.P5PrepareGeometryDefinitionWorkspace", "図形定義の境界節点数が3未満です。"
    If maxBoundaryNodes > 1999 Then Err.Raise vbObjectError + 3475, "Mesh.P5PrepareGeometryDefinitionWorkspace", "図形定義の境界節点数が旧デローニ経路の容量を超えています。"

    ' TRMODEL/ROUGH/TRFINEは内部で2001～2003の一時節点を使用し、
    ' 最悪時の要素数も含めて旧経路の上限を明示的に確保する。
    nodeCapacity = 60000
    elementCapacity = 60000
    adjacencyNodeCapacity = 2000

    ReDim ibex(1 To IIf(externalCount > 0, externalCount, 1))
    ReDim ibin(1 To IIf(internalCount > 0, internalCount, 1))
    ReDim mibex(1 To IIf(externalCount > 0, externalCount, 1))
    ReDim mdelx(1 To IIf(boundaryCount > 0, boundaryCount, 1))
    ReDim delx(1 To IIf(boundaryCount > 0, boundaryCount, 1))
    ReDim sbex(1 To IIf(externalCount > 0, externalCount, 1))
    ReDim mibno(1 To IIf(boundaryCount > 0, boundaryCount, 1), 1 To 2000)
    ReDim ibno(1 To IIf(boundaryCount > 0, boundaryCount, 1), 1 To 2000)

    ReDim mindex(1 To externalCount + 2)
    ReDim mtj(1 To elementCapacity, 1 To 9)
    ReDim jac(1 To elementCapacity, 1 To 4)
    ReDim idm(1 To elementCapacity)
    ReDim mmtj(1 To elementCapacity, 1 To 4)
    ReDim smtj(1 To elementCapacity, 1 To 4)
    ReDim pmtj(1 To elementCapacity, 1 To 4)
    ReDim kv(1 To elementCapacity)
    ReDim id(1 To elementCapacity)
    ReDim map(1 To elementCapacity)
    ReDim px(1 To nodeCapacity)
    ReDim py(1 To nodeCapacity)
    ReDim sx(1 To nodeCapacity)
    ReDim sy(1 To nodeCapacity)
    ReDim mpx(1 To nodeCapacity)
    ReDim mpy(1 To nodeCapacity)
    ReDim xlok(1 To nodeCapacity)
    ReDim ifix(1 To nodeCapacity)
    ReDim list(1 To nodeCapacity)
    ReDim istack(1 To nodeCapacity)
    ReDim ibr(1 To nodeCapacity)
    ReDim iadres(1 To nodeCapacity)
    ReDim ihen(1 To elementCapacity, 1 To 2)
    ReDim jhen(1 To elementCapacity)
    ReDim khen(1 To elementCapacity)
    ReDim jnb(1 To adjacencyNodeCapacity)
    ReDim nei(1 To adjacencyNodeCapacity, 1 To 100)
    Exit Sub
ErrorHandler:
    P5ReleaseGeometryDefinitionWorkspace
    Err.Raise Err.Number, Err.source, Err.Description
End Sub

Public Sub P5ReleaseGeometryDefinitionWorkspace()
    Erase ibex: Erase ibin: Erase mibex: Erase mdelx: Erase ibno: Erase mibno
    Erase mindex: Erase mtj: Erase delx: Erase px: Erase py: Erase sbex: Erase sx: Erase sy
    Erase mpx: Erase mpy: Erase xlok: Erase ifix: Erase jac: Erase idm: Erase mmtj: Erase smtj: Erase pmtj
    Erase kv: Erase id: Erase map: Erase ihen: Erase jhen: Erase khen: Erase jnb: Erase nei
    Erase list: Erase istack: Erase ibr: Erase iadres
End Sub

Private Function P5Q8WorkspaceMatches(ByVal nodeCount As Long, ByVal elementCount As Long) As Boolean
    On Error GoTo NotReady
    If UBound(px) <> nodeCount Or UBound(py) <> nodeCount Or UBound(ifix) <> nodeCount Then Exit Function
    If UBound(mtj, 1) <> elementCount Or UBound(mtj, 2) <> 9 Then Exit Function
    If UBound(jac, 1) <> elementCount Or UBound(jac, 2) <> 4 Then Exit Function
    If UBound(idm) <> elementCount Then Exit Function
    If UBound(jnb) <> nodeCount Or UBound(nei, 1) <> nodeCount Then Exit Function
    P5Q8WorkspaceMatches = True
    Exit Function
NotReady:
    P5Q8WorkspaceMatches = False
End Function

' P5: 要素・節点・辺・近傍・探索スタックの必要容量を一括確認する。
Public Sub P5ValidateMeshCapacity(ByVal nodeCount As Long, ByVal elementCount As Long)
    Dim edgeCount As Long
    If Not P5Q8DataWorkspaceSufficient(nodeCount, elementCount) Then
        P5EnsureQ8RepairWorkspace nodeCount, elementCount
    Else
        P5EnsureQ8AuxiliaryWorkspace nodeCount, elementCount
    End If
    If nodeCount < 1 Or nodeCount > UBound(px) Or nodeCount > UBound(py) Or nodeCount > UBound(ifix) Then _
        Err.Raise vbObjectError + 3501, "Mesh.P5ValidateMeshCapacity", "節点数が容量範囲外です。必要数=" & CStr(nodeCount) & "、容量=" & CStr(UBound(px))
    If elementCount < 1 Or elementCount > UBound(mtj, 1) Or elementCount > UBound(jac, 1) Or elementCount > UBound(idm) Then _
        Err.Raise vbObjectError + 3502, "Mesh.P5ValidateMeshCapacity", "要素数が容量範囲外です。必要数=" & CStr(elementCount) & "、容量=" & CStr(UBound(mtj, 1))
    If elementCount > UBound(ihen, 1) \ 4 Then _
        Err.Raise vbObjectError + 3503, "Mesh.P5ValidateMeshCapacity", "辺対応配列の必要容量を超えます。必要辺数上限=" & CStr(elementCount * 4) & "、容量=" & CStr(UBound(ihen, 1))
    edgeCount = elementCount * 4
    If edgeCount > UBound(jhen) Or edgeCount > UBound(khen) Then _
        Err.Raise vbObjectError + 3504, "Mesh.P5ValidateMeshCapacity", "辺の要素対応配列の容量を超えます。"
    If nodeCount > UBound(jnb) Or nodeCount > UBound(nei, 1) Then _
        Err.Raise vbObjectError + 3505, "Mesh.P5ValidateMeshCapacity", "節点近傍配列の容量を超えます。"
End Sub

' 外形は時計回り、穴は反時計回り。旧デローニの idm 判定(mp>0)がこの向きを前提にする。
' 図形定義が逆でもここで揃える。出力四角形は QUDATA で反時計回りへ戻す。
Private Sub P5NormalizeBoundaryWinding(ByVal externalCount As Long, ByVal internalCount As Long, _
                                       ByRef externalNodeCounts() As Long, ByRef internalNodeCounts() As Long, _
                                       ByRef boundaryNodes() As Long, ByRef x() As Double, ByRef y() As Double)
    Dim i As Long, loopCount As Long
    For i = 1 To externalCount
        loopCount = externalNodeCounts(i)
        If P5LoopSignedArea(i, loopCount, boundaryNodes, x, y) > 0# Then
            P5ReverseNodeLoop i, loopCount, boundaryNodes
        End If
    Next i
    For i = 1 To internalCount
        loopCount = internalNodeCounts(i)
        If P5LoopSignedArea(externalCount + i, loopCount, boundaryNodes, x, y) < 0# Then
            P5ReverseNodeLoop externalCount + i, loopCount, boundaryNodes
        End If
    Next i
End Sub

Private Function P5LoopSignedArea(ByVal boundaryIndex As Long, ByVal loopCount As Long, _
                                  ByRef boundaryNodes() As Long, ByRef x() As Double, ByRef y() As Double) As Double
    Dim j As Long, nextJ As Long, nodeA As Long, nodeB As Long
    P5LoopSignedArea = 0#
    If loopCount < 3 Then Exit Function
    For j = 1 To loopCount
        nextJ = j + 1
        If nextJ > loopCount Then nextJ = 1
        nodeA = boundaryNodes(boundaryIndex, j)
        nodeB = boundaryNodes(boundaryIndex, nextJ)
        If nodeA >= LBound(x) And nodeA <= UBound(x) And nodeB >= LBound(x) And nodeB <= UBound(x) Then
            P5LoopSignedArea = P5LoopSignedArea + x(nodeA) * y(nodeB) - x(nodeB) * y(nodeA)
        End If
    Next j
End Function

Private Sub P5ReverseNodeLoop(ByVal boundaryIndex As Long, ByVal loopCount As Long, ByRef boundaryNodes() As Long)
    Dim j As Long, swapValue As Long
    For j = 1 To loopCount \ 2
        swapValue = boundaryNodes(boundaryIndex, j)
        boundaryNodes(boundaryIndex, j) = boundaryNodes(boundaryIndex, loopCount - j + 1)
        boundaryNodes(boundaryIndex, loopCount - j + 1) = swapValue
    Next j
End Sub

' P5: 自動メッシュの境界入力を、シート出力の消去より前に検証する。
Public Sub P5ValidateBoundaryGeometry(ByVal externalCount As Long, ByVal internalCount As Long, _
                                       ByRef externalNodeCounts() As Long, ByRef internalNodeCounts() As Long, _
                                       ByRef boundaryNodes() As Long, ByRef x() As Double, ByRef y() As Double, _
                                       ByVal coordinateNodeCount As Long)
    Dim totalBoundary As Long, i As Long, j As Long, k As Long, l As Long
    Dim loopCount As Long, nodeA As Long, nodeB As Long, nodeC As Long, nodeD As Long
    Dim modelMinX As Double, modelMaxX As Double, modelMinY As Double, modelMaxY As Double
    Dim firstPoint As Boolean, pointX As Double, pointY As Double, insideOuter As Boolean
    Dim seen As Object, edgeLength As Double, nextJ As Long, nextK As Long, nextL As Long

    If externalCount < 1 Then Err.Raise vbObjectError + 3450, "Mesh.P5ValidateBoundaryGeometry", "外部境界は1つ以上必要です。"
    If internalCount < 0 Then Err.Raise vbObjectError + 3451, "Mesh.P5ValidateBoundaryGeometry", "内部境界数が負です。"
    totalBoundary = externalCount + internalCount
    If totalBoundary > UBound(boundaryNodes, 1) Then Err.Raise vbObjectError + 3452, "Mesh.P5ValidateBoundaryGeometry", "境界配列の容量を超えています。"
    firstPoint = True
    For i = 1 To coordinateNodeCount
        If firstPoint Then
            modelMinX = x(i): modelMaxX = x(i): modelMinY = y(i): modelMaxY = y(i): firstPoint = False
        Else
            If x(i) < modelMinX Then modelMinX = x(i)
            If x(i) > modelMaxX Then modelMaxX = x(i)
            If y(i) < modelMinY Then modelMinY = y(i)
            If y(i) > modelMaxY Then modelMaxY = y(i)
        End If
    Next i
    P5ModelScale = modelMaxX - modelMinX
    If modelMaxY - modelMinY > P5ModelScale Then P5ModelScale = modelMaxY - modelMinY
    If P5ModelScale <= 0# Then Err.Raise vbObjectError + 3453, "Mesh.P5ValidateBoundaryGeometry", "境界座標からモデル寸法を決定できません。"
    P5GeometryTolerance = P5ModelScale * 0.0000000001
    If P5GeometryTolerance < 0.000000000001 Then P5GeometryTolerance = 0.000000000001
    P5AreaTolerance = P5ModelScale * P5ModelScale * 0.000000000001
    If P5AreaTolerance < 1E-24 Then P5AreaTolerance = 1E-24
    P5BoundarySelfIntersectionCount = 0

    For i = 1 To totalBoundary
        If i <= externalCount Then loopCount = externalNodeCounts(i) Else loopCount = internalNodeCounts(i - externalCount)
        If loopCount < 3 Then Err.Raise vbObjectError + 3454, "Mesh.P5ValidateBoundaryGeometry", "境界点数は3以上必要です。境界=" & CStr(i)
        If loopCount > UBound(boundaryNodes, 2) - 1 Then Err.Raise vbObjectError + 3455, "Mesh.P5ValidateBoundaryGeometry", "境界点数が配列容量を超えています。境界=" & CStr(i)
        Set seen = CreateObject("Scripting.Dictionary")
        For j = 1 To loopCount
            nodeA = boundaryNodes(i, j)
            If nodeA < 1 Or nodeA > coordinateNodeCount Then Err.Raise vbObjectError + 3456, "Mesh.P5ValidateBoundaryGeometry", "境界の節点番号が範囲外です。境界=" & CStr(i) & "、点=" & CStr(j)
            If seen.Exists(CStr(nodeA)) Then Err.Raise vbObjectError + 3457, "Mesh.P5ValidateBoundaryGeometry", "境界内で節点番号が重複しています。境界=" & CStr(i) & "、節点=" & CStr(nodeA)
            seen.Add CStr(nodeA), True
            nextJ = (j Mod loopCount) + 1
            nodeB = boundaryNodes(i, nextJ)
            edgeLength = Sqr((x(nodeB) - x(nodeA)) ^ 2 + (y(nodeB) - y(nodeA)) ^ 2)
            If edgeLength <= P5GeometryTolerance Then Err.Raise vbObjectError + 3458, "Mesh.P5ValidateBoundaryGeometry", "極短辺または未閉鎖境界です。境界=" & CStr(i) & "、辺=" & CStr(j)
        Next j
        For j = 1 To loopCount
            nextJ = (j Mod loopCount) + 1
            nodeA = boundaryNodes(i, j): nodeB = boundaryNodes(i, nextJ)
            For k = j + 1 To loopCount
                nextK = (k Mod loopCount) + 1
                If nextJ <> k And nextK <> j Then
                    nodeC = boundaryNodes(i, k): nodeD = boundaryNodes(i, nextK)
                    If P5SegmentsIntersect(x(nodeA), y(nodeA), x(nodeB), y(nodeB), x(nodeC), y(nodeC), x(nodeD), y(nodeD)) Then
                        P5BoundarySelfIntersectionCount = P5BoundarySelfIntersectionCount + 1
                        Err.Raise vbObjectError + 3459, "Mesh.P5ValidateBoundaryGeometry", "境界が自己交差しています。境界=" & CStr(i) & "、辺=" & CStr(j) & "と" & CStr(k)
                    End If
                End If
            Next k
        Next j
    Next i

    ' 旧形式の「図形定義」では、複数の外部境界を隣接領域の分割線として
    ' 使用するため、外部境界どうしが共有辺・接触点を持つことがある。
    ' 外部境界どうしの交差検査は旧メッシャーの許容範囲とし、各境界自身の
    ' 自己交差、および外部境界と内部境界の交差は引き続き検査する。
    For i = 1 To totalBoundary - 1
        loopCount = P5BoundaryNodeCount(i, externalCount, externalNodeCounts, internalNodeCounts)
        For j = 1 To loopCount
            nextJ = (j Mod loopCount) + 1
            nodeA = boundaryNodes(i, j): nodeB = boundaryNodes(i, nextJ)
            For k = i + 1 To totalBoundary
                If Not (i <= externalCount And k <= externalCount) Then
                    l = P5BoundaryNodeCount(k, externalCount, externalNodeCounts, internalNodeCounts)
                    For nextK = 1 To l
                        nextL = (nextK Mod l) + 1
                        nodeC = boundaryNodes(k, nextK): nodeD = boundaryNodes(k, nextL)
                        If P5SegmentsIntersect(x(nodeA), y(nodeA), x(nodeB), y(nodeB), x(nodeC), y(nodeC), x(nodeD), y(nodeD)) Then
                            Err.Raise vbObjectError + 3460, "Mesh.P5ValidateBoundaryGeometry", "外部境界と内部境界が交差または接触しています。境界=" & CStr(i) & "と" & CStr(k)
                        End If
                    Next nextK
                End If
            Next k
        Next j
    Next i

    For i = externalCount + 1 To totalBoundary
        nodeA = boundaryNodes(i, 1): pointX = x(nodeA): pointY = y(nodeA): insideOuter = False
        For j = 1 To externalCount
            If P5PointInBoundary(pointX, pointY, j, externalCount, externalNodeCounts, internalNodeCounts, boundaryNodes, x, y) Then insideOuter = True
        Next j
        If Not insideOuter Then Err.Raise vbObjectError + 3461, "Mesh.P5ValidateBoundaryGeometry", "内部境界が外部境界の外側にあります。境界=" & CStr(i)
        For j = externalCount + 1 To totalBoundary
            If j <> i Then
                If P5PointInBoundary(pointX, pointY, j, externalCount, externalNodeCounts, internalNodeCounts, boundaryNodes, x, y) Then
                    Err.Raise vbObjectError + 3491, "Mesh.P5ValidateBoundaryGeometry", "内部境界が別の穴の内側にあります。穴どうしを入れ子にすることはできません。"
                End If
            End If
        Next j
    Next i
End Sub

Private Function P5BoundaryNodeCount(ByVal boundaryIndex As Long, ByVal externalCount As Long, _
                                     ByRef externalNodeCounts() As Long, ByRef internalNodeCounts() As Long) As Long
    If boundaryIndex <= externalCount Then P5BoundaryNodeCount = externalNodeCounts(boundaryIndex) Else P5BoundaryNodeCount = internalNodeCounts(boundaryIndex - externalCount)
End Function

Private Function P5PointInBoundary(ByVal pointX As Double, ByVal pointY As Double, ByVal boundaryIndex As Long, _
                                   ByVal externalCount As Long, ByRef externalNodeCounts() As Long, ByRef internalNodeCounts() As Long, _
                                   ByRef boundaryNodes() As Long, ByRef x() As Double, ByRef y() As Double) As Boolean
    Dim count As Long, i As Long, nextI As Long, nodeA As Long, nodeB As Long
    Dim intersects As Boolean
    count = P5BoundaryNodeCount(boundaryIndex, externalCount, externalNodeCounts, internalNodeCounts)
    For i = 1 To count
        nextI = (i Mod count) + 1
        nodeA = boundaryNodes(boundaryIndex, i): nodeB = boundaryNodes(boundaryIndex, nextI)
        If ((y(nodeA) > pointY) <> (y(nodeB) > pointY)) Then
            If pointX < (x(nodeB) - x(nodeA)) * (pointY - y(nodeA)) / (y(nodeB) - y(nodeA)) + x(nodeA) Then intersects = Not intersects
        End If
    Next i
    P5PointInBoundary = intersects
End Function

Private Function P5SegmentsIntersect(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double, _
                                     ByVal cx As Double, ByVal cy As Double, ByVal dx As Double, ByVal dy As Double) As Boolean
    Dim crossTol As Double, o1 As Double, o2 As Double, o3 As Double, o4 As Double
    crossTol = P5GeometryTolerance * P5ModelScale
    If crossTol < 1E-24 Then crossTol = 1E-24
    o1 = (bx - ax) * (cy - ay) - (by - ay) * (cx - ax)
    o2 = (bx - ax) * (dy - ay) - (by - ay) * (dx - ax)
    o3 = (dx - cx) * (ay - cy) - (dy - cy) * (ax - cx)
    o4 = (dx - cx) * (by - cy) - (dy - cy) * (bx - cx)
    If Abs(o1) <= crossTol And P5OnSegment(ax, ay, bx, by, cx, cy) Then P5SegmentsIntersect = True: Exit Function
    If Abs(o2) <= crossTol And P5OnSegment(ax, ay, bx, by, dx, dy) Then P5SegmentsIntersect = True: Exit Function
    If Abs(o3) <= crossTol And P5OnSegment(cx, cy, dx, dy, ax, ay) Then P5SegmentsIntersect = True: Exit Function
    If Abs(o4) <= crossTol And P5OnSegment(cx, cy, dx, dy, bx, by) Then P5SegmentsIntersect = True: Exit Function
    P5SegmentsIntersect = ((o1 > crossTol And o2 < -crossTol) Or (o1 < -crossTol And o2 > crossTol)) And _
                          ((o3 > crossTol And o4 < -crossTol) Or (o3 < -crossTol And o4 > crossTol))
End Function

Private Function P5OnSegment(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double, ByVal pxValue As Double, ByVal pyValue As Double) As Boolean
    P5OnSegment = pxValue >= IIf(ax < bx, ax, bx) - P5GeometryTolerance And pxValue <= IIf(ax > bx, ax, bx) + P5GeometryTolerance And _
                  pyValue >= IIf(ay < by, ay, by) - P5GeometryTolerance And pyValue <= IIf(ay > by, ay, by) + P5GeometryTolerance
End Function

' P5: Q8要素の品質と接続グラフを、生成後・平滑化前後に確認する。
Public Sub P5ValidateMeshQuality(ByVal nodeCount As Long, ByVal elementCount As Long, _
                                  ByRef connectivity() As Long, ByRef x() As Double, ByRef y() As Double)
    Dim i As Long, firstPoint As Boolean
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    P5ValidateMeshCapacity nodeCount, elementCount
    firstPoint = True
    For i = 1 To nodeCount
        If firstPoint Then
            minX = x(i): maxX = x(i): minY = y(i): maxY = y(i): firstPoint = False
        Else
            If x(i) < minX Then minX = x(i)
            If x(i) > maxX Then maxX = x(i)
            If y(i) < minY Then minY = y(i)
            If y(i) > maxY Then maxY = y(i)
        End If
    Next i
    P5ModelScale = maxX - minX
    If maxY - minY > P5ModelScale Then P5ModelScale = maxY - minY
    If P5ModelScale <= 0# Then Err.Raise vbObjectError + 3462, "Mesh.P5ValidateMeshQuality", _
        "メッシュ寸法が0です。節点数=" & CStr(nodeCount) & "、X範囲=" & Format$(minX, "0.000E+00") & "～" & Format$(maxX, "0.000E+00") & _
        "、Y範囲=" & Format$(minY, "0.000E+00") & "～" & Format$(maxY, "0.000E+00")
    P5GeometryTolerance = P5ModelScale * 0.0000000001
    If P5GeometryTolerance < 0.000000000001 Then P5GeometryTolerance = 0.000000000001
    P5AreaTolerance = P5ModelScale * P5ModelScale * 0.000000000001
    If P5AreaTolerance < 1E-24 Then P5AreaTolerance = 1E-24
    P5MinElementArea = 1E+308: P5MinJacobian = 1E+308
    P5MaxAspectRatio = 0#: P5MinAngleDeg = 180#: P5MaxAngleDeg = 0#
    P5BoundaryEdgeCount = 0: P5SharedEdgeCount = 0: P5ConnectedComponentCount = 0
    For i = 1 To elementCount
        P5EvaluateMeshQuadQuality i, connectivity, x, y
    Next i
    P5ValidateMeshTopology nodeCount, elementCount, connectivity
End Sub

Private Sub P5EvaluateMeshQuadQuality(ByVal elementId As Long, ByRef connectivity() As Long, ByRef x() As Double, ByRef y() As Double)
    Dim nodeNo(1 To 8) As Long, edgeLength As Double, minEdge As Double, maxEdge As Double
    Dim area As Double, jacobianValue As Double, minJacobian As Double, angleDeg As Double
    Dim i As Long, nextIndex As Long, previousIndex As Long
    Dim dx1 As Double, dy1 As Double, dx2 As Double, dy2 As Double, dotValue As Double, crossValue As Double
    Dim xiValues(1 To 8) As Double, etaValues(1 To 8) As Double
    For i = 1 To 8
        nodeNo(i) = connectivity(elementId, i)
        If nodeNo(i) < 1 Or nodeNo(i) > UBound(x) Then Err.Raise vbObjectError + 3463, "Mesh.P5ValidateMeshQuality", "要素接続番号が範囲外です。要素=" & CStr(elementId)
    Next i
    For i = 1 To 7
        For nextIndex = i + 1 To 8
            If nodeNo(i) = nodeNo(nextIndex) Then Err.Raise vbObjectError + 3464, "Mesh.P5ValidateMeshQuality", "要素内で節点番号が重複しています。要素=" & CStr(elementId)
        Next nextIndex
    Next i
    area = 0.5 * ((x(nodeNo(1)) * y(nodeNo(2)) + x(nodeNo(2)) * y(nodeNo(3)) + x(nodeNo(3)) * y(nodeNo(4)) + x(nodeNo(4)) * y(nodeNo(1))) - _
                   (y(nodeNo(1)) * x(nodeNo(2)) + y(nodeNo(2)) * x(nodeNo(3)) + y(nodeNo(3)) * x(nodeNo(4)) + y(nodeNo(4)) * x(nodeNo(1))))
    If area <= P5AreaTolerance Then Err.Raise vbObjectError + 3465, "Mesh.P5ValidateMeshQuality", "要素面積が退化しているか、向きが不正です。要素=" & CStr(elementId)
    If area < P5MinElementArea Then P5MinElementArea = area
    minEdge = 1E+308: maxEdge = 0#
    For i = 1 To 4
        nextIndex = (i Mod 4) + 1: previousIndex = ((i + 2) Mod 4) + 1
        edgeLength = Sqr((x(nodeNo(nextIndex)) - x(nodeNo(i))) ^ 2 + (y(nodeNo(nextIndex)) - y(nodeNo(i))) ^ 2)
        If edgeLength <= P5GeometryTolerance Then Err.Raise vbObjectError + 3466, "Mesh.P5ValidateMeshQuality", "極短辺を持つ要素です。要素=" & CStr(elementId) & "、辺=" & CStr(i)
        If edgeLength < minEdge Then minEdge = edgeLength
        If edgeLength > maxEdge Then maxEdge = edgeLength
        dx1 = x(nodeNo(previousIndex)) - x(nodeNo(i)): dy1 = y(nodeNo(previousIndex)) - y(nodeNo(i))
        dx2 = x(nodeNo(nextIndex)) - x(nodeNo(i)): dy2 = y(nodeNo(nextIndex)) - y(nodeNo(i))
        dotValue = dx1 * dx2 + dy1 * dy2: crossValue = dx1 * dy2 - dy1 * dx2
        angleDeg = P5Atan2(Abs(crossValue), dotValue) * 180# / 3.14159265358979
        If angleDeg < P5MinAngleDeg Then P5MinAngleDeg = angleDeg
        If angleDeg > P5MaxAngleDeg Then P5MaxAngleDeg = angleDeg
        If angleDeg < 0.001 Or angleDeg > 179.999 Then Err.Raise vbObjectError + 3467, "Mesh.P5ValidateMeshQuality", "要素角度が退化しています。要素=" & CStr(elementId) & "、頂点=" & CStr(i)
    Next i
    If minEdge > 0# And maxEdge / minEdge > P5MaxAspectRatio Then P5MaxAspectRatio = maxEdge / minEdge
    xiValues(1) = -1#: etaValues(1) = -1#: xiValues(2) = 1#: etaValues(2) = -1#: xiValues(3) = 1#: etaValues(3) = 1#: xiValues(4) = -1#: etaValues(4) = 1#
    xiValues(5) = -0.577350269189626: etaValues(5) = -0.577350269189626: xiValues(6) = 0.577350269189626: etaValues(6) = -0.577350269189626
    xiValues(7) = 0.577350269189626: etaValues(7) = 0.577350269189626: xiValues(8) = -0.577350269189626: etaValues(8) = 0.577350269189626
    minJacobian = 1E+308
    For i = 1 To 8
        jacobianValue = P5MeshQuadJacobian(x, y, nodeNo, xiValues(i), etaValues(i))
        If jacobianValue < minJacobian Then minJacobian = jacobianValue
    Next i
    If minJacobian < P5MinJacobian Then P5MinJacobian = minJacobian
    If minJacobian <= P5AreaTolerance Then Err.Raise vbObjectError + 3468, "Mesh.P5ValidateMeshQuality", "要素Jacobianが正でないか小さすぎます。要素=" & CStr(elementId)
End Sub

Private Function P5MeshQuadJacobian(ByRef x() As Double, ByRef y() As Double, ByRef nodeNo() As Long, ByVal xi As Double, ByVal eta As Double) As Double
    Dim dNxi(1 To 8) As Double, dNeta(1 To 8) As Double
    Dim i As Long, dXdxi As Double, dYdxi As Double, dXdeta As Double, dYdeta As Double
    dNxi(1) = 0.25 * (1# - eta) * (2# * xi + eta)
    dNeta(1) = 0.25 * (1# - xi) * (xi + 2# * eta)
    dNxi(2) = 0.25 * (1# - eta) * (2# * xi - eta)
    dNeta(2) = 0.25 * (1# + xi) * (-xi + 2# * eta)
    dNxi(3) = 0.25 * (1# + eta) * (2# * xi + eta)
    dNeta(3) = 0.25 * (1# + xi) * (xi + 2# * eta)
    dNxi(4) = 0.25 * (1# + eta) * (2# * xi - eta)
    dNeta(4) = 0.25 * (1# - xi) * (-xi + 2# * eta)
    dNxi(5) = -xi * (1# - eta)
    dNeta(5) = -0.5 * (1# - xi * xi)
    dNxi(6) = 0.5 * (1# - eta * eta)
    dNeta(6) = -eta * (1# + xi)
    dNxi(7) = -xi * (1# + eta)
    dNeta(7) = 0.5 * (1# - xi * xi)
    dNxi(8) = -0.5 * (1# - eta * eta)
    dNeta(8) = -eta * (1# - xi)
    For i = 1 To 8
        dXdxi = dXdxi + dNxi(i) * x(nodeNo(i)): dYdxi = dYdxi + dNxi(i) * y(nodeNo(i))
        dXdeta = dXdeta + dNeta(i) * x(nodeNo(i)): dYdeta = dYdeta + dNeta(i) * y(nodeNo(i))
    Next i
    P5MeshQuadJacobian = dXdxi * dYdeta - dYdxi * dXdeta
End Function

Public Function P5QualityMeetsSettings() As Boolean
    Dim scaleSquared As Double, minAreaRatio As Double, minJacobianRatio As Double
    Dim minAreaRatioTarget As Double, minJacobianRatioTarget As Double
    Dim minAngleTarget As Double, maxAngleTarget As Double, maxAspectTarget As Double
    P5QualityMeetsSettings = False
    If P5ModelScale <= 0# Then Exit Function
    scaleSquared = P5ModelScale * P5ModelScale
    minAreaRatio = P5MinElementArea / scaleSquared
    minJacobianRatio = P5MinJacobian / scaleSquared
    minAreaRatioTarget = P5ReadSetting("MESH_QUALITY_MIN_AREA_RATIO", 0.000001)
    minJacobianRatioTarget = P5ReadSetting("MESH_QUALITY_MIN_JACOBIAN_RATIO", 0.000001)
    minAngleTarget = P5ReadSetting("MESH_QUALITY_MIN_ANGLE_DEG", 10#)
    maxAngleTarget = P5ReadSetting("MESH_QUALITY_MAX_ANGLE_DEG", 170#)
    maxAspectTarget = P5ReadSetting("MESH_QUALITY_MAX_ASPECT_RATIO", 5#)
    P5QualityMeetsSettings = minAreaRatio >= minAreaRatioTarget And _
                             minJacobianRatio >= minJacobianRatioTarget And _
                             P5MinAngleDeg >= minAngleTarget And _
                             P5MaxAngleDeg <= maxAngleTarget And _
                             P5MaxAspectRatio <= maxAspectTarget And _
                             P5IsolatedNodeCount = 0 And P5ConnectedComponentCount = 1
End Function

Public Function P5QualitySummary() As String
    P5QualitySummary = "最小面積=" & Format$(P5MinElementArea, "0.000E+00") & _
                       "、最小Jacobian=" & Format$(P5MinJacobian, "0.000E+00") & _
                       "、最小角度=" & Format$(P5MinAngleDeg, "0.0") & "°" & _
                       "、最大アスペクト比=" & Format$(P5MaxAspectRatio, "0.00")
End Function

Public Sub P5WriteMeshQualitySettings(ByVal routeName As String, ByVal repairStatus As String, _
                                      ByVal nodeCount As Long, ByVal elementCount As Long, _
                                      ByVal iterationCount As Long, ByVal movedNodeCount As Long)
    P5WriteTextSetting "MESH_QUALITY_LAST_STATUS", IIf(P5QualityMeetsSettings(), "PASS", "WARN")
    P5WriteTextSetting "MESH_QUALITY_LAST_ROUTE", routeName
    P5WriteTextSetting "MESH_QUALITY_LAST_REPAIR_STATUS", repairStatus
    P5WriteNumericSetting "MESH_QUALITY_LAST_NODE_COUNT", CDbl(nodeCount)
    P5WriteNumericSetting "MESH_QUALITY_LAST_ELEMENT_COUNT", CDbl(elementCount)
    P5WriteNumericSetting "MESH_QUALITY_LAST_MIN_ELEMENT_AREA", P5MinElementArea
    P5WriteNumericSetting "MESH_QUALITY_LAST_MIN_JACOBIAN", P5MinJacobian
    P5WriteNumericSetting "MESH_QUALITY_LAST_MIN_ANGLE_DEG", P5MinAngleDeg
    P5WriteNumericSetting "MESH_QUALITY_LAST_MAX_ANGLE_DEG", P5MaxAngleDeg
    P5WriteNumericSetting "MESH_QUALITY_LAST_MAX_ASPECT_RATIO", P5MaxAspectRatio
    P5WriteNumericSetting "MESH_QUALITY_LAST_BOUNDARY_EDGES", CDbl(P5BoundaryEdgeCount)
    P5WriteNumericSetting "MESH_QUALITY_LAST_SHARED_EDGES", CDbl(P5SharedEdgeCount)
    P5WriteNumericSetting "MESH_QUALITY_LAST_ISOLATED_NODES", CDbl(P5IsolatedNodeCount)
    P5WriteNumericSetting "MESH_QUALITY_LAST_COMPONENTS", CDbl(P5ConnectedComponentCount)
    P5WriteNumericSetting "MESH_QUALITY_LAST_REPAIR_ITERATIONS", CDbl(iterationCount)
    P5WriteNumericSetting "MESH_QUALITY_LAST_MOVED_NODES", CDbl(movedNodeCount)
End Sub

Public Sub P5EvaluateCurrentSheet(ByVal routeName As String)
    Dim localNodeCount As Long, localElementCount As Long
    SQINPUT localNodeCount, localElementCount, mtj, px, py, ifix
    P5WriteMeshQualitySettings routeName, "NONE", localNodeCount, localElementCount, 0, 0
    P5ReleaseQ8RepairWorkspace
End Sub

Private Function P5Atan2(ByVal yValue As Double, ByVal xValue As Double) As Double
    If Abs(xValue) <= 1E-30 Then
        If yValue > 0# Then
            P5Atan2 = 1.5707963267949
        ElseIf yValue < 0# Then
            P5Atan2 = -1.5707963267949
        Else
            P5Atan2 = 0#
        End If
    Else
        P5Atan2 = Atn(yValue / xValue)
        If xValue < 0# Then
            If yValue >= 0# Then
                P5Atan2 = P5Atan2 + 3.14159265358979
            Else
                P5Atan2 = P5Atan2 - 3.14159265358979
            End If
        End If
    End If
End Function

Public Sub P5ValidateMeshTopology(ByVal nodeCount As Long, ByVal elementCount As Long, ByRef connectivity() As Long)
    Dim edgeMap As Object, parent() As Long, usedNode() As Boolean
    Dim i As Long, j As Long, a As Long, b As Long, edgeCount As Long, key As String
    Dim info As Variant
    Set edgeMap = CreateObject("Scripting.Dictionary"): edgeMap.CompareMode = 0
    ReDim parent(1 To nodeCount): ReDim usedNode(1 To nodeCount)
    For i = 1 To nodeCount: parent(i) = i: Next i
    For i = 1 To elementCount
        For j = 1 To 4
            a = connectivity(i, j): b = connectivity(i, (j Mod 4) + 1)
            If a < 1 Or a > nodeCount Or b < 1 Or b > nodeCount Then Err.Raise vbObjectError + 3469, "Mesh.P5ValidateMeshTopology", "要素接続番号が範囲外です。要素=" & CStr(i)
            usedNode(a) = True: usedNode(b) = True: P5MeshUnion parent, a, b
            If a < b Then key = CStr(a) & ":" & CStr(b) Else key = CStr(b) & ":" & CStr(a)
            If edgeMap.Exists(key) Then
                info = Split(CStr(edgeMap.item(key)), "|"): edgeCount = CLng(info(0))
                If edgeCount >= 2 Then Err.Raise vbObjectError + 3470, "Mesh.P5ValidateMeshTopology", "辺を3要素以上で共有しています。辺=" & key
                P5SharedEdgeCount = P5SharedEdgeCount + 1
                P5BoundaryEdgeCount = P5BoundaryEdgeCount - 1
                edgeMap.item(key) = CStr(edgeCount + 1)
            Else
                edgeMap.Add key, "1": P5BoundaryEdgeCount = P5BoundaryEdgeCount + 1
            End If
        Next j
        For j = 5 To 8
            a = connectivity(i, j)
            If j = 5 Then b = connectivity(i, 1)
            If j = 6 Then b = connectivity(i, 2)
            If j = 7 Then b = connectivity(i, 3)
            If j = 8 Then b = connectivity(i, 4)
            If a < 1 Or a > nodeCount Then Err.Raise vbObjectError + 3470, "Mesh.P5ValidateMeshTopology", "Q8中間節点番号が範囲外です。要素=" & CStr(i)
            usedNode(a) = True
            P5MeshUnion parent, a, b
        Next j
    Next i
    For i = 1 To nodeCount
        If Not usedNode(i) Then P5IsolatedNodeCount = P5IsolatedNodeCount + 1
        If P5MeshFindRoot(parent, i) = i Then P5ConnectedComponentCount = P5ConnectedComponentCount + 1
    Next i
    If P5IsolatedNodeCount > 0 Then Err.Raise vbObjectError + 3471, "Mesh.P5ValidateMeshTopology", "要素に接続していない節点があります。件数=" & CStr(P5IsolatedNodeCount)
    If P5ConnectedComponentCount > 1 Then Err.Raise vbObjectError + 3472, "Mesh.P5ValidateMeshTopology", "メッシュが複数の非接続領域に分かれています。領域数=" & CStr(P5ConnectedComponentCount)
End Sub

Private Sub P5MeshUnion(ByRef parent() As Long, ByVal leftNode As Long, ByVal rightNode As Long)
    Dim leftRoot As Long, rightRoot As Long
    leftRoot = P5MeshFindRoot(parent, leftNode): rightRoot = P5MeshFindRoot(parent, rightNode)
    If leftRoot <> rightRoot Then parent(rightRoot) = leftRoot
End Sub

Private Function P5MeshFindRoot(ByRef parent() As Long, ByVal nodeNo As Long) As Long
    Dim rootNo As Long
    rootNo = nodeNo
    Do While parent(rootNo) <> rootNo: rootNo = parent(rootNo): Loop
    Do While parent(nodeNo) <> nodeNo
        rootNo = parent(nodeNo): parent(nodeNo) = P5MeshFindRoot(parent, rootNo): nodeNo = rootNo
    Loop
    P5MeshFindRoot = rootNo
End Function

Private Function P5MeshQuadMetrics(ByVal elementNo As Long, ByRef connectivity() As Long, ByRef x() As Double, ByRef y() As Double, _
                                   ByRef area As Double, ByRef minJacobian As Double, ByRef aspectRatio As Double, _
                                   ByRef MinAngle As Double, ByRef MaxAngle As Double) As Boolean
    Dim nodeNo(1 To 8) As Long, edgeLength As Double, minEdge As Double, maxEdge As Double
    Dim i As Long, nextIndex As Long, previousIndex As Long, jacobianValue As Double
    Dim dx1 As Double, dy1 As Double, dx2 As Double, dy2 As Double, dotValue As Double, crossValue As Double, angleDeg As Double
    Dim xiValues(1 To 8) As Double, etaValues(1 To 8) As Double
    P5MeshQuadMetrics = False
    MaxAngle = 0#
    For i = 1 To 8
        nodeNo(i) = connectivity(elementNo, i)
        If nodeNo(i) < 1 Or nodeNo(i) > UBound(x) Then Exit Function
    Next i
    For i = 1 To 7
        For nextIndex = i + 1 To 8
            If nodeNo(i) = nodeNo(nextIndex) Then Exit Function
        Next nextIndex
    Next i
    area = 0.5 * ((x(nodeNo(1)) * y(nodeNo(2)) + x(nodeNo(2)) * y(nodeNo(3)) + x(nodeNo(3)) * y(nodeNo(4)) + x(nodeNo(4)) * y(nodeNo(1))) - _
                   (y(nodeNo(1)) * x(nodeNo(2)) + y(nodeNo(2)) * x(nodeNo(3)) + y(nodeNo(3)) * x(nodeNo(4)) + y(nodeNo(4)) * x(nodeNo(1))))
    If area <= P5AreaTolerance Then Exit Function
    minEdge = 1E+308: maxEdge = 0#
    For i = 1 To 4
        nextIndex = (i Mod 4) + 1
        edgeLength = Sqr((x(nodeNo(nextIndex)) - x(nodeNo(i))) ^ 2 + (y(nodeNo(nextIndex)) - y(nodeNo(i))) ^ 2)
        If edgeLength <= P5GeometryTolerance Then Exit Function
        If edgeLength < minEdge Then minEdge = edgeLength
        If edgeLength > maxEdge Then maxEdge = edgeLength
        previousIndex = ((i + 2) Mod 4) + 1
        dx1 = x(nodeNo(previousIndex)) - x(nodeNo(i)): dy1 = y(nodeNo(previousIndex)) - y(nodeNo(i))
        dx2 = x(nodeNo(nextIndex)) - x(nodeNo(i)): dy2 = y(nodeNo(nextIndex)) - y(nodeNo(i))
        dotValue = dx1 * dx2 + dy1 * dy2: crossValue = dx1 * dy2 - dy1 * dx2
        angleDeg = P5Atan2(Abs(crossValue), dotValue) * 180# / 3.14159265358979
        If i = 1 Or angleDeg < MinAngle Then MinAngle = angleDeg
        If i = 1 Or angleDeg > MaxAngle Then MaxAngle = angleDeg
        If angleDeg < 0.001 Or angleDeg > 179.999 Then Exit Function
    Next i
    aspectRatio = maxEdge / minEdge
    xiValues(1) = -1#: etaValues(1) = -1#: xiValues(2) = 1#: etaValues(2) = -1#: xiValues(3) = 1#: etaValues(3) = 1#: xiValues(4) = -1#: etaValues(4) = 1#
    xiValues(5) = -0.577350269189626: etaValues(5) = -0.577350269189626: xiValues(6) = 0.577350269189626: etaValues(6) = -0.577350269189626
    xiValues(7) = 0.577350269189626: etaValues(7) = 0.577350269189626: xiValues(8) = -0.577350269189626: etaValues(8) = 0.577350269189626
    minJacobian = 1E+308
    For i = 1 To 8
        jacobianValue = P5MeshQuadJacobian(x, y, nodeNo, xiValues(i), etaValues(i))
        If jacobianValue < minJacobian Then minJacobian = jacobianValue
    Next i
    If minJacobian <= P5AreaTolerance Then Exit Function
    P5MeshQuadMetrics = True
End Function

Sub CER()


    Erase ibex
    Erase ibin
    Erase mibex
    Erase mdelx
    Erase ibno
    Erase mibno
    
    Erase mindex
    Erase mtj
    Erase delx
    NEX = 0
    NIN = 0
    NOB = 0
    NIB = 0
    node = 0
    NELM = 0
    ntp = 0
    ia = 0
    ib = 0
    ic = 0
    meshXmin = 0#
    meshXmax = 0#
    meshYmin = 0#
    meshYmax = 0#
    rax = 0#
    ray = 0#
    dMax = 0#
    Erase list
    Erase istack
    Erase ibr
    Erase iadres
    Erase sx
    Erase sy
    Erase px
    Erase py
    Erase mpx
    Erase mpy
    Erase xlok

    Erase ifix
    Erase jac
    Erase idm
    Erase mmtj
    Erase smtj
    Erase pmtj
    
    Erase kv
    Erase id
    Erase map
    Erase ihen
    Erase jhen
    Erase khen
    Erase jnb
    Erase nei


End Sub
Sub PINPUT(NEX As Long, NIN As Long, ibex() As Long, ibin() As Long, ibno() As Long, NOB As Long, NIB As Long, _
            px() As Double, py() As Double, delx() As Double, xlok() As Long)
Dim ddx As Variant: Dim dx As Variant: Dim flg As Variant: Dim je As Variant: Dim lx As Variant
Dim newno As Variant: Dim no1 As Variant: Dim no2 As Variant: Dim rmax As Variant: Dim sa As Variant
Dim headerValues As Variant: Dim nodeCoordValues As Variant: Dim boundaryValues As Variant
Dim wsClear As Worksheet: Dim lastClearRow As Long
Dim pinStage As String: Dim pinErrorNumber As Long: Dim pinErrorDescription As String

    Dim ws0 As Worksheet
    Dim i As Long, j As Long, k As Long
    Dim xc As Double
    Dim r As Long
    
    Dim alpha As Double
    Dim imno() As Double
    Dim deno() As Long
    Dim sms() As Long
    
    Dim NMAX As Long
    On Error GoTo PinputErrorHandler
    Set ws0 = ThisWorkbook.Sheets("図形定義")


    
    pinStage = "ヘッダー読込"
    ' P4: 入力範囲を一括読込みし、セル単位の往復を避ける
    headerValues = ws0.range(ws0.Cells(1, 1), ws0.Cells(24, 14)).value2
    alpha = CDbl(Val(headerValues(1, 14)))
    NOB = CLng(Val(headerValues(1, 8)))
    NIB = CLng(Val(headerValues(1, 10)))
    NEX = CLng(Val(headerValues(3, 4)))
    NIN = CLng(Val(headerValues(24, 4)))
    ReDim imno(1 To IIf(NOB + NIB > 200, NOB + NIB, 200), 1 To IIf(NOB + NIB > 200, NOB + NIB, 200))
    ReDim deno(1 To IIf(NOB + NIB > 200, NOB + NIB, 200), 1 To IIf(NOB + NIB > 200, NOB + NIB, 200))
    ReDim sms(1 To IIf(NOB + NIB > 200, NOB + NIB, 200), 1 To IIf(NOB + NIB > 200, NOB + NIB, 200))
    ValidateMeshInputCapacity NEX, NIN, NOB, NIB
    pinStage = "節点座標読込"
    nodeCoordValues = ws0.range(ws0.Cells(2, 2), ws0.Cells(NOB + NIB + 1, 3)).value2
    For i = 1 To NOB + NIB
        px(i) = CDbl(Val(nodeCoordValues(i, 1))): py(i) = CDbl(Val(nodeCoordValues(i, 2)))
        sx(i) = px(i): sy(i) = py(i)
    Next i
    pinStage = "境界配列読込"
    boundaryValues = ws0.range(ws0.Cells(1, 5), ws0.Cells(23 + NIN, P5BoundaryReadWidth(ws0, NEX, NIN))).value2
    

    '---------------------------------読み込み
    
    ' Reading values for ibex and delx if NEX > 0
    If NEX > 0 Then
        For i = 1 To NEX
            ibex(i) = CLng(Val(boundaryValues(2 + i, 1)))
            sbex(i) = ibex(i)
            delx(i) = CDbl(Val(boundaryValues(2 + i, 2)))
            For j = 1 To ibex(i)
                mibno(i, j) = CLng(Val(boundaryValues(2 + i, 2 + j)))
            Next j
        Next i
    End If
    If NIN > 0 Then
        
        For i = 1 To NIN
            ibin(i) = CLng(Val(boundaryValues(23 + i, 1)))
            delx(i + NEX) = CDbl(Val(boundaryValues(23 + i, 2)))
            For j = 1 To ibin(i)
                mibno(i + NEX, j) = CLng(Val(boundaryValues(23 + i, 2 + j)))
            Next j
        Next i
        
    End If
    pinStage = "境界向き"
    P5NormalizeBoundaryWinding NEX, NIN, ibex, ibin, mibno, px, py
    pinStage = "境界検証"
    P5ValidateBoundaryGeometry NEX, NIN, ibex, ibin, mibno, px, py, NOB + NIB

    ' P5: 境界検証に合格した後で初めて既存の生成結果を消去する。
    Set wsClear = ThisWorkbook.Worksheets("節点データ")
    lastClearRow = wsClear.Cells(wsClear.rows.count, 1).End(xlUp).row
    If lastClearRow >= 2 Then wsClear.range(wsClear.Cells(2, 1), wsClear.Cells(lastClearRow, 14)).ClearContents
    Set wsClear = ThisWorkbook.Worksheets("要素データ")
    lastClearRow = wsClear.Cells(wsClear.rows.count, 1).End(xlUp).row
    If lastClearRow >= 2 Then wsClear.range(wsClear.Cells(2, 1), wsClear.Cells(lastClearRow, 6)).ClearContents
    pinStage = "中間点配列構築"
    '---------------------------------中間点の把握
    For i = 1 To (NEX + NIN)
        
        If i > NEX Then
            je = ibin(i - NEX) + 1
        Else
            je = ibex(i) + 1
        End If
        
        For j = 2 To je
            no1 = mibno(i, j - 1) '節点番号と節点番号
            If je = j Then
               no2 = mibno(i, 1)
            Else
               no2 = mibno(i, j)
            End If

            dx = imno(no1, no2)
            If dx = 0 Or delx(i) < dx Then
                imno(no1, no2) = delx(i) 'imnoには点間隔を入力する
            End If
            dx = imno(no2, no1)
            If dx = 0 Or delx(i) < dx Then
                imno(no2, no1) = delx(i) 'imnoには点間隔を入力する
            End If
            If imno(no1, no2) > imno(no2, no1) Then
               imno(no1, no2) = imno(no2, no1)
            Else
               imno(no2, no1) = imno(no1, no2)
            End If
        Next j
    Next i
    
    
    
    
    
    newno = NOB + NIB 'NOB + NIB
    ' 固定配列時のEraseはゼロ化でしたが、動的配列では未確保化になります。
    ' 後段で再利用するため、元節点数の範囲で確保し直してゼロ化する。
    ReDim deno(1 To NOB + NIB, 1 To NOB + NIB)
    ReDim sms(1 To NOB + NIB, 1 To NOB + NIB)
    
    pinStage = "細分化ノード生成"
    If NEX > 0 Then
        For i = 1 To (NEX + NIN)
             If i > NEX Then
               je = ibin(i - NEX) + 1
             Else
               je = ibex(i) + 1
             End If
             
             For j = 2 To je
                 no1 = mibno(i, j - 1)
                 If je = j Then
                     no2 = mibno(i, 1)
                 Else
                     no2 = mibno(i, j)
                 End If
                                  
                 xc = Sqr((px(no2) - px(no1)) ^ 2 + (py(no2) - py(no1)) ^ 2)
                 
                 lx = deno(no1, no2)
                 ddx = imno(no1, no2)


                 If lx = 0 Then
                     If Int(xc / ddx) >= 2 Then
                         dx = Int(xc / ddx) - 1
                         sms(no1, no2) = dx
                         sms(no2, no1) = dx
                         If deno(no1, no2) = 0 Or deno(no2, no1) = 0 Then
                            For k = 1 To dx
                               newno = newno + 1
                               If no1 < no2 Then
                                   If k = 1 Then
                                      deno(no1, no2) = newno
                                      deno(no2, no1) = newno + dx - 1
                                   End If
                                   px(newno) = px(no1) + (px(no2) - px(no1)) * k / (dx + 1)
                                   py(newno) = py(no1) + (py(no2) - py(no1)) * k / (dx + 1)
                               Else
                                   If k = 1 Then
                                      deno(no2, no1) = newno
                                      deno(no1, no2) = newno + (dx - 1)
                                   End If
                                   px(newno) = px(no2) + (px(no1) - px(no2)) * k / (dx + 1)
                                   py(newno) = py(no2) + (py(no1) - py(no2)) * k / (dx + 1)
                               End If
                            Next k
                         End If
                     End If
                 End If
             Next j
             
        Next i
    End If
    rmax = newno
     
    If NEX > 0 Then
        For i = 1 To (NEX + NIN)
            If i > NEX Then
               je = ibin(i - NEX) + 1
               sa = 21 - NEX
            Else
               je = ibex(i) + 1
               sa = 0
            End If
            
            
            
            r = 0
            For j = 2 To je
                 r = r + 1
                 ibno(i, r) = mibno(i, j - 1)
                 no1 = mibno(i, j - 1)
                 If je = j Then
                     no2 = mibno(i, 1)
                 Else
                     no2 = mibno(i, j)
                 End If
                 If sms(no1, no2) >= 1 Then
                     If no1 < no2 Then
                         For k = 1 To sms(no1, no2)
                              r = r + 1
                              If k = 1 Then
                                 newno = deno(no1, no2)
                              Else
                                 newno = newno + 1
                              End If
                              ibno(i, r) = newno
                         Next k
                     Else
                         For k = 1 To sms(no1, no2)
                              r = r + 1
                              If k = 1 Then
                                 newno = deno(no1, no2)
                              Else
                                 newno = newno - 1
                              End If
                              ibno(i, r) = newno
                         Next k
                     End If
                 End If
            Next j
            
            If i > NEX Then
               ibin(i - NEX) = r
            Else
               ibex(i) = r
            
            End If
            
            
            ibno(i, r + 1) = ibno(i, 1)
            
            
        Next i
    End If
    
    pinStage = "境界再番号化"
    If NEX > 0 Then
        r = 0
        For i = 1 To (NEX + NIN)
            If i = 1 Then
               
               For j = 1 To ibex(i)
                  r = r + 1
                  xlok(r) = ibno(i, j)
                  ibno(i, j) = r
               Next j
            Else
            
               If i > NEX Then
                  je = ibin(i - NEX)
               Else
                  je = ibex(i)
               End If
            
               For j = 1 To je
                  flg = 0
                  For xc = 1 To r
                     If xlok(xc) = ibno(i, j) Then
                        ibno(i, j) = xc
                        flg = 1
                     End If
                     If flg = 1 Then Exit For
                  Next xc
                  If flg = 0 Then
                     r = r + 1
                     xlok(r) = ibno(i, j)
                     ibno(i, j) = r
                  End If
               Next j
            End If
            If i > NEX Then
               je = ibin(i - NEX) + 1
            Else
               je = ibex(i) + 1
            End If
            ibno(i, je) = ibno(i, 1)
        
            If i = NEX Then NOB = r
            If i = (NEX + NIN) Then NIB = r - NOB
        
        Next i

    End If
    Exit Sub
PinputErrorHandler:
    pinErrorNumber = Err.Number
    pinErrorDescription = Err.Description
    On Error GoTo 0
    Err.Raise pinErrorNumber, "Mesh.PINPUT", pinStage & "：" & pinErrorDescription
End Sub

Sub TROUTPUT(NEX As Long, NIN As Long, ibex() As Long, ibin() As Long, ibno() As Long, NOB As Long, NIB As Long, _
            px() As Double, py() As Double, delx() As Double, xlok() As Long)

    Dim i As Long, totalNodes As Long
    Dim nodeX() As Double, nodeY() As Double
    totalNodes = NOB + NIB
    If totalNodes > 0 Then
        ReDim nodeX(1 To totalNodes): ReDim nodeY(1 To totalNodes)
        For i = 1 To totalNodes
            nodeX(i) = px(xlok(i)): nodeY(i) = py(xlok(i))
        Next i
        For i = 1 To totalNodes
            px(i) = nodeX(i): py(i) = nodeY(i)
        Next i
    End If
End Sub
Sub TRMODEL(NEX As Long, NIN As Long, ibex() As Long, ibin() As Long, ibno() As Long, NOB As Long, NIB As Long, _
            mindex() As Long, node As Long, px() As Double, py() As Double, NELM As Long, mtj() As Long, jac() As Long, _
            idm() As Long, ifix() As Long, delx() As Double, xlok() As Long)
    Dim i As Long

    ntp = NOB + NIB
    meshXmin = px(1)
    meshXmax = meshXmin
    meshYmin = py(1)
    meshYmax = meshYmin
    
    For i = 2 To ntp
        If px(i) < meshXmin Then meshXmin = px(i)
        If meshXmax < px(i) Then meshXmax = px(i)
        If py(i) < meshYmin Then meshYmin = py(i)
        If meshYmax < py(i) Then meshYmax = py(i)
    Next i
    
    rax = meshXmax - meshXmin
    ray = meshYmax - meshYmin
    dMax = rax
    If dMax < ray Then dMax = ray
    
    For i = 1 To ntp
        px(i) = (px(i) - meshXmin) / dMax
        py(i) = (py(i) - meshYmin) / dMax
    Next i
    
    NELM = 1
    ia = 2000 + 1
    ib = 2000 + 2
    ic = 2000 + 3
    mtj(1, 1) = ia
    mtj(1, 2) = ib
    mtj(1, 3) = ic
    jac(1, 1) = 0
    jac(1, 2) = 0
    jac(1, 3) = 0

    px(ia) = -1.23
    py(ia) = -0.5
    px(ib) = 2.23
    py(ib) = -0.5
    px(ic) = 0.5
    py(ic) = 2.5
    Call ROUGH(NEX, NIN, ibex, ibin, ibno, ntp, NIB, node, px, py, jnb, nei, NELM, mtj, jac, idm, list, iadres, istack, kv, ibr, map)
    For i = 1 To node
        ifix(i) = 1
    Next i
    Call TRFINE(NEX, dMax, node, px, py, jnb, nei, NELM, mtj, jac, idm, iadres, istack, delx)
    If (NELM <> 2 * node + 1) Then
        Err.Raise vbObjectError + 3490, "Mesh.TRMODEL", "境界メッシュ処理に失敗しました（TRMODEL）。"
        Exit Sub
    End If
    Call REMOVE(NEX, mindex, NELM, mtj, jac, idm)
    Call CHECK(NELM, mtj, jac)
    
    For i = 1 To node
        px(i) = px(i) * dMax + meshXmin
        py(i) = py(i) * dMax + meshYmin
    Next i

End Sub
Sub ROUGH(NEX As Long, NIN As Long, ibex() As Long, ibin() As Long, ibno() As Long, ntp As Long, NIB As Long, node As Long, _
           px() As Double, py() As Double, jnb() As Long, nei() As Long, NELM As Long, mtj() As Long, _
           jac() As Long, idm() As Long, list() As Long, iadres() As Long, istack() As Long, kv() As Long, _
           ibr() As Long, map() As Long)

    Dim i As Long, j As Long, k As Long, nb As Long, np As Long
    Dim ma As Long, mb As Long, mc As Long, mS As Long, mp As Long
    Dim js As Long, jz As Long, loc As Long, it As Long
    Dim xs As Double, ys As Double, xa As Double, ya As Double
    Dim s As Long
    
    For i = 1 To NEX
        jz = 0
        nb = i
        np = ibex(nb)
        For j = 1 To np
            list(j) = ibno(nb, j)
        Next j
        Call BOUGEN(jz, np, list, ntp, px, py, jnb, nei, NELM, mtj, jac, idm, iadres, istack, kv, ibr, map, node)
        
        For k = 1 To NELM
            If idm(k) = jz Then
                ma = iadres(mtj(k, 1))
                mb = iadres(mtj(k, 2))
                mc = iadres(mtj(k, 3))
                mS = ma * mb * mc
                mp = (mb - ma) * (mc - mb) * (ma - mc)
                If mS <> 0 And mp > 0 Then
                    idm(k) = nb
                End If
            End If
        Next k
    Next i
    
    For i = 1 To NIN
        nb = i
        np = ibin(nb)
        For j = 1 To np
            list(j) = ibno(NEX + i, j)
        Next j
        js = list(1)
        xs = px(js)
        ys = py(js)
        LOCATE xs, ys, px, py, mtj, jac, NELM, loc

        jz = idm(loc)
        BOUGEN jz, np, list, ntp, px, py, jnb, nei, NELM, mtj, jac, idm, iadres, istack, kv, ibr, map, node
        For k = 1 To NELM
            If idm(k) = jz Then
                ma = iadres(mtj(k, 1))
                mb = iadres(mtj(k, 2))
                mc = iadres(mtj(k, 3))
                mS = ma * mb * mc
                mp = (mb - ma) * (mc - mb) * (ma - mc)
                If mS <> 0 And mp < 0 Then
                    idm(k) = 0
                End If
            End If
        Next k
    Next i
    
    P5ZeroLongArray iadres
    
    ' PINPUT renumbers all boundary nodes; BOUGEN has inserted them already.
    ' Insert only any remaining uninserted nodes, never the hole boundary twice.
    Do While node < ntp
        node = node + 1
        xa = px(node): ya = py(node)
        Call LOCATE(xa, ya, px, py, mtj, jac, NELM, it)
        jz = idm(it)
        Call DELAUN(jz, node, node, ntp, px, py, jnb, nei, NELM, mtj, jac, idm, iadres, istack)
    Loop

    If node <> ntp Then
        ' Show ERROR message and stop the execution.
        s = s + 1
        Err.Raise vbObjectError + 3490, "Mesh.ROUGH", "境界メッシュ処理に失敗しました（ROUGH）。"

    End If

End Sub
Sub BOUGEN(jz As Long, np As Long, list() As Long, ntp As Long, px() As Double, py() As Double, _
             jnb() As Long, nei() As Long, NELM As Long, mtj() As Long, jac() As Long, idm() As Long, iadres() As Long, _
             istack() As Long, kv() As Long, ibr() As Long, map() As Long, node As Long)

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim inp As Long
    Dim js As Long
    Dim ip As Long
    Dim iq As Long
    Dim iv As Long

    inp = 0
    P5ZeroLongArray iadres

    For i = 1 To np
        iadres(list(i)) = i
    Next i

    js = list(1)

    If node < js Then
        inp = inp + 1
        Call DELAUN(jz, js, js, ntp, px, py, jnb, nei, NELM, mtj, jac, idm, iadres, istack)
    End If

    For i = 1 To np
        ip = list((i Mod np) + 1)
        iq = list(i)

        If (node < ip) And (i <> np) Then
            inp = inp + 1
            Call DELAUN(jz, ip, ip, ntp, px, py, jnb, nei, NELM, mtj, jac, idm, iadres, istack)
        End If

        For j = 1 To jnb(iq)
            For k = 1 To 3
                If mtj(nei(iq, j), k) = ip Then GoTo loopi_end
            Next k
        Next j

        Call SEARCH(jz, ip, iq, jnb, nei, NELM, mtj, jac, idm, istack, iv, kv, iadres, ibr)
        Call POLY(iq, ip, iv, kv, px, py, NELM, mtj, jac, jnb, nei, map)
loopi_end:
    Next i

    node = node + inp

End Sub
Sub DELAUN(jz As Long, js As Long, jg As Long, ntp As Long, px() As Double, py() As Double, _
           jnb() As Long, nei() As Long, NELM As Long, mtj() As Long, jac() As Long, idm() As Long, _
           iadres() As Long, istack() As Long)

    Dim i As Long, itop As Long, maxstk As Long, ip As Long, it As Long, ia As Long, ib As Long, ic As Long
    Dim iv1 As Long, iv2 As Long, iv3 As Long, idf As Long, mS As Long, iedge As Long, il As Long, ir As Long
    Dim iera As Long, ierb As Long, ierl As Long, iswap As Long
    Dim xp As Double, yp As Double
    
    itop = 0
    maxstk = ntp
    
    For i = js To jg
        ip = i
        xp = px(ip)
        yp = py(ip)
        ' Assuming LOCATE is another subroutine/function in the original code
        Call LOCATE(xp, yp, px(), py(), mtj(), jac, NELM, it)
        
        If idm(it) <> jz Then
            Err.Raise vbObjectError + 3490, "Mesh.DELAUN", "境界メッシュ処理に失敗しました（DELAUN）。"
            Exit Sub
        End If
        
        ia = jac(it, 1)
        ib = jac(it, 2)
        ic = jac(it, 3)
        iv1 = mtj(it, 1)
        iv2 = mtj(it, 2)
        iv3 = mtj(it, 3)
    
        mtj(it, 1) = ip
        mtj(it, 2) = iv1
        mtj(it, 3) = iv2
        jac(it, 1) = NELM + 2
        jac(it, 2) = ia
        jac(it, 3) = NELM + 1
    
        AssertMeshElementCapacity NELM + 1
        NELM = NELM + 1
        idm(NELM) = jz
        mtj(NELM, 1) = ip
        mtj(NELM, 2) = iv2
        mtj(NELM, 3) = iv3
        jac(NELM, 1) = it
        jac(NELM, 2) = ib
        jac(NELM, 3) = NELM + 1
    
        AssertMeshElementCapacity NELM + 1
        NELM = NELM + 1
        idm(NELM) = jz
        mtj(NELM, 1) = ip
        mtj(NELM, 2) = iv3
        mtj(NELM, 3) = iv1
        jac(NELM, 1) = NELM - 1
        jac(NELM, 2) = ic
        jac(NELM, 3) = it
    
        Call INCR(iv1, NELM, jnb, nei)
        Call INCR(iv2, NELM - 1, jnb, nei)
    
        If iv3 <= 2000 Then
            nei(iv3, NEIBOR(iv3, it, jnb(), nei())) = NELM - 1
        End If
    
        Call INCR(iv3, NELM, jnb, nei)
    
        jnb(ip) = 3
        nei(ip, 1) = it
        nei(ip, 2) = NELM - 1
        nei(ip, 3) = NELM
    
        If ia <> 0 Then
            mS = iadres(iv1) * iadres(iv2)
            idf = Abs(iadres(iv1) - iadres(iv2))
            If (idm(ia) = jz) And ((mS = 0) Or (idf <> 1)) Then
                itop = itop + 1
                istack(itop) = IPUSH(it, maxstk, itop)
            End If
        End If
    
        If ib <> 0 Then
            Call edge(ib, it, jac(), iedge)
            jac(ib, iedge) = NELM - 1
            mS = iadres(iv2) * iadres(iv3)
            idf = Abs(iadres(iv2) - iadres(iv3))
            If (idm(ib) = jz) And ((mS = 0) Or (idf <> 1)) Then
                itop = itop + 1
                istack(itop) = IPUSH(NELM - 1, maxstk, itop)
            End If
        End If
    
        If ic <> 0 Then
            Call edge(ic, it, jac(), iedge)
            jac(ic, iedge) = NELM
            mS = iadres(iv3) * iadres(iv1)
            idf = Abs(iadres(iv3) - iadres(iv1))
            If (idm(ic) = jz) And ((mS = 0) Or (idf <> 1)) Then
                itop = itop + 1
                istack(itop) = IPUSH(NELM, maxstk, itop)
            End If
        End If

        Do
            If itop > 0 Then
                il = istack(itop)
                itop = itop - 1
                ir = jac(il, 2)
                Call edge(ir, il, jac(), ierl)
                iera = (ierl Mod 3) + 1
                ierb = (iera Mod 3) + 1
                iv1 = mtj(ir, ierl)
                iv2 = mtj(ir, iera)
                iv3 = mtj(ir, ierb)
                Call swap(px(iv1), py(iv1), px(iv2), py(iv2), px(iv3), py(iv3), xp, yp, iswap)
                If iswap = 1 Then
                    ia = jac(ir, iera)
                    ib = jac(ir, ierb)
                    ic = jac(il, 3)
                    mtj(il, 3) = iv3
                    jac(il, 2) = ia
                    jac(il, 3) = ir
                    mtj(ir, 1) = ip
                    mtj(ir, 2) = iv3
                    mtj(ir, 3) = iv1
                    jac(ir, 1) = il
                    jac(ir, 2) = ib
                    jac(ir, 3) = ic
                    Call DECR(iv1, il, jnb, nei)
                    Call DECR(iv2, ir, jnb, nei)
                    Call INCR(ip, ir, jnb, nei)
                    Call INCR(iv3, il, jnb, nei)
                    If ia <> 0 Then
                        Call edge(ia, ir, jac(), iedge)
                        jac(ia, iedge) = il
                        mS = iadres(iv2) * iadres(iv3)
                        idf = Abs(iadres(iv2) - iadres(iv3))
                        If (idm(ia) = jz) And ((mS = 0) Or (idf <> 1)) Then
                            itop = itop + 1
                            istack(itop) = IPUSH(il, maxstk, itop)
                        End If
                    End If
                    If ib <> 0 Then
                        mS = iadres(iv3) * iadres(iv1)
                        idf = Abs(iadres(iv3) - iadres(iv1))
                        If (idm(ib) = jz) And ((mS = 0) Or (idf <> 1)) Then
                            itop = itop + 1
                            istack(itop) = IPUSH(ir, maxstk, itop)
                        End If
                    End If
                    If ic <> 0 Then
                        Call edge(ic, il, jac(), iedge)
                        jac(ic, iedge) = ir
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
    Next i
End Sub
Sub SEARCH(iz As Long, ip As Long, iq As Long, jnb() As Long, nei() As Long, _
           NELM As Long, mtj() As Long, jac() As Long, idm() As Long, istack() As Long, _
            iv As Long, kv() As Long, iadres() As Long, ibr() As Long)

    Dim i As Long, j As Long, ia As Long, ib As Long, ja As Long
    Dim jb As Long, mstk As Long, nbr As Long, idf As Long
    Dim mS As Long, ielm As Long, jelm As Long, kelm As Long
    Dim mmin As Long, jr As Long

    iv = 0
    mstk = 0
    nbr = 0


    P5ZeroLongArray kv
    P5ZeroLongArray ibr
    P5ZeroLongArray istack
    
    For i = 1 To jnb(iq)
        nbr = nbr + 1
        ibr(nei(iq, i)) = nbr
    Next i

    For i = 1 To jnb(iq)
        ielm = nei(iq, i)
        j = IVERT(ielm, iq, mtj())
        ja = j Mod 3 + 1
        jb = ja Mod 3 + 1
        ia = mtj(ielm, ja)
        ib = mtj(ielm, jb)
        idf = Abs(iadres(ia) - iadres(ib))
        mS = iadres(ia) * iadres(ib)
        If idf = 1 And mS <> 0 Then GoTo Continue1
        jelm = jac(ielm, ja)
        If jelm = 0 Then GoTo Continue1
        If idm(jelm) <> iz Then GoTo Continue1
        nbr = nbr + 1
        ibr(jelm) = nbr
        If mtj(jelm, 1) = ip Then GoTo label80
        If mtj(jelm, 2) = ip Then GoTo label80
        If mtj(jelm, 3) = ip Then GoTo label80
        mstk = mstk + 1
        istack(mstk) = jelm
Continue1:
    Next i

    Do
        If mstk = 0 Then
            'Show ERROR message and stop the execution
            Err.Raise vbObjectError + 3490, "Mesh.SEARCH", "境界メッシュ処理に失敗しました（SEARCH）。"

        End If
        ielm = istack(1)
        mstk = mstk - 1
        For i = 1 To mstk
            istack(i) = istack(i + 1)
        Next i
        istack(mstk + 1) = 0
        For j = 1 To 3
            ja = j Mod 3 + 1
            ia = mtj(ielm, j)
            ib = mtj(ielm, ja)
            idf = Abs(iadres(ia) - iadres(ib))
            mS = iadres(ia) * iadres(ib)
            If idf = 1 And mS <> 0 Then GoTo Continue2
            jelm = jac(ielm, j)
            If jelm = 0 Then GoTo Continue2
            If ibr(jelm) <> 0 Then GoTo Continue2
            If idm(jelm) <> iz Then GoTo Continue2
            nbr = nbr + 1
            ibr(jelm) = nbr
            If mtj(jelm, 1) = ip Then GoTo label80
            If mtj(jelm, 2) = ip Then GoTo label80
            If mtj(jelm, 3) = ip Then GoTo label80
            mstk = mstk + 1
            istack(mstk) = jelm
Continue2:
        Next j
    Loop

label80:
    Do
        iv = iv + 1
        kv(iv) = jelm
        If ibr(jelm) <= jnb(iq) Then Exit Do
        mmin = 4001 + 1
        For j = 1 To 3
            jr = jac(jelm, j)
            If jr = 0 Then GoTo Continue3
            If ibr(jr) = 0 Then GoTo Continue3
            ja = j Mod 3 + 1
            ia = mtj(jelm, j)
            ib = mtj(jelm, ja)
            idf = Abs(iadres(ia) - iadres(ib))
            mS = iadres(ia) * iadres(ib)
            If idf = 1 And mS <> 0 Then GoTo Continue3
            If ibr(jr) < mmin Then
                kelm = jr
                mmin = ibr(jr)
            End If
Continue3:
        Next j
        If mmin = 4001 + 1 Then
            'Show ERROR message and stop the execution
            Err.Raise vbObjectError + 3490, "Mesh.SEARCH", "境界メッシュ処理に失敗しました（SEARCH）。"

        End If
        jelm = kelm
    Loop

End Sub
Sub POLY(iq As Long, ip As Long, iv As Long, kv() As Long, _
          px() As Double, py() As Double, NELM As Long, mtj() As Long, _
          jac() As Long, jnb() As Long, nei() As Long, map() As Long)

    Dim i As Long, j As Long, k As Long, l As Long
    Dim ia As Long, ja As Long, ir As Long, jr As Long, iv1 As Long, iv2 As Long
    Dim iedge As Long, npa As Long, npb As Long, nta As Long, ntb As Long
    Dim ielm As Long, jelm As Long, kelm As Long, ivx As Long
    Dim ips As Long, ipg As Long, iva As Long, ivb As Long
    
    Dim iena(1 To 100, 1 To 3) As Long
    Dim ienb(1 To 100, 1 To 3) As Long
    Dim jeea(1 To 100, 1 To 3) As Long
    Dim jeeb(1 To 100, 1 To 3) As Long
    Dim nsra(1 To 100 + 2) As Long
    Dim nsrb(1 To 100 + 2) As Long
    Dim ihen(1 To 2 * 100 + 1, 1 To 2) As Long
    Dim jhen(1 To 2 * 100 + 1) As Long
    Dim iad(1 To 2 * 100 + 1) As Long
    Dim jstack(1 To 100 + 2) As Long

    If iv = 2 Then
        Call edge(kv(1), kv(2), jac, ia)
        Call edge(kv(2), kv(1), jac, ja)

        ir = jac(kv(1), (ia Mod 3) + 1)
        jr = jac(kv(2), (ja Mod 3) + 1)

        mtj(kv(1), (ia Mod 3) + 1) = iq
        jac(kv(1), ia) = jr
        jac(kv(1), (ja Mod 3) + 1) = kv(2)

        mtj(kv(2), (ja Mod 3) + 1) = ip
        jac(kv(2), ja) = ir
        jac(kv(2), (ja Mod 3) + 1) = kv(1)

        If ir <> 0 Then
            Call edge(ir, kv(1), jac, iedge)
            jac(ir, iedge) = kv(2)
        End If
        If jr <> 0 Then
            Call edge(jr, kv(2), jac, iedge)
            jac(jr, iedge) = kv(1)
        End If

        iv1 = mtj(kv(1), ia)
        iv2 = mtj(kv(2), ja)
        Call DECR(iv1, kv(2), jnb, nei)
        Call DECR(iv2, kv(1), jnb, nei)
        Call INCR(iq, kv(1), jnb, nei)
        Call INCR(ip, kv(2), jnb, nei)
    Else
        npa = 0
        npb = 0
        ' Loop to set nsra and nsrb to 0
        For i = 1 To 100 + 2
            nsra(i) = 0
            nsrb(i) = 0
        Next i
        ' Loop to set map to 0
        For i = 1 To NELM
            map(i) = 0
        Next i
        ' Loop for kv values
        For i = 1 To iv
            map(kv(i)) = 1
        Next i
        For i = 1 To iv
            ielm = kv(i)
            For j = 1 To 3
                ivx = mtj(ielm, j)
                Call DECR(ivx, ielm, jnb, nei)
            Next j
        Next i

        Call PICK(iq, ip, iv, kv, mtj, jac, map, npa, npb, nsra, nsrb)
        Call subdiv(npa, nsra, px, py, nta, iena, jeea, ihen, jhen, iad, jstack)
        Call subdiv(npb, nsrb, px, py, ntb, ienb, jeeb, ihen, jhen, iad, jstack)

        If iv <> nta + ntb Then
            Err.Raise vbObjectError + 3490, "Mesh.POLY", "境界メッシュ処理に失敗しました（POLY）。"

        End If
        
        For i = 1 To iv
            For j = 1 To 3
                jac(kv(i), j) = 0
            Next j
        Next i
    
        For i = 1 To nta
            ielm = kv(i)
            For j = 1 To 3
                mtj(ielm, j) = iena(i, j)
                If jeea(i, j) <> 0 Then
                    jac(ielm, j) = kv(jeea(i, j))
                End If
            Next j
        Next i
    
        For i = 1 To ntb
            ielm = kv(nta + i)
            For j = 1 To 3
                mtj(ielm, j) = ienb(i, j)
                If jeeb(i, j) <> 0 Then
                    jac(ielm, j) = kv(nta + jeeb(i, j))
                End If
            Next j
        Next i
    
        For i = 1 To iv
            ielm = kv(i)
            For j = 1 To 3
                ivx = mtj(ielm, j)
                Call INCR(ivx, ielm, jnb, nei)
            Next j
        Next i

        For i = 1 To iv
            ielm = kv(i)
            For j = 1 To 3
            jelm = jac(ielm, j)
            If jelm <> 0 Then GoTo ContinueLoop
                ips = mtj(ielm, j)
                ipg = mtj(ielm, (j Mod 3) + 1)
                For k = 1 To jnb(ipg)
                    kelm = nei(ipg, k)
                    For l = 1 To 3
                        iva = mtj(kelm, l)
                        ivb = mtj(kelm, (l Mod 3) + 1)
                        If iva = ipg And ivb = ips Then
                            jac(ielm, j) = kelm
                            jac(kelm, l) = ielm
                            GoTo ContinueLoop
                        End If
                    Next l
                Next k
ContinueLoop:
            Next j
        Next i
    End If

End Sub
Sub PICK(iq As Long, ip As Long, iv As Long, kv() As Long, mtj() As Long, _
         jac() As Long, map() As Long, npa As Long, npb As Long, _
          nsra() As Long, nsrb() As Long)

    Dim i As Long, ivx As Long, jvx As Long, jelm As Long

    ' Initialize npa and first two values of nsra
    npa = 1
    nsra(npa) = ip
    ivx = IVERT(kv(1), ip, mtj())
    npa = npa + 1
    nsra(npa) = mtj(kv(1), (ivx Mod 3) + 1)

    ' Loop through the elements of kv
    For i = 2 To iv - 1
        jvx = IVERT(kv(i), nsra(npa), mtj())
        jelm = jac(kv(i), jvx)
        If jelm = 0 Then
            npa = npa + 1
            nsra(npa) = mtj(kv(i), (jvx Mod 3) + 1)
        ElseIf map(jelm) = 0 Then
            npa = npa + 1
            nsra(npa) = mtj(kv(i), (jvx Mod 3) + 1)
        End If
    Next i

    ' Set final value of nsra
    npa = npa + 1
    nsra(npa) = iq

    ' Initialize npb and first two values of nsrb
    npb = 1
    nsrb(npb) = iq
    ivx = IVERT(kv(iv), iq, mtj())
    npb = npb + 1
    nsrb(npb) = mtj(kv(iv), (ivx Mod 3) + 1)

    ' Loop backward through the elements of kv
    For i = iv - 1 To 2 Step -1
        jvx = IVERT(kv(i), nsrb(npb), mtj())
        jelm = jac(kv(i), jvx)
        If jelm = 0 Or map(jelm) = 0 Then
            npb = npb + 1
            nsrb(npb) = mtj(kv(i), (jvx Mod 3) + 1)
        End If
    Next i

    ' Set final value of nsrb
    npb = npb + 1
    nsrb(npb) = ip

    ' Check for error
    If iv <> npa + npb - 4 Then
        Err.Raise vbObjectError + 3490, "Mesh.PICK", "境界メッシュ処理に失敗しました（PICK）。"

    End If

End Sub
Sub subdiv(npl As Long, nsr() As Long, px() As Double, py() As Double, _
           nte As Long, ien() As Long, jee() As Long, _
           ihen() As Long, jhen() As Long, iad() As Long, _
           jstack() As Long)
    Dim i As Long, j As Long, k As Long, nnpl As Long
    Dim nbs As Long, ia As Long, ib As Long, ic As Long, ix As Long
    Dim xa As Double, ya As Double, xb As Double, yb As Double, see As Double

    nte = 0
    nnpl = npl

    ' Zero out the arrays
    For i = 1 To 100
        For j = 1 To 3
            ien(i, j) = 0
            jee(i, j) = 0
        Next j
    Next i

    Do
        If nnpl >= 3 Then
            nbs = 1
            Do
                If nbs <= nnpl - 1 Then
                    ia = nsr(nbs)
                    ib = nsr(nbs + 1)
                    ic = nsr((nbs + 1) Mod nnpl + 1)
                    xa = px(ib) - px(ia)
                    ya = py(ib) - py(ia)
                    xb = px(ic) - px(ia)
                    yb = py(ic) - py(ia)
                    see = xa * yb - xb * ya
                    If see > 0.0000000001 Then
                        nte = nte + 1
                        ien(nte, 1) = ia
                        ien(nte, 2) = ib
                        ien(nte, 3) = ic
                        nnpl = nnpl - 1
                        For i = nbs + 1 To nnpl
                            nsr(i) = nsr(i + 1)
                        Next i
                    End If
                    nbs = nbs + 1
                Else
                    Exit Do
                End If
            Loop
        Else
            Exit Do
        End If
    Loop

    ix = 0
    For i = 1 To 2 * 100 + 1
        ihen(i, 1) = 0
        ihen(i, 2) = 0
        jhen(i) = 0
        iad(i) = 0
    Next i

    For i = 1 To nte
loopj:
        For j = 1 To 3
            ia = ien(i, j)
            ib = ien(i, (j Mod 3) + 1)
            For k = 1 To ix
                If ihen(k, 1) = ib And ihen(k, 2) = ia Then
                    jee(i, j) = jhen(k)
                    jee(jhen(k), iad(k)) = i
                    GoTo NextSubdivisionEdge
                End If
            Next k
            ix = ix + 1
            jhen(ix) = i
            iad(ix) = j
            ihen(ix, 1) = ia
            ihen(ix, 2) = ib
NextSubdivisionEdge:
        Next j
    Next i

    ' Make sure there's a corresponding subroutine or function for LAWSON in VBA
    LAWSON nte, ien(), jee(), npl, px(), py(), jstack()
    

End Sub
Sub LAWSON(nte As Long, ien() As Long, jee() As Long, npl, px() As Double, py() As Double, jstack() As Long)

    Dim i As Long, j As Long, itop As Long, maxstk As Long, ncount As Long
    Dim ielm As Long, il As Long, ir As Long, jl1 As Long, jl2 As Long
    Dim jl3 As Long, jr1 As Long, jr2 As Long, jr3 As Long
    Dim iv1 As Long, iv2 As Long, iv3 As Long, iv4 As Long
    Dim ia As Long, ib As Long, iswap As Long, iedge As Long
    Dim xx As Double, yy As Double

    itop = 0
    maxstk = npl
    ncount = 0

    For i = 1 To nte
        ielm = i
        itop = itop + 1
        jstack(itop) = IPUSH(ielm, maxstk, itop) ' Assuming IPUSH is another subroutine or function you've converted
    Next i

    Do
        If itop > 0 Then
            ncount = ncount + 1
            If ncount > 100 Then
                Err.Raise vbObjectError + 3490, "Mesh.LAWSON", "境界メッシュ処理に失敗しました（LAWSON）。"

            End If
            il = jstack(itop)
            itop = itop - 1

            For j = 1 To 3
                jl1 = j
                jl2 = (jl1 Mod 3) + 1
                jl3 = (jl2 Mod 3) + 1
                ir = jee(il, jl1)
                If ir = 0 Then
                    GoTo SkipCycle ' Equivalent to cycle in Fortran
                End If
                iv1 = ien(il, jl1)
                iv2 = ien(il, jl2)
                iv3 = ien(il, jl3)
                xx = px(ien(il, jl3))
                yy = py(ien(il, jl3))

                ' Assuming EDGE is another subroutine or function you've converted
                edge ir, il, jee(), jr1
                jr2 = (jr1 Mod 3) + 1
                jr3 = (jr2 Mod 3) + 1
                iv4 = ien(ir, jr3)

                ' Assuming SWAP is another subroutine or function you've converted
                swap px(iv2), py(iv2), px(iv1), py(iv1), px(iv4), py(iv4), xx, yy, iswap

                If iswap = 1 Then
                    ia = jee(il, jl2)
                    ib = jee(ir, jr2)
                    ien(il, jl2) = iv4
                    jee(il, jl1) = ib
                    jee(il, jl2) = ir
                    ien(ir, jr2) = iv3
                    jee(ir, jr1) = ia
                    jee(ir, jr2) = il

                    If ia <> 0 Then
                        edge ia, il, jee(), iedge
                        jee(ia, iedge) = ir
                    End If

                    If ib <> 0 Then
                        edge ib, ir, jee(), iedge
                        jee(ib, iedge) = il
                    End If

                    itop = itop + 1
                    jstack(itop) = IPUSH(il, maxstk, itop)
                    Exit For
                End If

SkipCycle:
            Next j
        Else
            Exit Do
        End If
    Loop
End Sub
Sub TRFINE(NEX As Long, dMax As Double, node As Long, _
            px() As Double, py() As Double, jnb() As Long, _
            nei() As Long, NELM As Long, mtj() As Long, _
            jac() As Long, idm() As Long, iadres() As Long, _
            istack() As Long, delx() As Double)


    Dim i As Long, j As Long, k As Long, jz As Long, js As Long, ndiv As Long
    Dim alpha As Double, range As Double
    Dim px1 As Double, py1 As Double, px2 As Double, py2 As Double
    Dim px3 As Double, py3 As Double, px4 As Double, py4 As Double
    Dim dltd As Double, delta As Double, xg As Double, yg As Double
    Dim xs As Double, ys As Double, xp As Double, yp As Double

    alpha = 3#
    range = dMax
    px1 = 0#
    py1 = 0#
    px2 = 1#
    py2 = 0#
    px3 = 1#
    py3 = 1#
    px4 = 0#
    py4 = 1#
    
    For i = 1 To NEX
        jz = i
        js = node + 1
        If dMax < delx(i) Then
            GoTo SkipCycle
        End If

        ndiv = Int(range / delx(i))
        delta = 1# / ndiv
        dltd = delta / alpha
        dltd = dltd * dltd

        For k = 1 To ndiv + 1
            xg = px2 + (px3 - px2) / ndiv * (k - 1)
            xs = px1 + (px4 - px1) / ndiv * (k - 1)
            yg = py2 + (py3 - py2) / ndiv * (k - 1)
            ys = py1 + (py4 - py1) / ndiv * (k - 1)

            For j = 1 To ndiv + 1
                xp = xs + (xg - xs) / ndiv * (j - 1)
                yp = ys + (yg - ys) / ndiv * (j - 1)

                ' Assuming TRPLACE is another function you've converted to VBA
                If TRPLACE(jz, xp, yp, NELM, mtj(), jac(), idm(), px(), py(), dltd) Then
                    AssertMeshNodeCapacity node + 1
                    node = node + 1
                    px(node) = xp
                    py(node) = yp
                End If
            Next j
        Next k

        ' Assuming DELAUN is another subroutine you've converted to VBA
        DELAUN jz, js, node, node, px, py, jnb, nei, NELM, mtj, jac, idm, iadres, istack

SkipCycle:
    Next i

End Sub
Function TRPLACE(iz As Long, xp As Double, yp As Double, NELM As Long, mtj() As Long, jac() As Long, idm() As Long, px() As Double, py() As Double, dltd As Double)

    Dim j As Long, loc As Long, ia As Long, ib As Long, jr As Long
    Dim xa As Double, ya As Double, xb As Double, yb As Double
    Dim dpp As Double, eps As Double, eps0 As Double, epl As Double
    
    Call LOCATE(xp, yp, px(), py(), mtj(), jac(), NELM, loc)

    
    If iz = idm(loc) Then
        For j = 1 To 3
            ia = mtj(loc, j)
            xa = px(ia)
            ya = py(ia)
            dpp = (xa - xp) * (xa - xp) + (ya - yp) * (ya - yp)
            
            If dpp < dltd Then
                TRPLACE = False
                Exit Function
            End If
            
            jr = jac(loc, j)
            If iz <> idm(jr) Then
                ib = mtj(loc, (j Mod 3) + 1)
                xb = px(ib)
                yb = py(ib)
                
                eps = (yb - ya) * xp - (xb - xa) * yp - xa * (yb - ya) + ya * (xb - xa)
                eps = eps * eps
                eps0 = (yb - ya) * (yb - ya) + (xb - xa) * (xb - xa)
                epl = eps / eps0
                
                If epl < dltd Then
                    TRPLACE = False
                    Exit Function
                End If
            End If
        Next j
        TRPLACE = True
    Else
        TRPLACE = False
    End If

End Function
Sub REMOVE(NEX As Long, mindex() As Long, NELM As Long, mtj() As Long, jac() As Long, idm() As Long)

    Dim i As Long, j As Long, k As Long, l As Long
    Dim iz As Long, inelm As Long, ielm As Long, jelm As Long
    Dim kelm As Long, lelm As Long, ikp As Long, kedg As Long, ledg As Long
    Dim mkp(1 To 3) As Long
    Dim jkp(1 To 3) As Long

    ielm = 0
    mindex(1) = 1
    For i = 1 To NEX
        iz = i
        inelm = 0
        For j = mindex(i) To NELM
            If idm(j) = iz Then
                ielm = ielm + 1
                jelm = j
                inelm = inelm + 1
                If ielm <> jelm Then
                    For k = 1 To 3
                        mkp(k) = mtj(ielm, k)
                        jkp(k) = jac(ielm, k)
                    Next k
                    ikp = idm(ielm)
                    For k = 1 To 3
                        kelm = jac(jelm, k)
                        If kelm <> 0 Then
                            Call edge(kelm, jelm, jac(), kedg) ' Assuming EDGE is another function you've converted
                            jac(kelm, kedg) = ielm + 4001 + 1
                        End If
                    Next k
                    For l = 1 To 3
                        lelm = jkp(l)
                        If lelm <> 0 Then
                            Call edge(lelm, ielm, jac(), ledg) ' Assuming EDGE is another function you've converted
                            jac(lelm, ledg) = jelm + 4001 + 1
                        End If
                    Next l
                    For k = 1 To 3
                        jac(ielm, k) = jac(ielm, k) Mod (4001 + 1)
                        jac(jelm, k) = jac(jelm, k) Mod (4001 + 1)
                        kelm = jac(ielm, k)
                        lelm = jac(jelm, k)
                        For l = 1 To 3
                            If kelm > 0 Then jac(kelm, l) = jac(kelm, l) Mod (4001 + 1)
                            If lelm > 0 Then jac(lelm, l) = jac(lelm, l) Mod (4001 + 1)
                        Next l
                    Next k
                    For k = 1 To 3
                        jkp(k) = jac(ielm, k)
                    Next k
                    For k = 1 To 3
                        mtj(ielm, k) = mtj(jelm, k)
                        jac(ielm, k) = jac(jelm, k)
                        mtj(jelm, k) = mkp(k)
                        jac(jelm, k) = jkp(k)
                    Next k
                    idm(ielm) = idm(jelm)
                    idm(jelm) = ikp
                End If
            End If
        Next j
        mindex(i + 1) = mindex(i) + inelm
    Next i

    For i = 1 To ielm
        For j = 1 To 3
            If ielm < jac(i, j) Then jac(i, j) = 0
        Next j
    Next i
    For i = ielm + 1 To NELM
        For j = 1 To 3
            mtj(i, j) = 0
            jac(i, j) = 0
        Next j
    Next i

    NELM = ielm

End Sub
Sub CHECK(NELM, mtj, jac)

    Dim i As Long, j As Long, k As Long
    Dim ielm As Long, jelm As Long, kelm As Long
    Dim ia As Long, ib As Long, ja As Long, jb As Long

    For i = 1 To NELM
        For j = 1 To 3
            ielm = i
            ia = mtj(i, j)
            ib = mtj(i, (j Mod 3) + 1)
            jelm = jac(i, j)
            If jelm = 0 Then GoTo loopj
            For k = 1 To 3
                kelm = jac(jelm, k)
                If kelm = ielm Then
                    ja = mtj(jelm, k)
                    jb = mtj(jelm, (k Mod 3) + 1)
                    If (ia = jb) And (ib = ja) Then GoTo loopj
                    'Show ERROR message and stop the execution
                    Err.Raise vbObjectError + 3490, "Mesh.CHECK", "境界メッシュ処理に失敗しました（CHECK）。"

                End If
            Next k
            'Show ERROR message and stop the execution
            Err.Raise vbObjectError + 3490, "Mesh.CHECK", "境界メッシュ処理に失敗しました（CHECK）。"

loopj:
        Next j
    Next i

End Sub
Sub INCR(n As Long, l As Long, jnb() As Long, nei() As Long)

    If n <= 2000 Then
        jnb(n) = jnb(n) + 1
        If 100 < jnb(n) Then
            Err.Raise vbObjectError + 3490, "Mesh.INCR", "境界メッシュ処理に失敗しました（INCR）。"

        End If
        nei(n, jnb(n)) = l
    End If

End Sub

Sub DECR(n As Long, l As Long, jnb() As Long, nei() As Long)

    Dim i As Long, inb As Long

    If n <= 2000 Then
        inb = NEIBOR(n, l, jnb(), nei())
        For i = inb To jnb(n) - 1
            nei(n, i) = nei(n, i + 1)
        Next i
        nei(n, jnb(n)) = 0
        jnb(n) = jnb(n) - 1
        If jnb(n) = 0 Then
            Err.Raise vbObjectError + 3490, "Mesh.DECR", "境界メッシュ処理に失敗しました（DECR）。"

        End If
    End If

End Sub

Sub LOCATE(xp As Double, yp As Double, px() As Double, py() As Double, mtj() As Long, jac() As Long, NELM As Long, itri As Long)
    Dim i As Long, ia As Long, ib As Long, neighborElement As Long
    Dim currentElement As Long, iterationNo As Long, maxIterations As Long
    Dim edgeCross As Double, edgeTolerance As Double, edgeLength As Double
    Dim visited() As Boolean, moved As Boolean

    P5LocateCallCount = P5LocateCallCount + 1
    If NELM < 1 Then
        P5LocateLastFailure = "要素数が1未満です。"
        Err.Raise vbObjectError + 3475, "Mesh.LOCATE", P5LocateLastFailure
    End If
    If P5LocateLastElement >= 1 And P5LocateLastElement <= NELM Then
        currentElement = P5LocateLastElement
    Else
        currentElement = NELM
    End If
    maxIterations = NELM + 1
    If maxIterations < 2 Then maxIterations = 2
    If locateCapacity < NELM Then
        locateCapacity = NELM + 256: ReDim locateStamp(1 To locateCapacity): locateEpoch = 0
    End If
    If locateEpoch = 2147483647 Then ReDim locateStamp(1 To locateCapacity): locateEpoch = 0
    locateEpoch = locateEpoch + 1
    For iterationNo = 1 To maxIterations
        If currentElement < 1 Or currentElement > NELM Then
            P5LocateLastFailure = "探索中に無効な要素添字へ遷移しました。要素=" & CStr(currentElement)
            P5LocateLastIterations = iterationNo
            Err.Raise vbObjectError + 3476, "Mesh.LOCATE", P5LocateLastFailure
        End If
        If locateStamp(currentElement) = locateEpoch Then
            P5LocateLastFailure = "要素近傍探索が循環しました。開始要素=" & CStr(currentElement)
            P5LocateLastIterations = iterationNo
            Err.Raise vbObjectError + 3477, "Mesh.LOCATE", P5LocateLastFailure
        End If
        locateStamp(currentElement) = locateEpoch
        moved = False
        For i = 1 To 3
            ia = mtj(currentElement, i)
            ib = mtj(currentElement, (i Mod 3) + 1)
            If ia < LBound(px) Or ia > UBound(px) Or ib < LBound(px) Or ib > UBound(px) Then
                P5LocateLastFailure = "要素の節点添字が範囲外です。要素=" & CStr(currentElement)
                P5LocateLastIterations = iterationNo
                Err.Raise vbObjectError + 3478, "Mesh.LOCATE", P5LocateLastFailure
            End If
            edgeCross = (px(ia) - xp) * (py(ib) - yp) - (py(ia) - yp) * (px(ib) - xp)
            edgeLength = Sqr((px(ib) - px(ia)) ^ 2 + (py(ib) - py(ia)) ^ 2)
            edgeTolerance = P5GeometryTolerance * edgeLength
            If edgeTolerance < 0.00000000000001 Then edgeTolerance = 0.00000000000001
            If edgeCross < -edgeTolerance Then
                neighborElement = jac(currentElement, i)
                If neighborElement < 1 Or neighborElement > NELM Then
                    P5LocateLastFailure = "領域外の点、または境界外の探索です。要素=" & CStr(currentElement) & "、辺=" & CStr(i)
                    P5LocateLastIterations = iterationNo
                    Err.Raise vbObjectError + 3479, "Mesh.LOCATE", P5LocateLastFailure
                End If
                If locateStamp(neighborElement) = locateEpoch Then
                    P5LocateLastFailure = "要素近傍探索が循環しました。要素=" & CStr(neighborElement)
                    P5LocateLastIterations = iterationNo
                    Err.Raise vbObjectError + 3480, "Mesh.LOCATE", P5LocateLastFailure
                End If
                currentElement = neighborElement
                moved = True
                Exit For
            End If
        Next i
        If Not moved Then
            itri = currentElement
            P5LocateLastIterations = iterationNo
            P5LocateTotalIterations = P5LocateTotalIterations + iterationNo
            If iterationNo > P5LocateMaxIterations Then P5LocateMaxIterations = iterationNo
            P5LocateLastElement = currentElement
            P5LocateLastFailure = vbNullString
            Exit Sub
        End If
    Next iterationNo
    P5LocateLastFailure = "探索回数上限を超えました。上限=" & CStr(maxIterations)
    P5LocateLastIterations = maxIterations
    P5LocateTotalIterations = P5LocateTotalIterations + maxIterations
    If maxIterations > P5LocateMaxIterations Then P5LocateMaxIterations = maxIterations
    Err.Raise vbObjectError + 3481, "Mesh.LOCATE", P5LocateLastFailure

End Sub

Sub swap(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, xp As Double, yp As Double, iswap As Long)

    Dim x13 As Double, y13 As Double, x23 As Double, y23 As Double
    Dim x1p As Double, y1p As Double, x2p As Double, y2p As Double
    Dim cosa As Double, cosb As Double, sina As Double, sinb As Double

    x13 = x1 - x3
    y13 = y1 - y3
    x23 = x2 - x3
    y23 = y2 - y3
    x1p = x1 - xp
    y1p = y1 - yp
    x2p = x2 - xp
    y2p = y2 - yp
    cosa = x13 * x23 + y13 * y23
    cosb = x2p * x1p + y1p * y2p
    
    If (0# <= cosa) And (0# <= cosb) Then
        iswap = 0
    ElseIf (cosa < 0#) And (cosb < 0#) Then
        iswap = 1
    Else
        sina = x13 * y23 - x23 * y13
        sinb = x2p * y1p - x1p * y2p
        If (sina * cosb + sinb * cosa) < 0# Then
            iswap = 1
        Else
            iswap = 0
        End If
    End If

End Sub
Sub edge(l As Long, k As Long, work() As Long, iedge As Long)

    Dim i As Long

    iedge = -1
    For i = 1 To 3
        If work(l, i) = k Then
            iedge = i
            Exit For
        End If
    Next i

    If iedge = -1 Then
        Err.Raise vbObjectError + 3415, "Mesh.EDGE", "要素境界が見つかりません。"
    End If

End Sub

Function IPUSH(item As Long, maxstk As Long, itop As Long) As Long

    If maxstk < itop Then
        Err.Raise vbObjectError + 3416, "Mesh.IPUSH", "メッシュ探索スタックが容量を超えました。"
    Else
        IPUSH = item
    End If

End Function

Function NEIBOR(n As Long, l As Long, jnb() As Long, nei() As Long) As Long

    Dim i As Long

    NEIBOR = -1
    For i = 1 To jnb(n)
Dim iki As Variant
        If nei(n, i) = l Then
            NEIBOR = i
            Exit Function
        End If
    Next i

    ' Show ERROR message and stop the execution
    If NEIBOR = -1 Then
        Err.Raise vbObjectError + 3417, "Mesh.NEIBOR", "近傍要素が見つかりません。"
    End If

End Function

Function IVERT(l As Long, k As Long, mtj() As Long) As Long

    Dim i As Long

    IVERT = -1
    For i = 1 To 3
        If mtj(l, i) = k Then
            IVERT = i
            Exit Function
        End If
    Next i

    ' Show ERROR message and stop the execution
    If IVERT = -1 Then
        Err.Raise vbObjectError + 3490, "Mesh.IVERT", "境界メッシュ処理に失敗しました（IVERT）。"

    End If

End Function
Sub QUEXELM(NELM As Long, mtj() As Long, jac() As Long, kv() As Long, idm() As Long, map() As Long, px() As Double, py() As Double)

    Dim i As Long, j As Long, k As Long
    Dim iv As Long, ip As Long, jn As Long, ie1 As Long, ie2 As Long
    Dim ie3 As Long, ia As Long, ib As Long, ic As Long
    Dim jacia As Long, jacib As Long, inn1 As Long, inn2 As Long
    Dim ien1 As Long, ien2 As Long
    Dim pxa As Double, pya As Double, pxb As Double, pyb As Double
    Dim pxc As Double, pyc As Double, px12 As Double, py12 As Double
    Dim px23 As Double, py23 As Double, pl12 As Double, pl23 As Double
    Dim pl(1 To 3) As Double

    iv = 0
    For i = 1 To NELM
        If mtj(i, 1) = 0 Then GoTo CycleLabel
        If mtj(i, 4) <> 0 Then GoTo CycleLabel

        pxa = px(mtj(i, 1)) - px(mtj(i, 2))
        pya = py(mtj(i, 1)) - py(mtj(i, 2))
        pxb = px(mtj(i, 2)) - px(mtj(i, 3))
        pyb = py(mtj(i, 2)) - py(mtj(i, 3))
        pxc = px(mtj(i, 3)) - px(mtj(i, 1))
        pyc = py(mtj(i, 3)) - py(mtj(i, 1))
        pl(1) = pxa * pxa + pya * pya
        pl(2) = pxb * pxb + pyb * pyb
        pl(3) = pxc * pxc + pyc * pyc

        For j = 1 To 3
            If (pl((j Mod 3) + 1) < pl(j)) And (pl(((j + 1) Mod 3) + 1) < pl(j)) Then
                ip = j
                Exit For
            End If
        Next j

        jn = jac(i, ip)
        
        If jn = 0 Then GoTo CycleLabel
        If idm(i) <> idm(jn) Then GoTo CycleLabel
        If mtj(jn, 4) <> 0 Then GoTo CycleLabel


        Call QUEDGE(jn, i, jac, ie1)

        ie2 = (ie1 Mod 3) + 1
        ie3 = (ie2 Mod 3) + 1
        px12 = px(mtj(jn, ie2)) - px(mtj(jn, ie3))
        py12 = py(mtj(jn, ie2)) - py(mtj(jn, ie3))
        px23 = px(mtj(jn, ie3)) - px(mtj(jn, ie1))
        py23 = py(mtj(jn, ie3)) - py(mtj(jn, ie1))
        pl12 = px12 * px12 + py12 * py12
        pl23 = px23 * px23 + py23 * py23
        If (pl(ip) < pl12) Or (pl(ip) < pl23) Then GoTo CycleLabel
        ia = mtj(i, (ip Mod 3) + 1)
        ib = mtj(i, ((ip + 1) Mod 3) + 1)
        ic = mtj(i, ip)
        jacia = jac(i, (ip Mod 3) + 1)
        jacib = jac(i, ((ip + 1) Mod 3) + 1)
        mtj(i, 1) = ia
        mtj(i, 2) = ib
        mtj(i, 3) = ic
        mtj(i, 4) = mtj(jn, ie3)
        jac(i, 1) = jacia
        jac(i, 2) = jacib
        inn1 = jac(jn, ie2)
        inn2 = jac(jn, ie3)

        If inn1 <> 0 Then
            Call QUEDGE(inn1, jn, jac, ien1)
            jac(inn1, ien1) = i
        End If
        If inn2 <> 0 Then
            Call QUEDGE(inn2, jn, jac, ien2)
            jac(inn2, ien2) = i
        End If
        jac(i, 3) = inn1
        jac(i, 4) = inn2

        For k = 1 To 3
            mtj(jn, k) = 0
            jac(jn, k) = 0
            idm(jn) = 0
        Next k
        iv = iv + 1
        kv(iv) = jn
CycleLabel:
    Next i

    Call QUDELETE(NELM, mtj, jac, idm, iv, kv, map)

End Sub
Sub QUDELETE(NELM As Long, mtj() As Long, jac() As Long, idm() As Long, iv As Long, kv() As Long, map() As Long)
    Dim i As Long, m As Long, n As Long, ia As Long

    m = 0
    n = 0
    
    For i = 1 To NELM
        map(i) = 1
    Next i
    
    For i = 1 To iv
        map(kv(i)) = 0
    Next i
    
    For i = 1 To NELM
        If map(i) <> 0 Then
            m = m + 1
            map(i) = m
        End If
    Next i

    For i = 1 To NELM
        If map(i) <> 0 Then
            n = n + 1
            For ia = 1 To 4
                mtj(n, ia) = mtj(i, ia)
                idm(n) = idm(i)
                If jac(i, ia) = 0 Then
                    jac(n, ia) = 0
                Else
                    jac(n, ia) = map(jac(i, ia))
                End If
            Next ia
        End If
    Next i

    For i = n + 1 To NELM
        For ia = 1 To 4
            mtj(i, ia) = 0
            jac(i, ia) = 0
            idm(i) = 0
        Next ia
    Next i

    NELM = NELM - iv

End Sub
Sub QUISOGEN(node As Long, NELM As Long, mtj() As Long, jac() As Long, px() As Double, py() As Double, id() As Long, mmtj() As Long, _
              ifix() As Long, idm() As Long)
             
    Dim i As Long, j As Long, k As Long, jn As Long, ien As Long
    Dim pxx As Double, pyy As Double
    
    For i = 1 To NELM
        If mtj(i, 4) = 0 Then
            id(i) = 3
            mtj(i, 5) = mtj(i, 3)
            mtj(i, 3) = mtj(i, 2)
            mtj(i, 2) = 0
        Else
            id(i) = 4
            mtj(i, 7) = mtj(i, 4)
            mtj(i, 5) = mtj(i, 3)
            mtj(i, 4) = 0
            mtj(i, 3) = mtj(i, 2)
            mtj(i, 2) = 0
        End If
    Next i
    
    For i = 1 To NELM
        For j = 1 To id(i)
            If mtj(i, 2 * j) <> 0 Then GoTo NextIteration1
            AssertMeshNodeCapacity node + 1
            node = node + 1
            mtj(i, 2 * j) = node
            px(node) = (px(mtj(i, 2 * j - 1)) + px(mtj(i, ((2 * j) Mod (2 * id(i))) + 1))) / 2
            py(node) = (py(mtj(i, 2 * j - 1)) + py(mtj(i, ((2 * j) Mod (2 * id(i))) + 1))) / 2
            If jac(i, j) = 0 Then
                ifix(node) = 1
                GoTo NextIteration1
            End If
            jn = jac(i, j)
            ' Assuming QUEDGE is another subroutine you want to call
            Call QUEDGE(jn, i, jac, ien)
            If mtj(jn, 2 * ien) <> 0 Then GoTo NextIteration1
            mtj(jn, 2 * ien) = node
            If idm(i) <> idm(jn) Then ifix(node) = 1
            
NextIteration1:
        Next j

        AssertMeshNodeCapacity node + 1
        node = node + 1
        mtj(i, 2 * id(i) + 1) = node
        pxx = 0
        pyy = 0
        For k = 1 To id(i)
            pxx = pxx + px(mtj(i, 2 * k))
            pyy = pyy + py(mtj(i, 2 * k))
        Next k
        px(node) = pxx / id(i)
        py(node) = pyy / id(i)
    Next i
    
    ' Assuming QUSQUARE is another subroutine you want to call
    Call QUSQUARE(NELM, mtj, id, mmtj)

End Sub
Sub QUEDGE(l As Long, k As Long, jac() As Long, iedge As Long)
    
    Dim i As Long
    
    iedge = -1
    For i = 1 To 4
        If jac(l, i) = k Then
            iedge = i
            Exit For
        End If
    Next i

    ' Error Message & Exit
    If iedge = -1 Then
        Err.Raise vbObjectError + 3490, "Mesh.QUEDGE", "境界メッシュ処理に失敗しました（QUEDGE）。"

    End If

End Sub

Sub QUSQUARE(NELM As Long, mtj() As Long, id() As Long, mmtj() As Long)
    
    Dim i As Long, j As Long, nelm1 As Long
    
    nelm1 = 0
    For i = 1 To NELM
        If mtj(i, 9) = 0 Then
            For j = 1 To id(i)
                nelm1 = nelm1 + 1
                mmtj(nelm1, 1) = mtj(i, ((2 * j + 3) Mod 6) + 1)
                mmtj(nelm1, 2) = mtj(i, 2 * j - 1)
                mmtj(nelm1, 3) = mtj(i, 2 * j)
                mmtj(nelm1, 4) = mtj(i, 7)
            Next j
        Else
            For j = 1 To id(i)
                nelm1 = nelm1 + 1
                mmtj(nelm1, 1) = mtj(i, ((2 * j + 5) Mod 8) + 1)
                mmtj(nelm1, 2) = mtj(i, 2 * j - 1)
                mmtj(nelm1, 3) = mtj(i, 2 * j)
                mmtj(nelm1, 4) = mtj(i, 9)
            Next j
        End If
    Next i

    NELM = nelm1

End Sub
Sub QUDATASQINPUT(NELM As Long, mtj() As Long, mmtj() As Long)
    Dim i As Long, j As Long
    For i = 1 To NELM
        For j = 1 To 4
            mtj(i, j) = mmtj(i, j)
        Next j
    Next i
End Sub
Sub QUCHECKDATA(node As Long, NNEX, px() As Double, py() As Double, mpx() As Double, mpy() As Double, mibex() As Long, mibno() As Long, ifix() As Long)
Dim no1 As Variant: Dim no2 As Variant
Dim i As Long, j As Long, k As Long


For i = 1 To node
    If ifix(i) = 0 Then
        For j = 1 To NNEX
           For k = 2 To mibex(j) + 1
              no2 = mibno(j, k)
              no1 = mibno(j, k - 1)
              If k = mibex(j) + 1 Then no2 = mibno(j, 1)
              If (py(i) - mpy(no1)) * (mpx(no2) - mpx(no1)) = (mpy(no2) - mpy(no1)) * (px(i) - mpx(no1)) Then ifix(i) = 1
           
           Next k
        Next j
    End If
Next i
    


End Sub
Sub QUDATA(node As Long, NELM As Long, mmtj() As Long, px() As Double, py() As Double, ifix() As Long, idm() As Long)
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim oldNodeCount As Long, oldElementCount As Long, clearNodeRows As Long, clearElementRows As Long
    Dim i As Long, j As Long, validElementCount As Long, newNodeCount As Long
    Dim iarea As Long, area As Double, nodeA As Long, nodeB As Long, midNode As Long
    Dim candidateElements() As Long, usedNode() As Boolean, oldToNew() As Long
    Dim newPx() As Double, newPy() As Double, newIfix() As Long
    Dim nodeOutput() As Variant, elementOutput() As Variant
    Dim edgeLookup As Object, key As String

    Set ws1 = ThisWorkbook.Worksheets("節点データ")
    Set ws2 = ThisWorkbook.Worksheets("要素データ")
    oldNodeCount = ws1.Cells(ws1.rows.count, 1).End(xlUp).row - 1
    oldElementCount = ws2.Cells(ws2.rows.count, 1).End(xlUp).row - 1
    If node < 1 Or NELM < 1 Then Err.Raise vbObjectError + 3440, "Mesh.QUDATA", "節点数または要素数が0です。"

    ReDim candidateElements(1 To NELM)
    ReDim usedNode(1 To node)
    For i = 1 To NELM
        P5EnsureQuadCounterClockwise i, mmtj, px, py
        mtj(i, 1) = mmtj(i, 1): mtj(i, 2) = mmtj(i, 2): mtj(i, 3) = mmtj(i, 3): mtj(i, 4) = mmtj(i, 4)
        area = SQAREA(i, mtj, px, py, iarea)
        If iarea <> 1 Then
            validElementCount = validElementCount + 1
            candidateElements(validElementCount) = i
            For j = 1 To 4
                If mmtj(i, j) < 1 Or mmtj(i, j) > node Then Err.Raise vbObjectError + 3441, "Mesh.QUDATA", "要素接続番号が範囲外です。"
                usedNode(mmtj(i, j)) = True
            Next j
        End If
    Next i
    If validElementCount < 1 Then Err.Raise vbObjectError + 3442, "Mesh.QUDATA", "有効な四角形要素がありません。"

    ReDim oldToNew(1 To node)
    For i = 1 To node
        If usedNode(i) Then
            newNodeCount = newNodeCount + 1
            oldToNew(i) = newNodeCount
        End If
    Next i
    ReDim newPx(1 To UBound(px)): ReDim newPy(1 To UBound(py)): ReDim newIfix(1 To UBound(ifix))
    For i = 1 To node
        If oldToNew(i) > 0 Then
            newPx(oldToNew(i)) = px(i)
            newPy(oldToNew(i)) = py(i)
            newIfix(oldToNew(i)) = ifix(i)
        End If
    Next i
    Set edgeLookup = CreateObject("Scripting.Dictionary")
    edgeLookup.CompareMode = 0
    For i = 1 To validElementCount
        For j = 1 To 4
            pmtj(i, j) = oldToNew(mmtj(candidateElements(i), j))
            smtj(i, j) = pmtj(i, j)
            mtj(i, j) = pmtj(i, j)
        Next j
        For j = 1 To 4
            nodeA = pmtj(i, j): nodeB = pmtj(i, (j Mod 4) + 1)
            If nodeA < nodeB Then key = CStr(nodeA) & ":" & CStr(nodeB) Else key = CStr(nodeB) & ":" & CStr(nodeA)
            If edgeLookup.Exists(key) Then
                midNode = CLng(edgeLookup.item(key))
            Else
                newNodeCount = newNodeCount + 1
                If newNodeCount > UBound(px) Then Err.Raise vbObjectError + 3443, "Mesh.QUDATA", "Q8中間節点の容量を超えました。"
                midNode = newNodeCount
                newPx(midNode) = (newPx(nodeA) + newPx(nodeB)) / 2#
                newPy(midNode) = (newPy(nodeA) + newPy(nodeB)) / 2#
                newIfix(midNode) = 0
                edgeLookup.Add key, midNode
            End If
            mtj(i, j + 4) = midNode
        Next j
    Next i

    node = newNodeCount
    NELM = validElementCount
    For i = 1 To node
        px(i) = newPx(i): py(i) = newPy(i): ifix(i) = newIfix(i)
    Next i
    P5ValidateMeshQuality node, NELM, mtj, px, py
    For i = node + 1 To UBound(px)
        px(i) = 0#: py(i) = 0#: ifix(i) = 0
    Next i
    clearNodeRows = oldNodeCount: If node > clearNodeRows Then clearNodeRows = node
    clearElementRows = oldElementCount: If NELM > clearElementRows Then clearElementRows = NELM
    If clearNodeRows > 0 Then ws1.range(ws1.Cells(2, 1), ws1.Cells(clearNodeRows + 1, 11)).ClearContents
    If clearElementRows > 0 Then ws2.range(ws2.Cells(2, 1), ws2.Cells(clearElementRows + 1, 10)).ClearContents

    ReDim nodeOutput(1 To node, 1 To 5)
    For i = 1 To node
        nodeOutput(i, 1) = i: nodeOutput(i, 2) = px(i): nodeOutput(i, 3) = py(i)
        nodeOutput(i, 4) = ifix(i)
    Next i
    ws1.range(ws1.Cells(2, 1), ws1.Cells(node + 1, 5)).value2 = nodeOutput

    ReDim elementOutput(1 To NELM, 1 To 10)
    For i = 1 To NELM
        elementOutput(i, 1) = i
        For j = 1 To 8: elementOutput(i, j + 1) = mtj(i, j): Next j
        elementOutput(i, 10) = EleCalc(mtj(i, 1), mtj(i, 2), mtj(i, 3), mtj(i, 4))
    Next i
    ws2.range(ws2.Cells(2, 1), ws2.Cells(NELM + 1, 10)).value2 = elementOutput
End Sub

Function EleCalc(p1 As Long, p2 As Long, p3 As Long, p4 As Long) As Long
   Dim i As Long, j As Long
   Dim ax As Double
   Dim ay As Double
   Dim vt As Double
    Dim cn As Long
   Dim zx1 As Double
   Dim zy1 As Double
   Dim zx2 As Double
   Dim zy2 As Double
   
   
   ax = (px(p1) + px(p2) + px(p3) + px(p4)) / 4#
   ay = (py(p1) + py(p2) + py(p3) + py(p4)) / 4#

   
   For i = 1 To NEX
       cn = 0
       For j = 1 To sbex(i)
Dim z1 As Variant: Dim z2 As Variant
           If j = sbex(i) Then
              z2 = mibno(i, 1)
           Else
              z2 = mibno(i, j + 1)
           End If
           z1 = mibno(i, j)
           
           zx1 = sx(z1)
           zy1 = sy(z1)
           
           zx2 = sx(z2)
           zy2 = sy(z2)
           
           
           
           If (zy1 <= ay And zy2 > ay) Or (zy1 > ay And zy2 <= ay) Then
               vt = (ay - zy1) / (zy2 - zy1)
               If ax < zx1 + vt * (zx2 - zx1) Then
                  cn = cn + 1
               End If
           End If
           
       Next j
       If (cn Mod 2) <> 0 Then
       
          EleCalc = i
          Exit For
       End If
   Next i
   If EleCalc = 0 Then
      On Error Resume Next
      EleCalc = MeshReadMaterialNumber()
      On Error GoTo 0
      If EleCalc < 1 Then EleCalc = 1
   End If
End Function

Sub SQINPUT(node As Long, NELM As Long, mtj() As Long, px() As Double, py() As Double, ifix() As Long)

    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim elementValues As Variant, nodeValues As Variant
    Dim i As Long, j As Long
    Erase mtj: Erase ifix: Erase px: Erase py
    Set ws1 = ThisWorkbook.Sheets("節点データ")
    Set ws2 = ThisWorkbook.Sheets("要素データ")
    node = ws1.Cells(ws1.rows.count, 1).End(xlUp).row - 1
    NELM = ws2.Cells(ws2.rows.count, 1).End(xlUp).row - 1
    P5ValidateMeshCapacity node, NELM
    ' Q8接続はB:I、材料番号はJ。ここで材料番号も保持し、
    ' 要素修正後に材料判定を再実行して別材料へ誤変更しない。
    elementValues = ws2.range(ws2.Cells(2, 2), ws2.Cells(NELM + 1, 10)).value2
    For i = 1 To NELM
        For j = 1 To 8: mtj(i, j) = CLng(elementValues(i, j)): Next j
        idm(i) = CLng(elementValues(i, 9))
    Next i
    nodeValues = ws1.range(ws1.Cells(2, 2), ws1.Cells(node + 1, 4)).value2
    For i = 1 To node
        px(i) = CDbl(Val(nodeValues(i, 1))): py(i) = CDbl(Val(nodeValues(i, 2))): ifix(i) = CLng(Val(nodeValues(i, 3)))
    Next i
    P5ValidateMeshQuality node, NELM, mtj, px, py

End Sub
Public Sub P5SelectBestGeneratedMesh()
    Dim kept As String
    ' MESH_GEN_COMPARE=0 still compares; 0 only changes the both-pass tie rule (BetterQuality).
    If Not MeshPrepareCandidate("SMOOTH") Then
        P5WriteTextSetting "MESH_GEN_LAST_METHOD", "COMPARE_FAIL"
        P5WriteTextSetting "MESH_GEN_LAST_COMPARE", MeshAdaptLastError
        Exit Sub
    End If
    kept = MeshAdaptSelectedMethod()
    ' Write LAST_* after apply. LAST_COMPARE is hashed by MeshInputSignature unless skipped,
    ' and writing it first made apply abort with "入力が変わりました".
    If kept = "既存接続" Then
        P5WriteTextSetting "MESH_GEN_LAST_METHOD", kept
        P5WriteTextSetting "MESH_GEN_LAST_COMPARE", MeshAdaptLastSummary
        Exit Sub
    End If
    If Not MeshApplyCandidate(True) Then
        P5WriteTextSetting "MESH_GEN_LAST_METHOD", "COMPARE_APPLY_FAIL"
        P5WriteTextSetting "MESH_GEN_LAST_COMPARE", MeshAdaptLastError & vbCrLf & MeshAdaptLastSummary
        Exit Sub
    End If
    P5WriteTextSetting "MESH_GEN_LAST_METHOD", kept
    P5WriteTextSetting "MESH_GEN_LAST_COMPARE", MeshAdaptLastSummary
    If Not SuppressUserMessages Then
        MsgBox "メッシュ比較で採用：" & kept & vbCrLf & vbCrLf & MeshAdaptLastSummary, vbInformation
    End If
End Sub

Function SQAREA(ielm As Long, mtj() As Long, px() As Double, py() As Double, iarea As Long) As Double
    Dim j1 As Long, j2 As Long, j3 As Long, j4 As Long
    Dim area1 As Double, area2 As Double, area3 As Double, area4 As Double

    iarea = 1
    SQAREA = 0#
    If ielm < LBound(mtj, 1) Or ielm > UBound(mtj, 1) Then Exit Function
    j1 = mtj(ielm, 1)
    j2 = mtj(ielm, 2)
    j3 = mtj(ielm, 3)
    j4 = mtj(ielm, 4)
    If j1 < 1 Or j2 < 1 Or j3 < 1 Or j4 < 1 Then Exit Function
    If j1 > UBound(px) Or j2 > UBound(px) Or j3 > UBound(px) Or j4 > UBound(px) Then Exit Function
    area1 = 0.5 * (px(j1) * py(j2) + px(j2) * py(j3) + px(j3) * py(j1) - px(j1) * py(j3) - px(j2) * py(j1) - px(j3) * py(j2))
    area2 = 0.5 * (px(j1) * py(j3) + px(j3) * py(j4) + px(j4) * py(j1) - px(j1) * py(j4) - px(j3) * py(j1) - px(j4) * py(j3))
    j1 = mtj(ielm, 2)
    j2 = mtj(ielm, 3)
    j3 = mtj(ielm, 4)
    j4 = mtj(ielm, 1)
    area3 = 0.5 * (px(j1) * py(j2) + px(j2) * py(j3) + px(j3) * py(j1) - px(j1) * py(j3) - px(j2) * py(j1) - px(j3) * py(j2))
    area4 = 0.5 * (px(j1) * py(j3) + px(j3) * py(j4) + px(j4) * py(j1) - px(j1) * py(j4) - px(j3) * py(j1) - px(j4) * py(j3))
    
    If (area1 > 0 And area2 > 0 And area3 > 0 And area4 > 0) Then iarea = 0
    
    SQAREA = area1 + area2
End Function

Private Sub P5EnsureQuadCounterClockwise(ByVal elementId As Long, ByRef connectivity() As Long, ByRef x() As Double, ByRef y() As Double)
    Dim iarea As Long, swapValue As Long
    SQAREA elementId, connectivity, x, y, iarea
    If iarea = 0 Then Exit Sub
    swapValue = connectivity(elementId, 2)
    connectivity(elementId, 2) = connectivity(elementId, 4)
    connectivity(elementId, 4) = swapValue
End Sub









Private Function P5BoundaryReadWidth(ByVal ws As Worksheet, ByVal externalCount As Long, ByVal internalCount As Long) As Long
    Dim rows As Variant, i As Long, n As Long
    rows = ws.range(ws.Cells(1, 5), ws.Cells(23 + internalCount, 5)).value2
    P5BoundaryReadWidth = 6
    For i = 1 To externalCount + internalCount
        If i <= externalCount Then n = CLng(Val(rows(i + 2, 1))) Else n = CLng(Val(rows(23 + i - externalCount, 1)))
        If n < 0 Or n > 1999 Then Err.Raise vbObjectError + 3475, , "境界節点数が容量を超えています。"
        If n + 6 > P5BoundaryReadWidth Then P5BoundaryReadWidth = n + 6
    Next i
End Function

Public Function MeshSheetHasJointElements() As Boolean
    Dim wsM As Worksheet, wsE As Worksheet
    Dim lastM As Long, lastE As Long, rowNo As Long, matId As Long
    Dim kindById() As String, rawKind As String
    MeshSheetHasJointElements = False
    On Error Resume Next
    Set wsM = ThisWorkbook.Worksheets("材料データ")
    Set wsE = ThisWorkbook.Worksheets("要素データ")
    On Error GoTo 0
    If wsM Is Nothing Or wsE Is Nothing Then Exit Function
    lastM = wsM.Cells(wsM.rows.count, 1).End(xlUp).row
    lastE = wsE.Cells(wsE.rows.count, 1).End(xlUp).row
    If lastM < 2 Or lastE < 2 Then Exit Function
    ReDim kindById(1 To lastM - 1)
    For rowNo = 2 To lastM
        matId = CLng(Val(CStr(wsM.Cells(rowNo, 1).value2)))
        If matId >= 1 And matId <= lastM - 1 Then
            rawKind = UCase$(Trim$(CStr(wsM.Cells(rowNo, 11).value2)))
            If rawKind = "JOINT" Or rawKind = "接合" Then kindById(matId) = "JOINT"
        End If
    Next rowNo
    For rowNo = 2 To lastE
        matId = CLng(Val(CStr(wsE.Cells(rowNo, 10).value2)))
        If matId >= 1 And matId <= UBound(kindById) Then
            If kindById(matId) = "JOINT" Then
                MeshSheetHasJointElements = True
                Exit Function
            End If
        End If
        If Len(Trim$(CStr(wsE.Cells(rowNo, 8).value2))) = 0 Or Len(Trim$(CStr(wsE.Cells(rowNo, 9).value2))) = 0 Then
            If CLng(Val(CStr(wsE.Cells(rowNo, 2).value2))) > 0 Then
                MeshSheetHasJointElements = True
                Exit Function
            End If
        End If
    Next rowNo
End Function

Public Function MeshJointPairsDefined() As Boolean
    Dim wsJ As Worksheet, lastJ As Long, rowNo As Long
    Dim matA As Long, matB As Long, Jmat As Long, enabled As Boolean
    MeshJointPairsDefined = False
    On Error Resume Next
    Set wsJ = ThisWorkbook.Worksheets("接合")
    On Error GoTo 0
    If wsJ Is Nothing Then Exit Function
    lastJ = wsJ.Cells(wsJ.rows.count, 1).End(xlUp).row
    For rowNo = 2 To lastJ
        matA = CLng(Val(CStr(wsJ.Cells(rowNo, 1).value2)))
        matB = CLng(Val(CStr(wsJ.Cells(rowNo, 2).value2)))
        Jmat = CLng(Val(CStr(wsJ.Cells(rowNo, 3).value2)))
        enabled = MeshParseEnabled(wsJ.Cells(rowNo, 4).value2)
        If enabled And matA > 0 And matB > 0 And Jmat > 0 And matA <> matB Then
            MeshJointPairsDefined = True
            Exit Function
        End If
    Next rowNo
End Function

Private Function MeshParseEnabled(ByVal rawValue As Variant) As Boolean
    Dim token As String
    MeshParseEnabled = True
    If IsError(rawValue) Then Exit Function
    If IsEmpty(rawValue) Then Exit Function
    If IsNumeric(rawValue) Then
        MeshParseEnabled = (Abs(CDbl(rawValue)) >= 0.5)
        Exit Function
    End If
    token = UCase$(Trim$(CStr(rawValue)))
    If token = "0" Or token = "OFF" Or token = "NO" Or token = "N" Or token = "しない" Or token = "無効" Then MeshParseEnabled = False
End Function

Public Sub MeshInsertJointsIfNeeded()
    If Not MeshJointPairsDefined() Then Exit Sub
    If MeshSheetHasJointElements() Then Exit Sub
    MeshInsertJoints
End Sub

Public Sub MeshInsertJoints()
    Dim wsJ As Worksheet, wsN As Worksheet, wsE As Worksheet, wsM As Worksheet
    Dim lastJ As Long, lastN As Long, lastE As Long, lastM As Long
    Dim rowNo As Long, pairN As Long, p As Long, e As Long, ee As Long, k As Long
    Dim matA As Long, matB As Long, Jmat As Long
    Dim nodeCount As Long, elemCount As Long, newNode As Long
    Dim c1 As Long, c2 As Long, midN As Long, nodeId As Long
    Dim key As String, info As String, parts() As String
    Dim masterElem As Long, masterEdge As Long, slaveElem As Long
    Dim origNode As Long, dupNode As Long, added As Long
    Dim kindById() As String, rawKind As String
    Dim pairA() As Long, pairB() As Long, pairJ() As Long
    Dim nodeX() As Double, nodeY() As Double
    Dim condX() As Long, condY() As Long
    Dim conn() As Long, MatNo() As Long
    Dim origConn() As Long
    Dim jointConn() As Long, jointMat() As Long
    Dim masterEdges As Object, slaveDup As Object
    Dim nodeOut() As Variant, elemOut() As Variant
    Dim n1 As Long, n2 As Long, n3 As Long, N4 As Long, n5 As Long, n6 As Long
    Dim origCount As Long, maxAdd As Long

    On Error Resume Next
    Set wsJ = ThisWorkbook.Worksheets("接合")
    On Error GoTo 0
    If wsJ Is Nothing Then Exit Sub
    Set wsN = ThisWorkbook.Worksheets("節点データ")
    Set wsE = ThisWorkbook.Worksheets("要素データ")
    Set wsM = ThisWorkbook.Worksheets("材料データ")
    lastJ = wsJ.Cells(wsJ.rows.count, 1).End(xlUp).row
    If lastJ < 2 Then Exit Sub
    lastN = wsN.Cells(wsN.rows.count, 1).End(xlUp).row
    lastE = wsE.Cells(wsE.rows.count, 1).End(xlUp).row
    lastM = wsM.Cells(wsM.rows.count, 1).End(xlUp).row
    If lastN < 2 Or lastE < 2 Then Exit Sub

    ReDim kindById(1 To 1)
    If lastM >= 2 Then
        ReDim kindById(1 To lastM - 1)
        For rowNo = 2 To lastM
            matA = CLng(Val(CStr(wsM.Cells(rowNo, 1).value2)))
            If matA >= 1 Then
                If matA > UBound(kindById) Then ReDim Preserve kindById(1 To matA)
                rawKind = UCase$(Trim$(CStr(wsM.Cells(rowNo, 11).value2)))
                If rawKind = "JOINT" Or rawKind = "接合" Then
                    kindById(matA) = "JOINT"
                ElseIf rawKind = "STRUCT" Or rawKind = "構造" Then
                    kindById(matA) = "STRUCT"
                Else
                    kindById(matA) = "SOIL"
                End If
            End If
        Next rowNo
    End If

    ReDim pairA(1 To lastJ): ReDim pairB(1 To lastJ): ReDim pairJ(1 To lastJ)
    pairN = 0
    For rowNo = 2 To lastJ
        If Not MeshParseEnabled(wsJ.Cells(rowNo, 4).value2) Then GoTo NextJointPair
        matA = CLng(Val(CStr(wsJ.Cells(rowNo, 1).value2)))
        matB = CLng(Val(CStr(wsJ.Cells(rowNo, 2).value2)))
        Jmat = CLng(Val(CStr(wsJ.Cells(rowNo, 3).value2)))
        If matA < 1 Or matB < 1 Or Jmat < 1 Or matA = matB Then GoTo NextJointPair
        If Jmat <= UBound(kindById) Then
            If kindById(Jmat) <> "JOINT" Then Err.Raise vbObjectError + 3488, "Mesh.MeshInsertJoints", "接合材料は種別JOINTにしてください。材料=" & CStr(Jmat)
        End If
        pairN = pairN + 1
        pairA(pairN) = matA
        pairB(pairN) = matB
        pairJ(pairN) = Jmat
NextJointPair:
    Next rowNo
    If pairN < 1 Then Exit Sub
    If MeshSheetHasJointElements() Then Exit Sub

    nodeCount = lastN - 1
    elemCount = lastE - 1
    ReDim nodeX(1 To nodeCount)
    ReDim nodeY(1 To nodeCount)
    ReDim condX(1 To nodeCount)
    ReDim condY(1 To nodeCount)
    For rowNo = 1 To nodeCount
        nodeX(rowNo) = CDbl(wsN.Cells(rowNo + 1, 2).value2)
        nodeY(rowNo) = CDbl(wsN.Cells(rowNo + 1, 3).value2)
        condX(rowNo) = CLng(Val(CStr(wsN.Cells(rowNo + 1, 4).value2)))
        condY(rowNo) = CLng(Val(CStr(wsN.Cells(rowNo + 1, 5).value2)))
    Next rowNo
    ReDim conn(1 To elemCount, 1 To 8)
    ReDim origConn(1 To elemCount, 1 To 8)
    ReDim MatNo(1 To elemCount)
    For rowNo = 1 To elemCount
        For k = 1 To 8
            conn(rowNo, k) = CLng(Val(CStr(wsE.Cells(rowNo + 1, k + 1).value2)))
            origConn(rowNo, k) = conn(rowNo, k)
        Next k
        MatNo(rowNo) = CLng(Val(CStr(wsE.Cells(rowNo + 1, 10).value2)))
    Next rowNo

    Set slaveDup = CreateObject("Scripting.Dictionary")
    slaveDup.CompareMode = 0
    added = 0
    origCount = elemCount
    maxAdd = origCount * 4
    If maxAdd < 8 Then maxAdd = 8
    ReDim jointConn(1 To maxAdd, 1 To 8)
    ReDim jointMat(1 To maxAdd)

    For p = 1 To pairN
        Set masterEdges = CreateObject("Scripting.Dictionary")
        masterEdges.CompareMode = 0
        For e = 1 To origCount
            If MatNo(e) = pairA(p) Then
                For k = 1 To 4
                    c1 = origConn(e, k)
                    c2 = origConn(e, (k Mod 4) + 1)
                    If c1 > 0 And c2 > 0 Then
                        If c1 < c2 Then key = CStr(c1) & ":" & CStr(c2) Else key = CStr(c2) & ":" & CStr(c1)
                        If Not masterEdges.Exists(key) Then masterEdges.Add key, CStr(e) & "|" & CStr(k)
                    End If
                Next k
            End If
        Next e
        For e = 1 To origCount
            If MatNo(e) = pairB(p) Then
                For k = 1 To 4
                    c1 = origConn(e, k)
                    c2 = origConn(e, (k Mod 4) + 1)
                    If c1 < 1 Or c2 < 1 Then GoTo NextSlaveEdge
                    If c1 < c2 Then key = CStr(c1) & ":" & CStr(c2) Else key = CStr(c2) & ":" & CStr(c1)
                    If Not masterEdges.Exists(key) Then GoTo NextSlaveEdge
                    info = CStr(masterEdges.item(key))
                    parts = Split(info, "|")
                    masterElem = CLng(parts(0))
                    masterEdge = CLng(parts(1))
                    n1 = origConn(masterElem, masterEdge)
                    n2 = origConn(masterElem, (masterEdge Mod 4) + 1)
                    n3 = origConn(masterElem, masterEdge + 4)
                    N4 = MeshJointDupNode(n1, nodeCount, nodeX, nodeY, condX, condY, slaveDup)
                    n5 = MeshJointDupNode(n2, nodeCount, nodeX, nodeY, condX, condY, slaveDup)
                    n6 = MeshJointDupNode(n3, nodeCount, nodeX, nodeY, condX, condY, slaveDup)
                    For ee = 1 To 8
                        If origConn(e, ee) = n1 Then conn(e, ee) = N4
                        If origConn(e, ee) = n2 Then conn(e, ee) = n5
                        If origConn(e, ee) = n3 Then conn(e, ee) = n6
                    Next ee
                    added = added + 1
                    If added > maxAdd Then Err.Raise vbObjectError + 3489, "Mesh.MeshInsertJoints", "接合要素の容量を超えました。"
                    jointConn(added, 1) = n1
                    jointConn(added, 2) = n2
                    jointConn(added, 3) = n3
                    jointConn(added, 4) = N4
                    jointConn(added, 5) = n5
                    jointConn(added, 6) = n6
                    jointConn(added, 7) = 0
                    jointConn(added, 8) = 0
                    jointMat(added) = pairJ(p)
NextSlaveEdge:
                Next k
            End If
        Next e
    Next p

    If added < 1 Then Exit Sub
    elemCount = origCount + added
    If nodeCount > 60000 Then Err.Raise vbObjectError + 3490, "Mesh.MeshInsertJoints", "接合挿入後の節点数が上限を超えます。"
    If elemCount > 60000 Then Err.Raise vbObjectError + 3491, "Mesh.MeshInsertJoints", "接合挿入後の要素数が上限を超えます。"

    If lastN - 1 > 0 Then wsN.range(wsN.Cells(2, 1), wsN.Cells(lastN, 9)).ClearContents
    If lastE - 1 > 0 Then wsE.range(wsE.Cells(2, 1), wsE.Cells(lastE, 10)).ClearContents
    ReDim nodeOut(1 To nodeCount, 1 To 9)
    For rowNo = 1 To nodeCount
        nodeOut(rowNo, 1) = rowNo
        nodeOut(rowNo, 2) = nodeX(rowNo)
        nodeOut(rowNo, 3) = nodeY(rowNo)
        nodeOut(rowNo, 4) = condX(rowNo)
        nodeOut(rowNo, 5) = condY(rowNo)
        nodeOut(rowNo, 6) = 0
        nodeOut(rowNo, 7) = 0
        nodeOut(rowNo, 8) = 0
        nodeOut(rowNo, 9) = 0
    Next rowNo
    wsN.range(wsN.Cells(2, 1), wsN.Cells(nodeCount + 1, 9)).value2 = nodeOut
    ReDim elemOut(1 To elemCount, 1 To 10)
    For rowNo = 1 To origCount
        elemOut(rowNo, 1) = rowNo
        For k = 1 To 8
            If conn(rowNo, k) > 0 Then elemOut(rowNo, k + 1) = conn(rowNo, k)
        Next k
        elemOut(rowNo, 10) = MatNo(rowNo)
    Next rowNo
    For rowNo = 1 To added
        elemOut(origCount + rowNo, 1) = origCount + rowNo
        For k = 1 To 8
            If jointConn(rowNo, k) > 0 Then elemOut(origCount + rowNo, k + 1) = jointConn(rowNo, k)
        Next k
        elemOut(origCount + rowNo, 10) = jointMat(rowNo)
    Next rowNo
    wsE.range(wsE.Cells(2, 1), wsE.Cells(elemCount + 1, 10)).value2 = elemOut
    FEMInvalidateInputCache
End Sub

Private Function MeshJointDupNode(ByVal origNode As Long, ByRef nodeCount As Long, ByRef nodeX() As Double, ByRef nodeY() As Double, ByRef condX() As Long, ByRef condY() As Long, ByVal slaveDup As Object) As Long
    Dim key As String, dupNode As Long
    If origNode < 1 Then
        MeshJointDupNode = 0
        Exit Function
    End If
    key = CStr(origNode)
    If slaveDup.Exists(key) Then
        MeshJointDupNode = CLng(slaveDup.item(key))
        Exit Function
    End If
    nodeCount = nodeCount + 1
    ReDim Preserve nodeX(1 To nodeCount)
    ReDim Preserve nodeY(1 To nodeCount)
    ReDim Preserve condX(1 To nodeCount)
    ReDim Preserve condY(1 To nodeCount)
    nodeX(nodeCount) = nodeX(origNode)
    nodeY(nodeCount) = nodeY(origNode)
    condX(nodeCount) = 0
    condY(nodeCount) = 0
    slaveDup.Add key, nodeCount
    MeshJointDupNode = nodeCount
End Function

