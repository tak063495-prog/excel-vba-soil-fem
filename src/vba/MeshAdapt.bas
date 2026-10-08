
Option Explicit
Private sourceGrid As Object, sourceCell As Double, sourceEpoch As Long
Private srcLoX() As Double, srcHiX() As Double, srcLoY() As Double, srcHiY() As Double
Private srcX() As Double, srcY() As Double, srcMat() As Long, srcSeen() As Long
Private terrainPrepared As Boolean, terrainPX() As Double, terrainPY() As Double, terrainNP As Long
Private terrainLeft As Double, terrainRight As Double, terrainBottom As Double
Private markHeap() As Double, markDue() As Double, markHeapN As Long, markLog() As Long, markLogN As Long

Private Type CandidateState
    countN As Long
    countE As Long
    x() As Double
    y() As Double
    nodes() As Variant
    elements() As Long
    materials() As Long
    score() As Double
    rebuilt As Boolean
    steps As Long
End Type
Private Type QualityRecord
    passN As Long
    warnN As Long
    failN As Long
    minA As Double
    maxA As Double
    aspect As Double
    scaledJ As Double
    grading As Double
    heatArea As Double
    countN As Long
    countE As Long
End Type
Private jacReady As Boolean, jacDu(0 To 24, 0 To 7) As Double, jacDv(0 To 24, 0 To 7) As Double
Private metricMax() As Double, metricSJ() As Double, metricClass() As Long
Private geomVersion As Long, heatVersion As Long, heatCache As Double
Private candidateRebuilt As Boolean, selectedMethod As String, comparisonText As String
Private sourceIndicator() As Double, parentNodes As Variant, parentElements As Variant
Private terrainReason As String, plasticHeatLimit As Double
Private segN As Long, sx1() As Double, sy1() As Double, sx2() As Double, sy2() As Double, sFixX() As Long, sFixY() As Long



Public MeshAdaptLastError As String
Public MeshAdaptLastSummary As String
Public MeshAdaptLastBackup As String
Private busy As Boolean, candidateReady As Boolean
Private nn As Long, ne As Long, originalNN As Long, originalNE As Long
Private xx() As Double, yy() As Double, conn() As Long, materialId() As Long
Private nodeRows() As Variant, oldNodes As Variant, oldElements As Variant
Private edgeMap As Object, edgeN As Long, ea() As Long, eb() As Long, em() As Long
Private ee1() As Long, ee2() As Long, el1() As Long, el2() As Long, elementEdge() As Long
Private areas() As Double, energies() As Double, angles() As Double, aspects() As Double
Private incident() As Collection, locked() As Boolean, slideA() As Long, slideB() As Long
Private candidateOptions As String
Private candidateSignature As String, candidateKind As String, meshScale As Double
Private movedCount As Long, sweepCount As Long, beforeAngle As Double, beforeAspect As Double
Private operationStep As String
Private oldAreaByMaterial As Object, newElements() As Variant
Private forceLimits As Object
Private refineScore() As Double, candidateScore() As Double
Private estimatedMemory As Double, estimatedBand As Long

Private Function LoadRowKey(ByRef rows As Variant, ByVal r As Long) As String
    Dim a As Double, b As Double, j As Long
    a = 17#: b = 31#
    For j = 1 To 9: HashText CStr(rows(r, j)) & "|", a, b: Next j
    LoadRowKey = Format$(a, "0") & "-" & Format$(b, "0")
End Function

Public Sub MeshLoadMapBegin()
    Dim ws As Worksheet, rows As Variant, data As Variant, r As Long, key As String
    Set forceLimits = CreateObject("Scripting.Dictionary")
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("メッシュ引継"): On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    rows = Table("載荷", 9): data = Table("メッシュ引継", 3)
    For r = 1 To UBound(data, 1)
        If CStr(data(r, 1)) = "FORCE" Then forceLimits(CStr(data(r, 2))) = CLng(data(r, 3))
    Next r
    For r = 1 To UBound(rows, 1)
        key = LoadRowKey(rows, r)
        If forceLimits.Exists(key) Then forceLimits("row" & CStr(r + 1)) = forceLimits(key)
    Next r
End Sub

Public Function MeshForceNodeAllowed(ByVal row As Long, ByVal originalNode As Long) As Boolean
    MeshForceNodeAllowed = True
    If forceLimits Is Nothing Then Exit Function
    If forceLimits.Exists("row" & CStr(row)) Then MeshForceNodeAllowed = (originalNode <= CLng(forceLimits("row" & CStr(row))))
End Function

Private Sub SaveForceAnchors()
    Dim ws As Worksheet, rows As Variant, data As Variant, old As Object, r As Long, writeRow As Long, key As String, limit As Long
    Set old = CreateObject("Scripting.Dictionary")
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("メッシュ引継"): On Error GoTo 0
    If Not ws Is Nothing Then
        data = Table("メッシュ引継", 3)
        For r = 1 To UBound(data, 1)
            If CStr(data(r, 1)) = "FORCE" Then old(CStr(data(r, 2))) = CLng(data(r, 3))
        Next r
    End If
    rows = Table("載荷", 9)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = "メッシュ引継"
    ws.Cells.ClearContents: ws.Cells(1, 1).value2 = "種類": ws.Cells(1, 2).value2 = "載荷行の署名": ws.Cells(1, 3).value2 = "元の節点上限"
    writeRow = 2
    For r = 1 To UBound(rows, 1)
        If UCase$(Trim$(CStr(rows(r, 4)))) = "FORCE" Then
            key = LoadRowKey(rows, r): limit = originalNN
            If old.Exists(key) Then limit = CLng(old(key))
            ws.Cells(writeRow, 1).value2 = "FORCE": ws.Cells(writeRow, 2).value2 = key: ws.Cells(writeRow, 3).value2 = limit
            writeRow = writeRow + 1
        End If
    Next r
    ws.Visible = xlSheetVeryHidden
End Sub

Private Function Table(ByVal sheetName As String, ByVal cols As Long) As Variant
    Dim ws As Worksheet, last As Long
    Set ws = ThisWorkbook.Worksheets(sheetName)
    last = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If last < 2 Then last = 2
    Table = ws.range(ws.Cells(2, 1), ws.Cells(last, cols)).value2
End Function

Private Function d(ByVal v As Variant) As Double
    If IsError(v) Then Err.Raise vbObjectError + 3800, "MeshAdapt", "入力セルにエラーがあります。"
    If IsEmpty(v) Or Len(CStr(v)) = 0 Then Exit Function
    If Not IsNumeric(v) Then Err.Raise vbObjectError + 3800, "MeshAdapt", "数値入力が不正です。"
    d = CDbl(v)
End Function

Private Function EdgeKey(ByVal a As Long, ByVal b As Long) As String
    If a < b Then EdgeKey = CStr(a) & ":" & CStr(b) Else EdgeKey = CStr(b) & ":" & CStr(a)
End Function

Private Sub HashText(ByVal s As String, ByRef a As Double, ByRef b As Double)
    Dim i As Long, c As Long
    For i = 1 To Len(s)
        c = AscW(mid$(s, i, 1)): If c < 0 Then c = c + 65536
        a = a * 257# + c + 1#: a = a - Fix(a / 2147483629#) * 2147483629#
        b = b * 263# + c + 1#: b = b - Fix(b / 2147483587#) * 2147483587#
    Next i
End Sub

Public Function MeshInputSignature() As String
    Dim names As Variant, widths As Variant, rows As Variant, i As Long, r As Long, c As Long
    Dim a As Double, b As Double, token As String, settingKey As String, meta As Worksheet
    names = Array("節点データ", "要素データ", "材料データ", "ステージ", "載荷", "設定")
    widths = Array(14, 10, 10, 5, 9, 5): a = 17#: b = 31#
    For i = 0 To UBound(names)
        rows = Table(CStr(names(i)), CLng(widths(i)))
        HashText CStr(names(i)) & ":", a, b
        For r = 1 To UBound(rows, 1)
            If i = 5 Then
                settingKey = UCase$(Trim$(CStr(rows(r, 5))))
                If Len(settingKey) = 0 Then GoTo NextRow
                If left$(settingKey, 5) = "VIEW_" Or left$(settingKey, 13) = "MESH_QUALITY_" Or left$(settingKey, 6) = "ADAPT_" Or left$(settingKey, 4) = "RCM_" Or left$(settingKey, 12) = "MESH_REPAIR_" Or left$(settingKey, 14) = "MESH_GEN_LAST_" Then GoTo NextRow
                HashText settingKey & "=" & CStr(rows(r, 3)) & ";", a, b
            Else
                For c = 1 To UBound(rows, 2)
                    If IsError(rows(r, c)) Then Err.Raise vbObjectError + 3800, , "入力にセルエラーがあります。"
                    token = CStr(rows(r, c))
                    HashText CStr(Len(token)) & ":" & token & "|", a, b
                Next c
            End If
NextRow:
        Next r
    Next i
    On Error Resume Next: Set meta = ThisWorkbook.Worksheets("メッシュ引継"): On Error GoTo 0
    If Not meta Is Nothing Then
        rows = Table("メッシュ引継", 3)
        For r = 1 To UBound(rows, 1)
            If CStr(rows(r, 1)) = "FORCE" Then
                HashText "FORCE:" & CStr(rows(r, 2)) & ":" & CStr(rows(r, 3)), a, b
            End If
        Next r
    End If
    MeshInputSignature = Format$(a, "0") & "-" & Format$(b, "0")
End Function

Private Function SavedSignature() As String
    On Error Resume Next
    SavedSignature = Replace$(Replace$(ThisWorkbook.names("_FEMResultSignature").RefersTo, "=", ""), """", "")
    On Error GoTo 0
End Function

Public Sub MeshStampResult()
    Dim sig As String
    sig = MeshInputSignature()
    On Error Resume Next
    ThisWorkbook.names("_FEMResultSignature").Delete
    On Error GoTo 0
    ThisWorkbook.names.Add name:="_FEMResultSignature", RefersTo:="=""" & sig & """", Visible:=False
End Sub

Private Sub LoadMesh()
    Dim i As Long, j As Long, k As Long, cap As Long, loX As Double, hiX As Double, loY As Double, hiY As Double
    oldNodes = Table("節点データ", 14): oldElements = Table("要素データ", 10)
    originalNN = UBound(oldNodes, 1): originalNE = UBound(oldElements, 1)
    nn = originalNN: ne = originalNE
    candidateRebuilt = False: plasticHeatLimit = 0#: comparisonText = ""
    geomVersion = 0: heatVersion = -1
    Set sourceGrid = Nothing
    terrainPrepared = False
    If nn < 8 Or ne < 1 Then Err.Raise vbObjectError + 3801, , "Q8メッシュがありません。"
    cap = nn + ne * 25
    If cap > 300000 Or ne > 50000 Then Err.Raise vbObjectError + 3801, , "再メッシュの作業上限を超えます。"
    ReDim xx(1 To cap): ReDim yy(1 To cap): ReDim nodeRows(1 To cap, 1 To 14)
    ReDim conn(1 To ne * 4, 1 To 8): ReDim materialId(1 To ne * 4)
    ReDim candidateScore(1 To ne * 4)
    For i = 1 To nn
        If d(oldNodes(i, 1)) <> i Then Err.Raise vbObjectError + 3801, , "節点番号は行順の1始まり連番にしてください。"
        For j = 1 To 14: nodeRows(i, j) = oldNodes(i, j): Next j
        xx(i) = d(oldNodes(i, 2)): yy(i) = d(oldNodes(i, 3))
        If i = 1 Then loX = xx(i): hiX = xx(i): loY = yy(i): hiY = yy(i)
        If xx(i) < loX Then loX = xx(i)
        If xx(i) > hiX Then hiX = xx(i)
        If yy(i) < loY Then loY = yy(i)
        If yy(i) > hiY Then hiY = yy(i)
    Next i
    meshScale = hiX - loX: If hiY - loY > meshScale Then meshScale = hiY - loY
    If meshScale <= 0# Then Err.Raise vbObjectError + 3801, , "座標範囲が不正です。"
    For i = 1 To ne
        If d(oldElements(i, 1)) <> i Then Err.Raise vbObjectError + 3801, , "要素番号は行順の1始まり連番にしてください。"
        For j = 1 To 8
            k = CLng(d(oldElements(i, j + 1)))
            If k < 1 Or k > nn Then Err.Raise vbObjectError + 3801, , "要素の接続節点が不正です。"
            conn(i, j) = k
        Next j
        materialId(i) = CLng(d(oldElements(i, 10)))
        If materialId(i) < 1 Then Err.Raise vbObjectError + 3801, , "材料番号が不正です。"
    Next i
    BuildEdges
    CacheMetrics
    Set oldAreaByMaterial = MaterialAreas()
    movedCount = 0: sweepCount = 0
    beforeAngle = MinAngle(): beforeAspect = MaxAspect()
End Sub

Private Sub BuildEdges()
    Dim e As Long, j As Long, a As Long, b As Long, m As Long, k As Long, key As String
    Set edgeMap = CreateObject("Scripting.Dictionary")
    ReDim ea(1 To ne * 4): ReDim eb(1 To ne * 4): ReDim em(1 To ne * 4)
    ReDim ee1(1 To ne * 4): ReDim ee2(1 To ne * 4): ReDim el1(1 To ne * 4): ReDim el2(1 To ne * 4)
    ReDim elementEdge(1 To ne, 0 To 3): edgeN = 0
    ReDim incident(1 To nn)
    For j = 1 To nn: Set incident(j) = New Collection: Next j
    For e = 1 To ne
        For j = 0 To 3
            a = conn(e, j + 1): b = conn(e, (j + 1) Mod 4 + 1): m = conn(e, j + 5)
            incident(a).Add e
            key = EdgeKey(a, b)
            If edgeMap.Exists(key) Then
                k = edgeMap(key)
                If ee2(k) <> 0 Or em(k) <> m Then Err.Raise vbObjectError + 3802, , "辺の共有状態またはQ8中間節点が不正です。"
                ee2(k) = e: el2(k) = j
            Else
                edgeN = edgeN + 1: k = edgeN: edgeMap.Add key, k
                ea(k) = a: eb(k) = b: em(k) = m: ee1(k) = e: el1(k) = j
            End If
            elementEdge(e, j) = k
        Next j
    Next e
End Sub

Private Function AcosValue(ByVal v As Double) As Double
    If v >= 1# Then Exit Function
    If v <= -1# Then AcosValue = 3.14159265358979: Exit Function
    AcosValue = 1.5707963267949 - Atn(v / Sqr(1# - v * v))
End Function

Private Function Metric(ByVal e As Long, ByRef ar As Double, ByRef q As Double, ByRef amin As Double, ByRef asp As Double) As Boolean
    Dim j As Long, a As Long, b As Long, c As Long, ux As Double, uy As Double, vx As Double, vy As Double
    Dim cross As Double, l1 As Double, l2 As Double, shortest As Double, longest As Double, angle As Double
    ar = 0#: q = 0#: amin = 180#: shortest = 1E+100: longest = 0#
    metricMax(e) = 0#: metricSJ(e) = 1#: metricClass(e) = 0
    For j = 0 To 3
        a = conn(e, j + 1): b = conn(e, (j + 1) Mod 4 + 1): c = conn(e, (j + 3) Mod 4 + 1)
        ux = xx(b) - xx(a): uy = yy(b) - yy(a): vx = xx(c) - xx(a): vy = yy(c) - yy(a)
        cross = ux * vy - uy * vx
        If cross <= meshScale * meshScale * 0.00000000000001 Then Exit Function
        l1 = ux * ux + uy * uy: l2 = vx * vx + vy * vy
        If l1 <= 0# Or l2 <= 0# Then Exit Function
        q = q + (l1 + l2) / (2# * cross)
        angle = AcosValue((ux * vx + uy * vy) / Sqr(l1 * l2)) * 57.2957795130823
        If angle < amin Then amin = angle
        If angle > metricMax(e) Then metricMax(e) = angle
        If cross / Sqr(l1 * l2) < metricSJ(e) Then metricSJ(e) = cross / Sqr(l1 * l2)
        If angle < 30# - 0.0000001 Or angle > 150# + 0.0000001 Then
            metricClass(e) = 2
        ElseIf angle < 45# - 0.0000001 Or angle > 135# + 0.0000001 Then
            If metricClass(e) < 1 Then metricClass(e) = 1
        End If
        If l1 < shortest Then shortest = l1
        If l1 > longest Then longest = l1
        ar = ar + xx(a) * yy(b) - xx(b) * yy(a)
    Next j
    ar = ar / 2#: asp = Sqr(longest / shortest)
    Metric = ar > 0#
End Function

Private Sub CacheMetrics()
    Dim e As Long
    ReDim areas(1 To ne): ReDim energies(1 To ne): ReDim angles(1 To ne): ReDim aspects(1 To ne)
    ReDim metricMax(1 To ne): ReDim metricSJ(1 To ne): ReDim metricClass(1 To ne)
    geomVersion = geomVersion + 1: heatVersion = -1
    For e = 1 To ne
        If Not Metric(e, areas(e), energies(e), angles(e), aspects(e)) Then Err.Raise vbObjectError + 3803, , "反転・退化した四角形です。要素=" & CStr(e)
    Next e
End Sub

Private Function MaterialAreas() As Object
    Dim result As Object, e As Long, key As String
    Set result = CreateObject("Scripting.Dictionary")
    For e = 1 To ne
        key = CStr(materialId(e))
        If Not result.Exists(key) Then result(key) = 0#
        result(key) = CDbl(result(key)) + areas(e)
    Next e
    Set MaterialAreas = result
End Function

Private Function MinAngle() As Double
    Dim e As Long: MinAngle = 180#
    For e = 1 To ne: If angles(e) < MinAngle Then MinAngle = angles(e)
    Next e
End Function

Private Function MaxAspect() As Double
    Dim e As Long
    For e = 1 To ne: If aspects(e) > MaxAspect Then MaxAspect = aspects(e)
    Next e
End Function

Private Sub MakeLocks()
    Dim bn() As Long, ba() As Long, bb() As Long, k As Long, a As Long, b As Long, i As Long, j As Long, e As Long
    Dim ux As Double, uy As Double, vx As Double, vy As Double, mid As Long
    Dim loadRows As Variant, protectOriginal As Boolean, token As String
    ReDim locked(1 To nn): ReDim slideA(1 To nn): ReDim slideB(1 To nn)
    ReDim bn(1 To nn): ReDim ba(1 To nn): ReDim bb(1 To nn)
    loadRows = Table("載荷", 9)
    For i = 1 To UBound(loadRows, 1)
        token = UCase$(Trim$(CStr(loadRows(i, 2))))
        If UCase$(Trim$(CStr(loadRows(i, 4)))) = "FORCE" Or left$(token, 3) = "BOX" Or IsNumeric(token) Then protectOriginal = True
    Next i
    If protectOriginal Then
        For i = 1 To originalNN: locked(i) = True: Next i
    End If
    For k = 1 To edgeN
        a = ea(k): b = eb(k): mid = em(k)
        If protectOriginal And mid <= originalNN Then locked(a) = True: locked(b) = True
        If ee2(k) = 0 Then
            bn(a) = bn(a) + 1: bn(b) = bn(b) + 1
            If ba(a) = 0 Then ba(a) = b Else bb(a) = b
            If ba(b) = 0 Then ba(b) = a Else bb(b) = a
        ElseIf materialId(ee1(k)) <> materialId(ee2(k)) Then
            locked(a) = True: locked(b) = True
        End If
        If Abs(xx(mid) - (xx(a) + xx(b)) / 2#) + Abs(yy(mid) - (yy(a) + yy(b)) / 2#) > meshScale * 0.000000001 Then
            ' 曲線Q8は幾何を変えず保持。初版の平滑化は直線辺だけを対象にする。
            For j = 1 To 4: locked(conn(ee1(k), j)) = True: Next j
            If ee2(k) > 0 Then
                For j = 1 To 4: locked(conn(ee2(k), j)) = True: Next j
            End If
        End If
        For j = 6 To 14
            If Abs(d(nodeRows(mid, j))) > 0# Then locked(a) = True: locked(b) = True
        Next j
        If d(nodeRows(mid, 4)) <> 0# Or d(nodeRows(mid, 5)) <> 0# Then
            If ee2(k) > 0 Then locked(a) = True: locked(b) = True
        End If
    Next k
    For i = 1 To nn
        If incident(i).count = 0 Then locked(i) = True
        For j = 6 To 14
            If Abs(d(nodeRows(i, j))) > 0# Then locked(i) = True
        Next j
        If bn(i) = 0 Then
            If d(nodeRows(i, 4)) <> 0# Or d(nodeRows(i, 5)) <> 0# Then locked(i) = True
        ElseIf bn(i) <> 2 Then
            locked(i) = True
        Else
            a = ba(i): b = bb(i)
            ux = xx(a) - xx(i): uy = yy(a) - yy(i): vx = xx(b) - xx(i): vy = yy(b) - yy(i)
            If Abs(ux * vy - uy * vx) > meshScale * meshScale * 0.0000000001 Or ux * vx + uy * vy >= 0# Then
                locked(i) = True
            Else
                slideA(i) = a: slideB(i) = b
                For j = 4 To 5
                    If d(nodeRows(i, j)) <> d(nodeRows(a, j)) Or d(nodeRows(i, j)) <> d(nodeRows(b, j)) Then locked(i) = True
                Next j
            End If
        End If
    Next i
End Sub

Private Sub ShapeQ8(ByVal xi As Double, ByVal eta As Double, ByRef shape() As Double)
    shape(0) = -0.25 * (1# - xi) * (1# - eta) * (1# + xi + eta)
    shape(1) = -0.25 * (1# + xi) * (1# - eta) * (1# - xi + eta)
    shape(2) = -0.25 * (1# + xi) * (1# + eta) * (1# - xi - eta)
    shape(3) = -0.25 * (1# - xi) * (1# + eta) * (1# + xi - eta)
    shape(4) = 0.5 * (1# - xi * xi) * (1# - eta)
    shape(5) = 0.5 * (1# + xi) * (1# - eta * eta)
    shape(6) = 0.5 * (1# - xi * xi) * (1# + eta)
    shape(7) = 0.5 * (1# - xi) * (1# - eta * eta)
End Sub

Private Sub ParentPoint(ByVal e As Long, ByVal u As Double, ByVal v As Double, ByRef xp As Double, ByRef yp As Double)
    Dim shape(0 To 7) As Double, j As Long, n As Long
    ShapeQ8 u, v, shape: xp = 0#: yp = 0#
    For j = 0 To 7
        n = CLng(parentElements(e, j + 2))
        xp = xp + shape(j) * d(parentNodes(n, 2)): yp = yp + shape(j) * d(parentNodes(n, 3))
    Next j
End Sub

Private Function AddPoint(ByVal xp As Double, ByVal yp As Double) As Long
    nn = nn + 1
    If nn > UBound(xx) Then Err.Raise vbObjectError + 3804, , "節点作業配列の上限を超えました。"
    xx(nn) = xp: yy(nn) = yp: nodeRows(nn, 1) = nn
    AddPoint = nn
End Function

Private Sub InheritEdgeConstraint(ByVal newNode As Long, ByVal parent As Long, ByVal u As Double, ByVal v As Double)
    Dim edge As Long, a As Long, b As Long, m As Long, c As Long, t As Double
    edge = -1
    If Abs(v + 1#) < 0.0000000001 Then edge = 0: t = u
    If Abs(u - 1#) < 0.0000000001 Then edge = 1: t = v
    If Abs(v - 1#) < 0.0000000001 Then edge = 2: t = -u
    If Abs(u + 1#) < 0.0000000001 Then edge = 3: t = -v
    If edge < 0 Then Exit Sub
    a = CLng(parentElements(parent, edge + 2)): b = CLng(parentElements(parent, (edge + 1) Mod 4 + 2)): m = CLng(parentElements(parent, edge + 6))
    For c = 4 To 5
        If d(parentNodes(a, c)) = 1# And d(parentNodes(b, c)) = 1# And d(parentNodes(m, c)) = 1# Then
            nodeRows(newNode, c) = 1
            nodeRows(newNode, c + 2) = 0.5 * t * (t - 1#) * d(parentNodes(a, c + 2)) + (1# - t * t) * d(parentNodes(m, c + 2)) + 0.5 * t * (t + 1#) * d(parentNodes(b, c + 2))
        End If
    Next c
    ' 節点力は元の作用点に保持する。新節点に複製して合力を増やさない。
End Sub

Private Sub AddChild(ByVal parent As Long, ByVal vertices As Variant, ByVal us As Variant, ByVal vs As Variant, ByVal mids As Object, ByRef count As Long)
    Dim j As Long, a As Long, b As Long, m As Long, u As Double, v As Double, xp As Double, yp As Double, key As String
    count = count + 1: newElements(count, 1) = count: newElements(count, 10) = parentElements(parent, 10)
    candidateScore(count) = refineScore(parent)
    For j = 0 To 3
        a = CLng(vertices(j)): b = CLng(vertices((j + 1) Mod 4))
        newElements(count, j + 2) = a: key = EdgeKey(a, b)
        If mids.Exists(key) Then
            m = mids(key)
        Else
            u = (CDbl(us(j)) + CDbl(us((j + 1) Mod 4))) / 2#: v = (CDbl(vs(j)) + CDbl(vs((j + 1) Mod 4))) / 2#
            ParentPoint parent, u, v, xp, yp
            m = AddPoint(xp, yp): mids.Add key, m
            InheritEdgeConstraint m, parent, u, v
        End If
        newElements(count, j + 6) = m
    Next j
End Sub

Private Function CloseMarks(ByRef marked() As Boolean, ByVal budget As Long) As Long
    Dim changed As Boolean, e As Long, j As Long, n As Long, last As Long, total As Long
    Do
        changed = False
        For e = 1 To ne
            n = 0: last = 0
            For j = 0 To 3: If marked(elementEdge(e, j)) Then n = n + 1: last = j
            Next j
            If n = 1 Then
                marked(elementEdge(e, (last + 2) Mod 4)) = True: changed = True
            ElseIf n = 3 Then
                For j = 0 To 3: marked(elementEdge(e, j)) = True: Next j
                changed = True
            End If
        Next e
    Loop While changed
    For e = 1 To ne
        n = 0: last = -1
        For j = 0 To 3
            If marked(elementEdge(e, j)) Then n = n + 1: If last < 0 Then last = j
        Next j
        Select Case n
            Case 0: total = total + 1
            Case 2
                If marked(elementEdge(e, (last + 2) Mod 4)) Then total = total + 2 Else total = total + 3
            Case 4: total = total + 4
            Case Else: Err.Raise vbObjectError + 3805, , "移行要素の辺分割が不整合です。"
        End Select
    Next e
    CloseMarks = total
End Function

Private Function FinalMaterialStates() As Object
    Dim states As Object, rows As Variant, r As Long, token As String, v As Variant, enabled As Boolean
    Set states = CreateObject("Scripting.Dictionary")
    rows = Table("材料データ", 10)
    For r = 1 To UBound(rows, 1)
        If Not IsEmpty(rows(r, 1)) Then
            v = rows(r, 10): enabled = True
            If Not IsEmpty(v) Then
                If IsNumeric(v) Then
                    enabled = Abs(CDbl(v)) >= 0.5
                Else
                    token = UCase$(Trim$(CStr(v)))
                    If token = "OFF" Or token = "NO" Or token = "N" Or token = "FALSE" Or token = "無効" Or token = "しない" Then enabled = False
                End If
            End If
            states(CLng(rows(r, 1))) = enabled
        End If
    Next r
    rows = Table("ステージ", 5)
    For r = 1 To UBound(rows, 1)
        If IsNumeric(rows(r, 5)) Then
            If CDbl(rows(r, 5)) <> 0# Then
                token = UCase$(Trim$(CStr(rows(r, 2))))
                Select Case token
                    Case "BIRTH", "誕生", "盛土": states(CLng(rows(r, 3))) = True
                    Case "DEATH", "死滅", "掘削", "KILL": states(CLng(rows(r, 3))) = False
                End Select
            End If
        End If
    Next r
    Set FinalMaterialStates = states
End Function

Private Sub RefinePlastic(Optional ByVal transferred As Boolean = False, Optional ByVal absoluteBudget As Long = 0)
    Dim ws As Worksheet, p2 As Variant, scores() As Double, originalScores() As Double, ranked() As Long, marks() As Boolean, trialMarks() As Boolean
    Dim e As Long, r As Long, j As Long, k As Long, id As Long, last As Long, gp As Long, countGP() As Long, gpSeen As Object
    Dim q As Double, delta As Double, maxQ As Double, best As Double, budget As Long, targetCount As Long, proposed As Long
    Dim n As Long, start As Long, count As Long, a(0 To 3) As Long, m(0 To 3) As Long, us(0 To 3) As Double, vs(0 To 3) As Double
    Dim mids As Object, center As Long, xp As Double, yp As Double, c0 As Long, c1 As Long, c2 As Long, c3 As Long, upper As Double
    Dim dirty As String, nodeResults As Variant, finalStates As Object
    Call CaptureParents
    EnsureRefinementCapacity
    If transferred Then
        scores = refineScore
        ReDim ranked(1 To ne)
        For e = 1 To ne: ranked(e) = e: Next e
        GoTo ScoresReady
    End If
    Set ws = ThisWorkbook.Worksheets("診断")
    If CStr(ws.Cells(2, 2).value2) <> "PASS" Then Err.Raise vbObjectError + 3806, , "正常終了した計算結果が必要です。"
    On Error Resume Next: dirty = ThisWorkbook.names("_FEMViewerResultsDirty").RefersTo: On Error GoTo 0
    If dirty = "=TRUE" Then Err.Raise vbObjectError + 3806, , "入力変更後です。再解析してください。"
    If SavedSignature() <> MeshInputSignature() Then Err.Raise vbObjectError + 3806, , "入力と保存結果の署名が一致しません。再解析してください。"
    If CStr(ws.Cells(106, 40).value2) <> "Element" Or CStr(ws.Cells(106, 48).value2) <> "Yielded" Then Err.Raise vbObjectError + 3806, , "塑性積分点結果がありません。再解析してください。"
    p2 = ws.range(ws.Cells(107, 40), ws.Cells(106 + ne * 4, 53)).value2
    operationStep = "塑性結果の読込"
    ReDim scores(1 To ne): ReDim ranked(1 To ne): ReDim countGP(1 To ne)
    Set gpSeen = CreateObject("Scripting.Dictionary")
    Set finalStates = FinalMaterialStates()
    For r = 1 To UBound(p2, 1)
        id = CLng(d(p2(r, 1))): gp = CLng(d(p2(r, 2)))
        If id < 1 Or id > ne Or gp < 0 Or gp > 3 Then Err.Raise vbObjectError + 3806, , "塑性結果の要素番号が不正です。"
        If gpSeen.Exists(CStr(id) & ":" & CStr(gp)) Then Err.Raise vbObjectError + 3806, , "塑性結果の積分点が重複しています。"
        gpSeen.Add CStr(id) & ":" & CStr(gp), True: countGP(id) = countGP(id) + 1
        If Not finalStates.Exists(materialId(id)) Then Err.Raise vbObjectError + 3806, , "材料の初期状態が不明です。"
        If CBool(p2(r, 9)) And CBool(finalStates(materialId(id))) Then
            q = Sqr(d(p2(r, 3)) ^ 2 + d(p2(r, 4)) ^ 2 + d(p2(r, 5)) ^ 2 + 0.5 * d(p2(r, 6)) ^ 2)
            If q > scores(id) Then scores(id) = q
        End If
    Next r
    For e = 1 To ne
        If countGP(e) <> 4 Then Err.Raise vbObjectError + 3806, , "要素あたり4積分点の結果が必要です。"
        If scores(e) > maxQ Then maxQ = scores(e)
    Next e
    If maxQ <= 1E-18 Then Err.Raise vbObjectError + 3806, , "細分化対象となる塑性更新がありません。"
    For e = 1 To ne: scores(e) = scores(e) / maxQ: ranked(e) = e: Next e
    originalScores = scores
    For k = 1 To edgeN
        If ee2(k) > 0 Then
            If materialId(ee1(k)) = materialId(ee2(k)) Then
                delta = Abs(originalScores(ee1(k)) - originalScores(ee2(k))) * 0.1
                scores(ee1(k)) = scores(ee1(k)) + delta: scores(ee2(k)) = scores(ee2(k)) + delta
            End If
        End If
    Next k
    For e = 1 To ne
        If scores(e) > 0# Then scores(e) = scores(e) + 0.05 * (energies(e) / 4# - 1#)
    Next e
ScoresReady:
    refineScore = scores
    RankScores scores, ranked, ne
    upper = FEMReadSetting("ADAPT_MAX_ELEMENT_RATIO", 2#)
    If upper < 1.05 Or upper > 4# Then Err.Raise vbObjectError + 3806, , "最大要素数倍率は1.05～4で指定してください。"
    budget = CLng(Fix(ne * upper))
    If absoluteBudget > 0 Then budget = absoluteBudget
    ReDim marks(1 To edgeN)
    operationStep = "分割辺の選定"
    targetCount = CLng(ne * FEMReadSetting("ADAPT_MARK_FRACTION", 0.15))
    If targetCount < 1 Then targetCount = 1
    If targetCount > ne Then targetCount = ne
    count = 0
    For r = 1 To ne
        e = ranked(r): If scores(e) <= 0# Then Exit For
        n = 0: For j = 0 To 3: If marks(elementEdge(e, j)) Then n = n + 1
        Next j
        If n = 4 Then GoTo NextRank
        If TryMarkElement(e, marks, budget) Then count = count + 1
        If count >= targetCount Then Exit For
NextRank:
    Next r
    If count = 0 Then Err.Raise vbObjectError + 3806, , "適合性を保った細分化が要素数上限に入りません。倍率を増やしてください。"
    ReDim newElements(1 To ne * 4, 1 To 10)
    ReDim candidateScore(1 To ne * 4)
    Set mids = CreateObject("Scripting.Dictionary")
    For k = 1 To edgeN: If Not marks(k) Then mids(EdgeKey(ea(k), eb(k))) = em(k)
    Next k
    us(0) = -1#: us(1) = 1#: us(2) = 1#: us(3) = -1#
    vs(0) = -1#: vs(1) = -1#: vs(2) = 1#: vs(3) = 1#
    count = 0
    For e = 1 To ne
        operationStep = "Q8生成 要素=" & CStr(e)
        n = 0: start = -1
        For j = 0 To 3
            a(j) = conn(e, j + 1): m(j) = conn(e, j + 5)
            If marks(elementEdge(e, j)) Then n = n + 1: If start < 0 Then start = j
        Next j
        If n = 0 Then
            AddChild e, Array(a(0), a(1), a(2), a(3)), Array(-1#, 1#, 1#, -1#), Array(-1#, -1#, 1#, 1#), mids, count
        ElseIf n = 4 Then
            ParentPoint e, 0#, 0#, xp, yp: center = AddPoint(xp, yp)
            For j = 0 To 3
                c0 = j: c1 = (j + 1) Mod 4: c3 = (j + 3) Mod 4
                AddChild e, Array(a(c0), m(c0), center, m(c3)), Array(us(c0), (us(c0) + us(c1)) / 2#, 0#, (us(c3) + us(c0)) / 2#), Array(vs(c0), (vs(c0) + vs(c1)) / 2#, 0#, (vs(c3) + vs(c0)) / 2#), mids, count
            Next j
        ElseIf n = 2 Then
            If marks(elementEdge(e, (start + 2) Mod 4)) Then
                c0 = start: c1 = (start + 1) Mod 4: c2 = (start + 2) Mod 4: c3 = (start + 3) Mod 4
                AddChild e, Array(a(c0), m(c0), m(c2), a(c3)), Array(us(c0), (us(c0) + us(c1)) / 2#, (us(c2) + us(c3)) / 2#, us(c3)), Array(vs(c0), (vs(c0) + vs(c1)) / 2#, (vs(c2) + vs(c3)) / 2#, vs(c3)), mids, count
                AddChild e, Array(m(c0), a(c1), a(c2), m(c2)), Array((us(c0) + us(c1)) / 2#, us(c1), us(c2), (us(c2) + us(c3)) / 2#), Array((vs(c0) + vs(c1)) / 2#, vs(c1), vs(c2), (vs(c2) + vs(c3)) / 2#), mids, count
            Else
                For j = 0 To 3
                    If marks(elementEdge(e, j)) And marks(elementEdge(e, (j + 1) Mod 4)) Then start = j: Exit For
                Next j
                c0 = start: c1 = (start + 1) Mod 4: c2 = (start + 2) Mod 4: c3 = (start + 3) Mod 4
                ParentPoint e, 0#, 0#, xp, yp: center = AddPoint(xp, yp)
                AddChild e, Array(center, m(c0), a(c1), m(c1)), Array(0#, (us(c0) + us(c1)) / 2#, us(c1), (us(c1) + us(c2)) / 2#), Array(0#, (vs(c0) + vs(c1)) / 2#, vs(c1), (vs(c1) + vs(c2)) / 2#), mids, count
                AddChild e, Array(center, m(c1), a(c2), a(c3)), Array(0#, (us(c1) + us(c2)) / 2#, us(c2), us(c3)), Array(0#, (vs(c1) + vs(c2)) / 2#, vs(c2), vs(c3)), mids, count
                AddChild e, Array(center, a(c3), a(c0), m(c0)), Array(0#, us(c3), us(c0), (us(c0) + us(c1)) / 2#), Array(0#, vs(c3), vs(c0), (vs(c0) + vs(c1)) / 2#), mids, count
            End If
        End If
    Next e
    ne = count
    operationStep = "生成後の接続構築"
    For e = 1 To ne
        materialId(e) = CLng(newElements(e, 10))
        For j = 1 To 8: conn(e, j) = CLng(newElements(e, j + 1)): Next j
    Next e
    Call BuildEdges
    Call CacheMetrics
End Sub

Private Sub CheckCandidate()
    Dim e As Long, i As Long, j As Long, k As Long, key As Variant, sums As Object
    Dim shapeP(0 To 7) As Double, shapeM(0 To 7) As Double, dxu As Double, dyu As Double, dxv As Double, dyv As Double
    Dim u As Double, v As Double, h As Double, det As Double, used() As Boolean
    Set sums = MaterialAreas()
    If nn > 60000 Or ne > 60000 Then Err.Raise vbObjectError + 3807, , "解析の節点・要素上限60000を超えます。"
    estimatedBand = CLng(8# * Sqr(CDbl(nn)))
    If estimatedBand > 2 * nn - 1 Then estimatedBand = 2 * nn - 1
    estimatedMemory = 4# * 8# * (2# * nn) * (estimatedBand + 1#) + ne * 16384#
    If estimatedMemory > FEMReadSetting("SOLVER_MEMORY_LIMIT_MB", 512#) * 1048576# Then Err.Raise vbObjectError + 3807, , "概算の作業メモリが上限を超えます。要素数倍率を下げてください。"
    For Each key In oldAreaByMaterial.Keys
        If Not sums.Exists(key) Then Err.Raise vbObjectError + 3807, , "材料領域が失われました。"
        If Abs(sums(key) - oldAreaByMaterial(key)) > 0.00000001 * (1# + Abs(oldAreaByMaterial(key))) Then Err.Raise vbObjectError + 3807, , "材料ごとの面積が変わりました。"
    Next key
    ReDim used(1 To nn): h = 0.00001
    EnsureJacobianCoefficients
    For e = 1 To ne
        For j = 1 To 8: used(conn(e, j)) = True: Next j
        For i = 0 To 4
            u = -1# + i * 0.5
            For j = 0 To 4
                v = -1# + j * 0.5
                ' Shape differences are invariant across elements.
                dxu = 0#: dyu = 0#: dxv = 0#: dyv = 0#
                For k = 0 To 7
                    dxu = dxu + jacDu(i * 5 + j, k) * xx(conn(e, k + 1)) / (2# * h)
                    dyu = dyu + jacDu(i * 5 + j, k) * yy(conn(e, k + 1)) / (2# * h)
                Next k
                ' Same 25 points and finite-difference step as before.
                For k = 0 To 7
                    dxv = dxv + jacDv(i * 5 + j, k) * xx(conn(e, k + 1)) / (2# * h)
                    dyv = dyv + jacDv(i * 5 + j, k) * yy(conn(e, k + 1)) / (2# * h)
                Next k
                det = dxu * dyv - dxv * dyu
                If det <= meshScale * meshScale * 0.0000000000001 Then Err.Raise vbObjectError + 3807, , "Q8のJacobian検査で候補を棄却しました。要素=" & CStr(e)
            Next j
        Next i
    Next e
    For i = 1 To nn: If Not used(i) Then Err.Raise vbObjectError + 3807, , "候補に孤立節点があります。"
    Next i
    If candidateRebuilt Then
        Call ValidateTerrainConditions
    Else
        For i = 1 To originalNN
            For j = 4 To 14
                If CStr(nodeRows(i, j)) <> CStr(oldNodes(i, j)) Then Err.Raise vbObjectError + 3807, , "既存の節点条件が変化しました。"
            Next j
        Next i
    End If
End Sub

Private Function CandidateOptionSignature() As String
    CandidateOptionSignature = CStr(FEMReadSetting("ADAPT_MAX_ELEMENT_RATIO", 2#)) & "|" & _
        CStr(FEMReadSetting("ADAPT_MARK_FRACTION", 0.15)) & "|" & _
        CStr(FEMReadSetting("MESH_REPAIR_MAX_ITERATIONS", 20#)) & "|" & _
        CStr(FEMReadSetting("MESH_REPAIR_ANGLE_ITERATIONS", 2000#)) & "|" & CStr(FEMReadSetting("MESH_REPAIR_TERRAIN_COMPARE", 1#)) & "|" & _
        CStr(FEMReadSetting("MESH_GEN_COMPARE", 1#)) & "|" & CStr(FEMReadSetting("MESH_TERRAIN_COUNT_MIN", 0.55)) & "|" & CStr(FEMReadSetting("MESH_TERRAIN_COUNT_MAX", 1.8))
End Function

Public Function MeshPrepareCandidate(ByVal kind As String) As Boolean
    Dim oldStatus As Variant
    If busy Or AnalysisRunning Then MeshAdaptLastError = "解析またはメッシュ処理中です。": Exit Function
    busy = True: candidateReady = False: MeshAdaptLastError = "": oldStatus = Application.StatusBar
    On Error GoTo Failed
    Application.StatusBar = "メッシュ候補：入力と境界条件を確認"
    candidateOptions = CandidateOptionSignature()
    candidateSignature = MeshInputSignature(): candidateKind = UCase$(kind)
    operationStep = "入力読込"
    LoadMesh
    If candidateKind <> "PLASTIC" And candidateKind <> "SMOOTH" Then Err.Raise vbObjectError + 3808, , "不明なメッシュ操作です。"
    If Not GenerateAndCompare() Then Err.Raise vbObjectError + 3808, , "有効な候補がありません。"
    candidateReady = True: MeshPrepareCandidate = True
    GoTo Finish
Failed:
    MeshAdaptLastError = operationStep & ": " & Err.Description
Finish:
    Application.StatusBar = oldStatus: busy = False
End Function

Public Function MeshCandidateError() As String
    MeshCandidateError = MeshAdaptLastError
End Function

Public Function MeshAdaptSelectedMethod() As String
    MeshAdaptSelectedMethod = selectedMethod
End Function

Public Sub MeshCancelCandidate()
    candidateReady = False
    On Error Resume Next
    ThisWorkbook.Worksheets("再メッシュ候補").Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

Public Function MeshApplyCandidate(Optional ByVal skipBackup As Boolean = False) As Boolean
    Dim wsN As Worksheet, wsE As Worksheet, outN() As Variant, outE() As Variant
    Dim i As Long, j As Long, oldEvents As Boolean, oldScreen As Boolean, changed As Boolean, folder As String
    Dim settingsSaved As Variant, metaSaved As Variant, metaSheet As Worksheet, hadMeta As Boolean
    Dim wasOK As Boolean, wasP1 As Boolean, oldRevision As Long, dirtySaved As String, hadDirty As Boolean
    Dim oldAlerts As Boolean, settingsLast As Long
    If Not candidateReady Or busy Or AnalysisRunning Then MeshAdaptLastError = "有効な候補がありません。": Exit Function
    busy = True: oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    If candidateSignature <> MeshInputSignature() Or candidateOptions <> CandidateOptionSignature() Then Err.Raise vbObjectError + 3809, , "候補作成後に入力が変わりました。候補を作り直してください。"
    settingsSaved = Table("設定", 5)
    wasOK = AnalysisOK: wasP1 = P1ResultReady: oldRevision = ResultRevision
    On Error Resume Next
    dirtySaved = ThisWorkbook.names("_FEMViewerResultsDirty").RefersTo: hadDirty = (Err.Number = 0): Err.Clear
    Set metaSheet = ThisWorkbook.Worksheets("メッシュ引継")
    On Error GoTo Failed
    hadMeta = Not metaSheet Is Nothing
    If hadMeta Then metaSaved = metaSheet.UsedRange.value2
    MeshAdaptLastBackup = ""
    If Not skipBackup Then
        folder = Environ$("USERPROFILE") & "\Documents\WORK\03_バックアップ"
        If Len(Dir$(folder, vbDirectory)) > 0 Then
            folder = folder & Application.PathSeparator & "2DSoilFEM_適応メッシュ"
        Else
            folder = ThisWorkbook.path & Application.PathSeparator & "MeshBackups"
        End If
        If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
        MeshAdaptLastBackup = folder & Application.PathSeparator & "before_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 100#)) & ".xlsm"
        ThisWorkbook.SaveCopyAs MeshAdaptLastBackup
    End If
    Application.EnableEvents = False: Application.ScreenUpdating = False
    ReDim outN(1 To nn, 1 To 14): ReDim outE(1 To ne, 1 To 10)
    For i = 1 To nn
        For j = 1 To 14: outN(i, j) = nodeRows(i, j): Next j
        outN(i, 1) = i: outN(i, 2) = xx(i): outN(i, 3) = yy(i)
    Next i
    For i = 1 To ne
        outE(i, 1) = i: outE(i, 10) = materialId(i)
        For j = 1 To 8: outE(i, j + 1) = conn(i, j): Next j
    Next i
    Set wsN = ThisWorkbook.Worksheets("節点データ"): Set wsE = ThisWorkbook.Worksheets("要素データ")
    changed = True
    If nn < originalNN Then wsN.range(wsN.Cells(nn + 2, 1), wsN.Cells(originalNN + 1, 14)).ClearContents
    If ne < originalNE Then wsE.range(wsE.Cells(ne + 2, 1), wsE.Cells(originalNE + 1, 10)).ClearContents
    wsN.range(wsN.Cells(2, 1), wsN.Cells(nn + 1, 14)).value2 = outN
    wsE.range(wsE.Cells(2, 1), wsE.Cells(ne + 1, 10)).value2 = outE
    Call FEMInvalidateRemeshedModel
    Call FEMInvalidateSettingCache
    P5EvaluateCurrentSheet "全体最適化・適応Q8"
    FEMWriteNumericSetting "MESH_QUALITY_LAST_REPAIR_ITERATIONS", sweepCount
    FEMWriteNumericSetting "MESH_QUALITY_LAST_MOVED_NODES", movedCount
    AnalysisOK = False: P1ResultReady = False: ResultRevision = 0
    FEMViewerMarkDirty
    SaveForceAnchors
    candidateReady = False: MeshApplyCandidate = True
    GoTo Finish
Failed:
    MeshAdaptLastError = Err.Description
    Resume Rollback
Rollback:
    If changed Then
        On Error Resume Next
        wsN.range(wsN.Cells(2, 1), wsN.Cells(nn + 1, 14)).ClearContents
        wsE.range(wsE.Cells(2, 1), wsE.Cells(ne + 1, 10)).ClearContents
        wsN.range(wsN.Cells(2, 1), wsN.Cells(originalNN + 1, 14)).value2 = oldNodes
        wsE.range(wsE.Cells(2, 1), wsE.Cells(originalNE + 1, 10)).value2 = oldElements
        With ThisWorkbook.Worksheets("設定")
            settingsLast = .Cells(.rows.count, 5).End(xlUp).row
            If settingsLast >= 2 Then .range(.Cells(2, 1), .Cells(settingsLast, 5)).ClearContents
            .range(.Cells(2, 1), .Cells(UBound(settingsSaved, 1) + 1, 5)).value2 = settingsSaved
        End With
        Set metaSheet = Nothing: Set metaSheet = ThisWorkbook.Worksheets("メッシュ引継")
        If Not metaSheet Is Nothing Then
            If hadMeta Then
                metaSheet.Cells.ClearContents
                metaSheet.range(metaSheet.Cells(1, 1), metaSheet.Cells(UBound(metaSaved, 1), UBound(metaSaved, 2))).value2 = metaSaved
            Else
                oldAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False: metaSheet.Delete: Application.DisplayAlerts = oldAlerts
            End If
        End If
        AnalysisOK = wasOK: P1ResultReady = wasP1: ResultRevision = oldRevision
        If hadDirty Then ThisWorkbook.names("_FEMViewerResultsDirty").RefersTo = dirtySaved
        FEMInvalidateInputCache
        FEMInvalidateSettingCache
        On Error GoTo 0
    End If
Finish:
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen: busy = False
    If MeshApplyCandidate Then
        MeshCancelCandidate
        FEMViewAll
    End If
End Function

Private Sub PreviewCandidate()
    Dim ws As Worksheet, part() As String, text As String, e As Long, j As Long, n As Long, order As Variant
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double, drawScale As Double, points As String
    Dim temp As String, stream As Object, shape As Object, key As String, i As Long, shade As Long, colorText As String
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("再メッシュ候補"): On Error GoTo 0
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = "再メッシュ候補"
    ws.Visible = xlSheetVisible
    For i = ws.Shapes.count To 1 Step -1: ws.Shapes(i).Delete: Next i
    ws.Cells.ClearContents
    ws.range("A1:P1").Merge
    ws.range("A1:P1").Interior.Color = RGB(31, 70, 105)
    ws.range("A1:P1").Font.Color = vbWhite
    ws.range("A1:P1").Font.bOld = True
    ws.Cells(1, 1).value2 = "候補です。元のメッシュは採用するまで変更されません。"
    ws.range("A2:P10").Merge: ws.range("A2:P10").WrapText = True
    ws.range("A2:P10").Font.size = 10
    ws.range("A2:P10").value2 = MeshAdaptLastSummary
    ws.rows("1:10").RowHeight = 20
    ws.columns("A:P").ColumnWidth = 9
    minX = xx(1): maxX = xx(1): minY = yy(1): maxY = yy(1)
    For i = 2 To nn
        If xx(i) < minX Then minX = xx(i)
        If xx(i) > maxX Then maxX = xx(i)
        If yy(i) < minY Then minY = yy(i)
        If yy(i) > maxY Then maxY = yy(i)
    Next i
    drawScale = 760# / (maxX - minX + 0.000000000001)
    If drawScale > 370# / (maxY - minY + 0.000000000001) Then drawScale = 370# / (maxY - minY + 0.000000000001)
    ReDim part(0 To ne + 1): order = Array(1, 5, 2, 6, 3, 7, 4, 8)
    part(0) = "<svg xmlns=""http://www.w3.org/2000/svg"" width=""820"" height=""430""><rect width=""820"" height=""430"" fill=""white""/>"
    For e = 1 To ne
        points = ""
        For j = 0 To 7
            n = conn(e, order(j))
            points = points & Replace$(CStr(30# + (xx(n) - minX) * drawScale), Application.International(xlDecimalSeparator), ".") & "," & Replace$(CStr(30# + (maxY - yy(n)) * drawScale), Application.International(xlDecimalSeparator), ".") & " "
        Next j
        Select Case ElementClass(e)
            Case 0: colorText = "#b7e2c1"
            Case 1: colorText = "#ffd578"
            Case 2: colorText = "#ef8787"
        End Select
        If candidateKind = "PLASTIC" Then
            shade = CLng(245# - 150# * candidateScore(e))
            If shade < 80 Then shade = 80
            If shade > 245 Then shade = 245
            colorText = "#FF" & right$("0" & Hex$(shade), 2) & right$("0" & Hex$(shade), 2)
        End If
        part(e) = "<polygon points=""" & points & """ fill=""" & colorText & """ stroke=""#334155"" stroke-width=""0.5""/>"
    Next e
    part(ne + 1) = "<text x=""30"" y=""424"" font-family=""Meiryo"" font-size=""11"">緑：合格45～135° / 黄：警告30～150°内 / 赤：不合格</text></svg>"
    If candidateKind = "PLASTIC" Then part(ne + 1) = "<text x=""30"" y=""424"" font-family=""Meiryo"" font-size=""11"">赤ほど直近の塑性更新指標が大きい（累積塑性ひずみではありません）</text></svg>"
    text = Join(part, "")
    temp = ThisWorkbook.path & Application.PathSeparator & "~mesh_candidate.svg"
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.WriteText text: stream.SaveToFile temp, 2: stream.Close
    Set shape = ws.Shapes.AddPicture(temp, msoFalse, msoTrue, 15, 270, 820, 430): Kill temp
    ws.Activate
    Set shape = ws.Shapes.AddFormControl(xlButtonControl, 15, 220, 140, 28)
    ws.Buttons(shape.name).caption = "候補を採用"
    shape.OnAction = "'" & Replace$(ThisWorkbook.name, "'", "''") & "'!MeshApplyCandidateUI"
    Set shape = ws.Shapes.AddFormControl(xlButtonControl, 175, 220, 100, 28)
    ws.Buttons(shape.name).caption = "取消"
    shape.OnAction = "'" & Replace$(ThisWorkbook.name, "'", "''") & "'!MeshCancelCandidateUI"
    ws.Activate: ActiveWindow.Zoom = 90: ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
End Sub

Public Sub MeshPlasticRemesh()
    If MeshPrepareCandidate("PLASTIC") Then
        PreviewCandidate
    ElseIf Not SuppressUserMessages Then
        MsgBox MeshAdaptLastError, vbExclamation
    End If
End Sub

Public Sub MeshOptimizeExisting()
    If MeshPrepareCandidate("SMOOTH") Then
        PreviewCandidate
    ElseIf Not SuppressUserMessages Then
        MsgBox MeshAdaptLastError, vbExclamation
    End If
End Sub

Public Sub MeshApplyCandidateUI()
    If MeshApplyCandidate() Then
        If Not SuppressUserMessages Then MsgBox "候補を採用しました。計算実行で初期状態から再解析してください。" & vbCrLf & "元のブック：" & MeshAdaptLastBackup, vbInformation
    ElseIf Not SuppressUserMessages Then
        MsgBox MeshAdaptLastError, vbExclamation
    End If
End Sub

Public Sub MeshCancelCandidateUI()
    ThisWorkbook.Worksheets("設定").Activate
    MeshCancelCandidate
End Sub




Private Sub CaptureState(ByRef s As CandidateState)
    Dim i As Long, j As Long
    s.countN = nn: s.countE = ne
    ReDim s.x(1 To nn): ReDim s.y(1 To nn): ReDim s.nodes(1 To nn, 1 To 14)
    ReDim s.elements(1 To ne, 1 To 8): ReDim s.materials(1 To ne): ReDim s.score(1 To ne)
    For i = 1 To nn
        s.x(i) = xx(i): s.y(i) = yy(i)
        For j = 1 To 14: s.nodes(i, j) = nodeRows(i, j): Next j
    Next i
    For i = 1 To ne
        For j = 1 To 8: s.elements(i, j) = conn(i, j): Next j
        s.materials(i) = materialId(i): s.score(i) = candidateScore(i)
    Next i
    s.rebuilt = candidateRebuilt: s.steps = sweepCount
End Sub
Private Sub RestoreState(ByRef s As CandidateState)
    nn = s.countN: ne = s.countE: xx = s.x: yy = s.y: nodeRows = s.nodes: conn = s.elements: materialId = s.materials
    candidateScore = s.score: candidateRebuilt = s.rebuilt: sweepCount = s.steps
    Call BuildEdges
    Call CacheMetrics
End Sub
Private Sub CaptureParents()
    Dim i As Long, j As Long
    ReDim parentNodes(1 To nn, 1 To 14): ReDim parentElements(1 To ne, 1 To 10)
    For i = 1 To nn
        For j = 1 To 14: parentNodes(i, j) = nodeRows(i, j): Next j
        parentNodes(i, 2) = xx(i): parentNodes(i, 3) = yy(i)
    Next i
    For i = 1 To ne
        For j = 1 To 8: parentElements(i, j + 1) = conn(i, j): Next j
        parentElements(i, 10) = materialId(i)
    Next i
End Sub
Private Function MaxAngle() As Double
    Dim e As Long
    For e = 1 To ne: If metricMax(e) > MaxAngle Then MaxAngle = metricMax(e)
    Next e
End Function
Private Function ElementClass(ByVal e As Long) As Long
    ElementClass = metricClass(e)
End Function
Private Sub GetQuality(ByRef q As QualityRecord, Optional ByVal withHeat As Boolean = True)
    Dim e As Long, k As Long, cls As Long, ratio As Double, j As Long, a As Long, b As Long, c As Long
    Dim ux As Double, uy As Double, vx As Double, vy As Double, sj As Double
    q.passN = 0: q.warnN = 0: q.failN = 0: q.scaledJ = 1#: q.grading = 1#: q.heatArea = 0#
    q.countN = nn: q.countE = ne: q.minA = MinAngle(): q.maxA = MaxAngle(): q.aspect = MaxAspect()
    For e = 1 To ne
        cls = ElementClass(e)
        Select Case cls
            Case 0: q.passN = q.passN + 1
            Case 1: q.warnN = q.warnN + 1
            Case 2: q.failN = q.failN + 1
        End Select
        If metricSJ(e) < q.scaledJ Then q.scaledJ = metricSJ(e)
    Next e
    For k = 1 To edgeN
        If ee2(k) > 0 Then
            ratio = areas(ee1(k)) / areas(ee2(k)): If ratio < 1# Then ratio = 1# / ratio
            If ratio > q.grading Then q.grading = ratio
        End If
    Next k
    If withHeat And candidateKind = "PLASTIC" Then q.heatArea = PlasticResolution(False)
End Sub
Private Function BetterQuality(ByRef a As QualityRecord, ByRef b As QualityRecord) As Boolean
    Dim am As Double, bm As Double
    If a.failN <> b.failN Then BetterQuality = a.failN < b.failN: Exit Function
    If a.warnN <> b.warnN Then BetterQuality = a.warnN < b.warnN: Exit Function
    am = a.minA - 45#: If 135# - a.maxA < am Then am = 135# - a.maxA
    bm = b.minA - 45#: If 135# - b.maxA < bm Then bm = 135# - b.maxA
    If Abs(am - bm) > 0.0001 Then BetterQuality = am > bm: Exit Function
    If Abs(a.scaledJ - b.scaledJ) > 0.0001 Then BetterQuality = a.scaledJ > b.scaledJ: Exit Function
    If Abs(a.aspect - b.aspect) > 0.001 Then BetterQuality = a.aspect < b.aspect: Exit Function
    If Abs(a.grading - b.grading) > 0.001 Then BetterQuality = a.grading < b.grading: Exit Function
    BetterQuality = a.countN < b.countN
End Function

Private Function QualityPasses(ByRef q As QualityRecord) As Boolean
    QualityPasses = (q.failN = 0)
End Function

' MESH_GEN_COMPARE when both pass (failN=0):
'   0 = pick the better by BetterQuality
'   1 = keep Delaunay/angle-optimized
'   2 = take the terrain candidate
' Two terrain candidates are always ranked by BetterQuality.
Private Function PreferTerrainOver(ByRef terrainQ As QualityRecord, ByRef bestQ As QualityRecord, ByVal bestAlreadyTerrain As Boolean) As Boolean
    Dim policy As Double
    policy = FEMReadSetting("MESH_GEN_COMPARE", 1#)
    ' C64=2: keep a passing raw terrain grid; angle-smoothing interiors makes it bumpy.
    If bestAlreadyTerrain Then
        If policy >= 1.5 And QualityPasses(bestQ) Then
            PreferTerrainOver = (terrainQ.failN < bestQ.failN)
            Exit Function
        End If
        PreferTerrainOver = BetterQuality(terrainQ, bestQ)
        Exit Function
    End If
    If QualityPasses(terrainQ) And QualityPasses(bestQ) Then
        If policy >= 1.5 Then
            PreferTerrainOver = True
            Exit Function
        End If
        If policy >= 0.5 Then
            PreferTerrainOver = False
            Exit Function
        End If
    End If
    PreferTerrainOver = BetterQuality(terrainQ, bestQ)
End Function
Private Function QualityLine(ByVal title As String, ByRef q As QualityRecord) As String
    QualityLine = title & "：" & q.countE & "要素、合格" & q.passN & " / 警告" & q.warnN & " / 不合格" & q.failN & _
        "、角度 " & Format$(q.minA, "0.000") & "～" & Format$(q.maxA, "0.000") & "°"
End Function
Private Sub OptimizeAngles()
    Dim k As Long, a As Long, b As Long, limit As Long, it As Long, bx() As Double, by() As Double, oldSteps As Long, q0 As QualityRecord, q1 As QualityRecord
    bx = xx: by = yy: oldSteps = sweepCount
    Call GetQuality(q0)
    On Error GoTo Reject
    Call MakeLocks
    ' Preserve every boundary vertex, including holes, and material interfaces.
    For k = 1 To edgeN
        If ee2(k) = 0 Then locked(ea(k)) = True: locked(eb(k)) = True
    Next k
    limit = CLng(FEMReadSetting("MESH_REPAIR_ANGLE_ITERATIONS", 2000#))
    If limit < 20 Then limit = 20
    If limit > 5000 Then limit = 5000
    Call MeshAngleOptimize(xx, yy, conn, nn, ne, locked, limit, it)
    sweepCount = it
    Call CorrectRightCorners
    For k = 1 To edgeN
        a = ea(k): b = eb(k)
        If Not locked(a) Or Not locked(b) Then xx(em(k)) = (xx(a) + xx(b)) / 2#: yy(em(k)) = (yy(a) + yy(b)) / 2#
    Next k
    On Error GoTo Reject
    Call CacheMetrics
    Call CheckCandidate
    Call GetQuality(q1)
    If BetterQuality(q0, q1) Then GoTo Reject
    If candidateKind = "PLASTIC" And plasticHeatLimit > 0# Then
        If q1.heatArea > plasticHeatLimit Then GoTo Reject
    End If
    Exit Sub
Reject:
    xx = bx: yy = by: sweepCount = oldSteps
    Call CacheMetrics
End Sub
Private Sub CorrectRightCorners()
    Dim bn() As Long, ba() As Long, bb() As Long, ni() As Long, k As Long, a As Long, b As Long, n As Long, p As Long
    Dim ux As Double, uy As Double, vx As Double, vy As Double, ln As Double, t As Double, ox As Double, oy As Double
    Dim item As Variant, ar As Double, energy As Double, angle As Double, aspect As Double, valid As Boolean
    ReDim bn(1 To nn): ReDim ba(1 To nn): ReDim bb(1 To nn): ReDim ni(1 To nn)
    For k = 1 To edgeN
        a = ea(k): b = eb(k)
        If ee2(k) = 0 Then
            bn(a) = bn(a) + 1: bn(b) = bn(b) + 1
            If ba(a) = 0 Then ba(a) = b Else bb(a) = b
            If ba(b) = 0 Then ba(b) = a Else bb(b) = a
        Else
            ni(a) = b: ni(b) = a
        End If
    Next k
    For n = 1 To nn
        If bn(n) = 2 And incident(n).count = 2 And ni(n) > 0 Then
            p = ni(n): If locked(p) Then GoTo NextCorner
            a = ba(n): b = bb(n): ux = xx(a) - xx(n): uy = yy(a) - yy(n): vx = xx(b) - xx(n): vy = yy(b) - yy(n)
            ln = Sqr(ux * ux + uy * uy): ux = ux / ln: uy = uy / ln
            ln = Sqr(vx * vx + vy * vy): vx = vx / ln: vy = vy / ln
            If Abs(ux * vx + uy * vy) < 0.00000001 Then
                ux = ux + vx: uy = uy + vy: ln = Sqr(ux * ux + uy * uy): ux = ux / ln: uy = uy / ln
                t = (xx(p) - xx(n)) * ux + (yy(p) - yy(n)) * uy
                If t <= 0# Then GoTo NextCorner
                ox = xx(p): oy = yy(p): xx(p) = xx(n) + t * ux: yy(p) = yy(n) + t * uy
                valid = True
                For Each item In incident(p)
                    If Not Metric(CLng(item), ar, energy, angle, aspect) Then valid = False: Exit For
                Next item
                If Not valid Then xx(p) = ox: yy(p) = oy
            End If
        End If
NextCorner:
    Next n
End Sub

Private Function GenerateAndCompare() As Boolean
    Dim source As CandidateState, best As CandidateState, terrainBase As CandidateState
    Dim bestQ As QualityRecord, q As QualityRecord, baseQ As QualityRecord, record As String, method As String, found As Boolean
    Dim off As Long, target As Long, ratio As Double, bestIsTerrain As Boolean
    bestIsTerrain = False
    Call CaptureState(source)
    If candidateKind = "PLASTIC" Then
        Call RefinePlastic
        sourceIndicator = refineScore
    End If
    Call CheckCandidate
    Call GetQuality(bestQ)
    If candidateKind = "PLASTIC" Then plasticHeatLimit = bestQ.heatArea * 1.05 Else plasticHeatLimit = 0#
    Call CaptureState(best): selectedMethod = "既存接続"
    comparisonText = QualityLine("最適化前", bestQ)
    Application.StatusBar = "メッシュ候補：外周・接続を維持して角度を最適化"
    Call OptimizeAngles
    Call GetQuality(q)
    comparisonText = comparisonText & vbCrLf & QualityLine("角度最適化", q)
    If BetterQuality(q, bestQ) Then bestQ = q: Call CaptureState(best): selectedMethod = "角度最適化（外周・接続維持）"
    If FEMReadSetting("MESH_REPAIR_TERRAIN_COMPARE", 1#) < 0.5 Then
        comparisonText = comparisonText & vbCrLf & "地形追従：比較設定OFF"
        GoTo Done
    End If
    Call RestoreState(source)
    Application.StatusBar = "メッシュ候補：地形追従格子の適用条件を確認"
    If Not TerrainEligible() Then
        comparisonText = comparisonText & vbCrLf & "地形追従：対象外（" & terrainReason & "）"
        GoTo Done
    End If
    On Error GoTo TerrainFailed
    For off = -16 To 16
        If BuildTerrain(off, False) Then
            Call CheckCandidate
            Call GetQuality(q, False)
            If Not found Then
                found = True: baseQ = q: Call CaptureState(terrainBase)
            ElseIf BetterQuality(q, baseQ) Then
                baseQ = q: Call CaptureState(terrainBase)
            End If
        End If
    Next off
    If Not found Then
        For off = -16 To 16
            If BuildTerrain(off, True) Then
                Call CheckCandidate
                Call GetQuality(q, False)
                If Not found Then
                    found = True: baseQ = q: Call CaptureState(terrainBase)
                ElseIf BetterQuality(q, baseQ) Then
                    baseQ = q: Call CaptureState(terrainBase)
                End If
            End If
        Next off
    End If
    If Not found Then terrainReason = "地形追従格子を分割できません": GoTo TerrainSkipped
    Call RestoreState(terrainBase)
    If candidateKind = "PLASTIC" Then
        Call PlasticResolution(True)
        target = CLng(Fix(originalNE * FEMReadSetting("ADAPT_MAX_ELEMENT_RATIO", 2#)))
        Call RefinePlastic(True, target)
    End If
    Application.StatusBar = "メッシュ候補：地形追従候補の角度とQ8品質を検査"
    Call CheckCandidate
    Call GetQuality(q)
    comparisonText = comparisonText & vbCrLf & QualityLine("地形追従", q)
    If candidateKind = "PLASTIC" And q.heatArea > plasticHeatLimit Then
        comparisonText = comparisonText & "（塑性域の分解能不足で除外）"
    ElseIf PreferTerrainOver(q, bestQ, bestIsTerrain) Then
        bestQ = q: Call CaptureState(best): selectedMethod = "地形追従格子": bestIsTerrain = True
    End If
    Call OptimizeAngles
    Call GetQuality(q)
    comparisonText = comparisonText & vbCrLf & QualityLine("地形追従＋角度最適化", q)
    If candidateKind = "PLASTIC" And q.heatArea > plasticHeatLimit Then
        comparisonText = comparisonText & "（塑性域の分解能不足で除外）"
    ElseIf PreferTerrainOver(q, bestQ, bestIsTerrain) Then
        bestQ = q: Call CaptureState(best): selectedMethod = "地形追従＋角度最適化": bestIsTerrain = True
    End If
    GoTo Done
TerrainFailed:
    terrainReason = Err.Description
    Resume TerrainSkipped
TerrainSkipped:
    comparisonText = comparisonText & vbCrLf & "地形追従：候補除外（" & terrainReason & "）"
Done:
    On Error GoTo 0
    If FEMReadSetting("MESH_GEN_COMPARE", 1#) >= 1.5 Then
        comparisonText = comparisonText & vbCrLf & "採用優先：両方合格なら地形格子 (MESH_GEN_COMPARE=2)"
    ElseIf FEMReadSetting("MESH_GEN_COMPARE", 1#) >= 0.5 Then
        comparisonText = comparisonText & vbCrLf & "採用優先：両方合格ならデローニ修正 (MESH_GEN_COMPARE=1)"
    Else
        comparisonText = comparisonText & vbCrLf & "採用優先：両方合格なら品質の良い方 (MESH_GEN_COMPARE=0)"
    End If
    Call RestoreState(best)
    Call CheckCandidate
    MeshAdaptLastSummary = "選択：" & selectedMethod & " / 要素 " & originalNE & " → " & ne & " / 節点 " & originalNN & " → " & nn & vbCrLf & _
        comparisonText & vbCrLf & "合格45～135° / 警告30～150°内 / その他は不合格。採用後は再解析が必要です。"
    If bestQ.failN > 0 Then MeshAdaptLastSummary = MeshAdaptLastSummary & vbCrLf & "不合格要素が残っています。外形・拘束と角度基準を確認してください。"
    GenerateAndCompare = True
End Function

Private Function OnSegment(ByVal x As Double, ByVal y As Double, ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal endY As Double) As Boolean
    Dim tol As Double: tol = meshScale * 0.000000001
    If Abs((c - a) * (y - b) - (endY - b) * (x - a)) > tol * meshScale Then Exit Function
    OnSegment = (x - a) * (x - c) + (y - b) * (y - endY) <= tol * tol
End Function
Private Function TerrainEligible() As Boolean
    Dim e As Long, k As Long, i As Long, j As Long, n As Long, fx As Long, fy As Long, found As Boolean, rows As Variant, token As String
    terrainReason = "": segN = 0
    If HasInteriorBoundary() Then terrainReason = "穴の境界・接続を保持するため、角度最適化を使用": Exit Function
    For e = 2 To ne
        If materialId(e) <> materialId(1) Then terrainReason = "複数材料の領域境界を保持するため": Exit Function
    Next e
    For i = 1 To nn
        For j = 6 To 14
            If d(nodeRows(i, j)) <> 0# Then terrainReason = "指定変位・集中荷重・追加節点条件を保持するため": Exit Function
        Next j
    Next i
    rows = Table("載荷", 9)
    For i = 1 To UBound(rows, 1)
        If left$(Trim$(CStr(rows(i, 1))), 1) <> "(" Then
            token = UCase$(Trim$(CStr(rows(i, 4))))
            If token = "FORCE" Then terrainReason = "節点ごとのFORCE荷重の作用点を保持するため": Exit Function
            If token = "FIX" Or token = "DISP" Then
                token = UCase$(Trim$(CStr(rows(i, 2))))
                If token <> "TOP" And token <> "BOTTOM" And token <> "LEFT" And token <> "RIGHT" Then terrainReason = "位置・番号指定の載荷条件を保持するため": Exit Function
            End If
        End If
    Next i
    ReDim sx1(1 To edgeN): ReDim sy1(1 To edgeN): ReDim sx2(1 To edgeN): ReDim sy2(1 To edgeN)
    ReDim sFixX(1 To edgeN): ReDim sFixY(1 To edgeN)
    For k = 1 To edgeN
        If Abs(xx(em(k)) - (xx(ea(k)) + xx(eb(k))) / 2#) + Abs(yy(em(k)) - (yy(ea(k)) + yy(eb(k))) / 2#) > meshScale * 0.000000001 Then terrainReason = "曲線Q8の形状を保持するため": Exit Function
        If ee2(k) = 0 Then
            segN = segN + 1: sx1(segN) = xx(ea(k)): sy1(segN) = yy(ea(k)): sx2(segN) = xx(eb(k)): sy2(segN) = yy(eb(k))
            sFixX(segN) = CLng(d(nodeRows(em(k), 4))): sFixY(segN) = CLng(d(nodeRows(em(k), 5)))
        End If
    Next k
    For i = 1 To nn
        fx = 0: fy = 0
        For k = 1 To segN
            If OnSegment(xx(i), yy(i), sx1(k), sy1(k), sx2(k), sy2(k)) Then
                If sFixX(k) <> 0 Then fx = sFixX(k)
                If sFixY(k) <> 0 Then fy = sFixY(k)
            End If
        Next k
        If fx <> CLng(d(nodeRows(i, 4))) Or fy <> CLng(d(nodeRows(i, 5))) Then terrainReason = "内部拘束・局所的な境界条件を保持するため": Exit Function
    Next i
    TerrainEligible = True
End Function
Private Sub SetTerrainConditions(ByVal n As Long)
    Dim k As Long
    nodeRows(n, 4) = 0#: nodeRows(n, 5) = 0#
    For k = 1 To segN
        If OnSegment(xx(n), yy(n), sx1(k), sy1(k), sx2(k), sy2(k)) Then
            If sFixX(k) <> 0 Then nodeRows(n, 4) = sFixX(k)
            If sFixY(k) <> 0 Then nodeRows(n, 5) = sFixY(k)
        End If
    Next k
End Sub
Private Sub ValidateTerrainConditions()
    Dim i As Long, k As Long, fx As Long, fy As Long
    For i = 1 To nn
        fx = 0: fy = 0
        For k = 1 To segN
            If OnSegment(xx(i), yy(i), sx1(k), sy1(k), sx2(k), sy2(k)) Then
                If sFixX(k) <> 0 Then fx = sFixX(k)
                If sFixY(k) <> 0 Then fy = sFixY(k)
            End If
        Next k
        If fx <> CLng(d(nodeRows(i, 4))) Or fy <> CLng(d(nodeRows(i, 5))) Then Err.Raise vbObjectError + 3860, , "地形追従候補の境界条件が一致しません。"
    Next i
End Sub
Private Function BuildTerrain(ByVal offset As Long, Optional ByVal ignoreCountWindow As Boolean = False) As Boolean
    Dim xmin As Double, xmax As Double, bottom As Double, tol As Double, k As Long, i As Long, j As Long, t As Long, n As Long
    Dim tx1() As Double, ty1() As Double, tx2() As Double, ty2() As Double, tfX() As Long, tfY() As Long, order() As Long, nt As Long
    Dim px() As Double, py() As Double, masksX() As Long, masksY() As Long, np As Long, a As Double, b As Double, c As Double, endY As Double, z As Double
    Dim nx As Long, ny As Long, count() As Long, total As Long, target As Long, ideal As Double, best As Double, at As Long, h As Double
    Dim gridX() As Double, gridH() As Double, col As Long, row As Long, e As Long, v1 As Long, v2 As Long, v3 As Long, v4 As Long, mid As Long, key As String, mids As Object
    Dim bcBottomX As Long, bcBottomY As Long, bcLeftX As Long, bcLeftY As Long, bcRightX As Long, bcRightY As Long
    If terrainPrepared Then
        xmin = terrainLeft: xmax = terrainRight: bottom = terrainBottom
        px = terrainPX: py = terrainPY: np = terrainNP
        GoTo TerrainGeometryReady
    End If
    xmin = d(oldNodes(1, 2)): xmax = xmin: bottom = d(oldNodes(1, 3)): tol = meshScale * 0.000000001
    For i = 1 To originalNN
        z = d(oldNodes(i, 2)): If z < xmin Then xmin = z
        If z > xmax Then xmax = z
        z = d(oldNodes(i, 3)): If z < bottom Then bottom = z
    Next i
    ReDim tx1(1 To segN): ReDim ty1(1 To segN): ReDim tx2(1 To segN): ReDim ty2(1 To segN): ReDim order(1 To segN)
    ReDim tfX(1 To segN): ReDim tfY(1 To segN)
    bcBottomX = -1: bcLeftX = -1: bcRightX = -1
    For k = 1 To segN
        a = sx1(k): b = sy1(k): c = sx2(k): endY = sy2(k)
        If Abs(b - bottom) < tol And Abs(endY - bottom) < tol Then
            If bcBottomX >= 0 Then
                If bcBottomX <> sFixX(k) Or bcBottomY <> sFixY(k) Then terrainReason = "底辺の部分拘束を保持するため": Exit Function
            End If
            bcBottomX = sFixX(k): bcBottomY = sFixY(k)
        ElseIf Abs(a - xmin) < tol And Abs(c - xmin) < tol Then
            If bcLeftX >= 0 Then
                If bcLeftX <> sFixX(k) Or bcLeftY <> sFixY(k) Then terrainReason = "側辺の部分拘束を保持するため": Exit Function
            End If
            bcLeftX = sFixX(k): bcLeftY = sFixY(k)
        ElseIf Abs(a - xmax) < tol And Abs(c - xmax) < tol Then
            If bcRightX >= 0 Then
                If bcRightX <> sFixX(k) Or bcRightY <> sFixY(k) Then terrainReason = "側辺の部分拘束を保持するため": Exit Function
            End If
            bcRightX = sFixX(k): bcRightY = sFixY(k)
        Else
            If Abs(c - a) < tol Or b <= bottom + tol Or endY <= bottom + tol Then terrainReason = "水平底辺・鉛直側辺・高さ関数の地表ではありません": Exit Function
            If c < a Then z = a: a = c: c = z: z = b: b = endY: endY = z
            nt = nt + 1: tx1(nt) = a: ty1(nt) = b: tx2(nt) = c: ty2(nt) = endY: order(nt) = nt
            tfX(nt) = sFixX(k): tfY(nt) = sFixY(k)
        End If
    Next k
    If nt = 0 Or bcBottomX < 0 Or bcLeftX < 0 Or bcRightX < 0 Then terrainReason = "地形追従格子の外形条件に適合しません": Exit Function
    For i = 1 To nt - 1
        For j = i + 1 To nt
            If tx1(order(j)) < tx1(order(i)) Then t = order(i): order(i) = order(j): order(j) = t
        Next j
    Next i
    If Abs(tx1(order(1)) - xmin) > tol Or Abs(tx2(order(nt)) - xmax) > tol Then terrainReason = "地表が側辺につながりません": Exit Function
    ReDim px(0 To nt): ReDim py(0 To nt): ReDim masksX(1 To nt): ReDim masksY(1 To nt)
    px(0) = tx1(order(1)): py(0) = ty1(order(1))
    For i = 1 To nt
        k = order(i)
        If i > 1 Then
            j = order(i - 1)
            If Abs(tx1(k) - tx2(j)) > tol Or Abs(ty1(k) - ty2(j)) > tol Then terrainReason = "穴・複数領域・張り出しを含むため": Exit Function
        End If
        If np > 0 Then
            a = px(np) - px(np - 1): b = py(np) - py(np - 1): c = tx2(k) - px(np): endY = ty2(k) - py(np)
            If Abs(a * endY - b * c) < tol * meshScale And masksX(np) = tfX(k) And masksY(np) = tfY(k) Then
                px(np) = tx2(k): py(np) = ty2(k): GoTo NextTop
            End If
        End If
        np = np + 1: px(np) = tx2(k): py(np) = ty2(k): masksX(np) = tfX(k): masksY(np) = tfY(k)
NextTop:
    Next i
    terrainPX = px: terrainPY = py: terrainNP = np
    terrainLeft = xmin: terrainRight = xmax: terrainBottom = bottom: terrainPrepared = True
TerrainGeometryReady:
    ideal = 0#
    For i = 1 To np: ideal = ideal + (px(i) - px(i - 1)) * ((py(i) + py(i - 1)) / 2# - bottom): Next i
    ny = CLng(Sqr(originalNE * ideal / (xmax - xmin) ^ 2)) + offset
    If ny < 1 Then Exit Function
    nx = CLng(CDbl(originalNE) / CDbl(ny))
    If nx < np Then nx = np
    If Not ignoreCountWindow Then
        If CDbl(nx * ny) < CDbl(originalNE) * FEMReadSetting("MESH_TERRAIN_COUNT_MIN", 0.55) Then Exit Function
        If CDbl(nx * ny) > CDbl(originalNE) * FEMReadSetting("MESH_TERRAIN_COUNT_MAX", 1.8) Then Exit Function
    End If
    ReDim count(1 To np): total = np
    For i = 1 To np: count(i) = 1: Next i
    Do While total < nx
        best = -1#: at = 1
        For i = 1 To np
            ideal = (px(i) - px(i - 1)) / count(i)
            If ideal > best Then best = ideal: at = i
        Next i
        count(at) = count(at) + 1: total = total + 1
    Loop
    ReDim gridX(0 To nx): ReDim gridH(0 To nx): col = 0
    For i = 1 To np
        For j = 0 To count(i) - 1
            gridX(col) = px(i - 1) + (px(i) - px(i - 1)) * j / count(i)
            gridH(col) = py(i - 1) + (py(i) - py(i - 1)) * j / count(i) - bottom: col = col + 1
        Next j
    Next i
    gridX(nx) = xmax: gridH(nx) = py(np) - bottom
    ne = nx * ny: nn = 0
    ReDim xx(1 To ne * 30 + 100): ReDim yy(1 To UBound(xx)): ReDim nodeRows(1 To UBound(xx), 1 To 14)
    ReDim conn(1 To ne * 4, 1 To 8): ReDim materialId(1 To ne * 4): ReDim candidateScore(1 To ne * 4)
    For row = 0 To ny
        For col = 0 To nx
            n = AddPoint(gridX(col), bottom + gridH(col) * row / ny): Call SetTerrainConditions(n)
        Next col
    Next row
    Set mids = CreateObject("Scripting.Dictionary"): e = 0
    For row = 0 To ny - 1
        For col = 0 To nx - 1
            e = e + 1: v1 = row * (nx + 1) + col + 1: v2 = v1 + 1: v4 = v1 + nx + 1: v3 = v4 + 1
            conn(e, 1) = v1: conn(e, 2) = v2: conn(e, 3) = v3: conn(e, 4) = v4: materialId(e) = CLng(oldElements(1, 10))
            For j = 1 To 4
                v1 = conn(e, j): v2 = conn(e, j Mod 4 + 1): key = EdgeKey(v1, v2)
                If mids.Exists(key) Then
                    mid = CLng(mids(key))
                Else
                    mid = AddPoint((xx(v1) + xx(v2)) / 2#, (yy(v1) + yy(v2)) / 2#): mids.Add key, mid: Call SetTerrainConditions(mid)
                End If
                conn(e, j + 4) = mid
            Next j
        Next col
    Next row
    candidateRebuilt = True
    Call BuildEdges
    Call CacheMetrics
    BuildTerrain = True
End Function

' Exact overlap of straight convex parent/candidate quadrilaterals for indicator transport.
Private Function OverlapArea(ByVal e As Long, ByVal source As Long) As Double
    Dim px(0 To 15) As Double, py(0 To 15) As Double, qx(0 To 15) As Double, qy(0 To 15) As Double
    Dim n As Long, m As Long, i As Long, j As Long, k As Long, a As Long, b As Long, ax As Double, ay As Double, bx As Double, by As Double
    Dim d1 As Double, d2 As Double, t As Double, ux As Double, uy As Double
    n = 4
    For i = 0 To 3: px(i) = xx(conn(e, i + 1)): py(i) = yy(conn(e, i + 1)): Next i
    For k = 0 To 3
        a = CLng(oldElements(source, k + 2)): b = CLng(oldElements(source, (k + 1) Mod 4 + 2))
        ax = srcX(source, k): ay = srcY(source, k): bx = srcX(source, (k + 1) Mod 4): by = srcY(source, (k + 1) Mod 4): ux = bx - ax: uy = by - ay
        m = 0
        For i = 0 To n - 1
            j = (i + 1) Mod n: d1 = ux * (py(i) - ay) - uy * (px(i) - ax): d2 = ux * (py(j) - ay) - uy * (px(j) - ax)
            If d1 >= -0.000000000001 Then qx(m) = px(i): qy(m) = py(i): m = m + 1
            If (d1 > 0# And d2 < 0#) Or (d1 < 0# And d2 > 0#) Then
                t = d1 / (d1 - d2): qx(m) = px(i) + t * (px(j) - px(i)): qy(m) = py(i) + t * (py(j) - py(i)): m = m + 1
            End If
        Next i
        n = m: If n < 3 Then Exit Function
        For i = 0 To n - 1: px(i) = qx(i): py(i) = qy(i): Next i
    Next k
    For i = 0 To n - 1: j = (i + 1) Mod n: OverlapArea = OverlapArea + px(i) * py(j) - px(j) * py(i): Next i
    OverlapArea = Abs(OverlapArea) / 2#
End Function
Private Function PlasticResolution(ByVal transport As Boolean) As Double
    Dim e As Long, s As Long, j As Long, n As Long, ix As Long, iy As Long, key As String, item As Variant
    Dim a As Double, b As Double, c As Double, endY As Double, xp As Double, yp As Double, weight As Double, total As Double, overlap As Double
    Dim hits() As Long, count As Long, hit As Long
    If Not transport And heatVersion = geomVersion Then PlasticResolution = heatCache: Exit Function
    PrepareSourceGrid
    ReDim hits(1 To originalNE)
    If transport Then ReDim refineScore(1 To ne)
    For e = 1 To ne
        a = 1E+100: b = -1E+100: c = 1E+100: endY = -1E+100
        For j = 1 To 4
            n = conn(e, j): xp = xx(n): yp = yy(n)
            If xp < a Then a = xp
            If xp > b Then b = xp
            If yp < c Then c = yp
            If yp > endY Then endY = yp
        Next j
        If sourceEpoch = 2147483647 Then ReDim srcSeen(1 To originalNE): sourceEpoch = 0
        sourceEpoch = sourceEpoch + 1: count = 0
        For ix = CLng(Int(a / sourceCell)) To CLng(Int(b / sourceCell))
            For iy = CLng(Int(c / sourceCell)) To CLng(Int(endY / sourceCell))
                key = CStr(ix) & ":" & CStr(iy)
                If sourceGrid.Exists(key) Then
                    For Each item In sourceGrid(key)
                        s = CLng(item)
                        If srcSeen(s) <> sourceEpoch Then
                            srcSeen(s) = sourceEpoch
                            If materialId(e) = srcMat(s) Then
                                If srcHiX(s) > a And srcLoX(s) < b And srcHiY(s) > c And srcLoY(s) < endY Then count = count + 1: hits(count) = s
                            End If
                        End If
                    Next item
                End If
            Next iy
        Next ix
        If count > 1 Then SortIds hits, 1, count
        weight = 0#
        For hit = 1 To count
            s = hits(hit): overlap = OverlapArea(e, s)
            If overlap > meshScale * meshScale * 0.0000000000001 Then
                weight = weight + overlap * sourceIndicator(s)
                If transport Then If sourceIndicator(s) > refineScore(e) Then refineScore(e) = sourceIndicator(s)
            End If
        Next hit
        PlasticResolution = PlasticResolution + weight * areas(e): total = total + weight
    Next e
    If total > 0# Then PlasticResolution = PlasticResolution / total Else Err.Raise vbObjectError + 3861, , "塑性域の位置対応に失敗しました。"
    If Not transport Then heatVersion = geomVersion: heatCache = PlasticResolution
End Function


Private Sub EnsureJacobianCoefficients()
    Dim i As Long, j As Long, k As Long, h As Double, u As Double, v As Double, p(0 To 7) As Double, m(0 To 7) As Double
    If jacReady Then Exit Sub
    h = 0.00001
    For i = 0 To 4
        u = -1# + i * 0.5
        For j = 0 To 4
            v = -1# + j * 0.5
            ShapeQ8 u + h, v, p: ShapeQ8 u - h, v, m
            For k = 0 To 7: jacDu(i * 5 + j, k) = p(k) - m(k): Next k
            ShapeQ8 u, v + h, p: ShapeQ8 u, v - h, m
            For k = 0 To 7: jacDv(i * 5 + j, k) = p(k) - m(k): Next k
        Next j
    Next i
    jacReady = True
End Sub
Private Sub EnsureRefinementCapacity()
    Dim cap As Long, i As Long, j As Long, rows() As Variant, elements() As Long
    cap = nn + ne * 25
    If cap > 300000 Then Err.Raise vbObjectError + 3801, , "再メッシュの作業上限を超えます。"
    If UBound(xx) < cap Then ReDim Preserve xx(1 To cap): ReDim Preserve yy(1 To cap)
    If UBound(nodeRows, 1) < cap Then
        ReDim rows(1 To cap, 1 To 14)
        For i = 1 To nn: For j = 1 To 14: rows(i, j) = nodeRows(i, j): Next j: Next i
        nodeRows = rows
    End If
    If UBound(conn, 1) < ne * 4 Then
        ReDim elements(1 To ne * 4, 1 To 8)
        For i = 1 To ne: For j = 1 To 8: elements(i, j) = conn(i, j): Next j: Next i
        conn = elements
    End If
    If UBound(materialId) < ne * 4 Then ReDim Preserve materialId(1 To ne * 4)
End Sub

Private Sub PrepareSourceGrid()
    Dim e As Long, j As Long, n As Long, ix As Long, iy As Long, key As String, bucket As Collection
    If Not sourceGrid Is Nothing Then Exit Sub
    Set sourceGrid = CreateObject("Scripting.Dictionary")
    sourceCell = meshScale / Sqr(CDbl(originalNE)): If sourceCell <= 0# Then Err.Raise vbObjectError + 3862, , "空間索引の寸法が不正です。"
    ReDim srcLoX(1 To originalNE): ReDim srcHiX(1 To originalNE): ReDim srcLoY(1 To originalNE): ReDim srcHiY(1 To originalNE)
    ReDim srcX(1 To originalNE, 0 To 3): ReDim srcY(1 To originalNE, 0 To 3): ReDim srcMat(1 To originalNE)
    ReDim srcSeen(1 To originalNE): sourceEpoch = 0
    For e = 1 To originalNE
        srcLoX(e) = 1E+100: srcLoY(e) = 1E+100: srcHiX(e) = -1E+100: srcHiY(e) = -1E+100
        srcMat(e) = CLng(oldElements(e, 10))
        For j = 0 To 3
            n = CLng(oldElements(e, j + 2)): srcX(e, j) = d(oldNodes(n, 2)): srcY(e, j) = d(oldNodes(n, 3))
            If srcX(e, j) < srcLoX(e) Then srcLoX(e) = srcX(e, j)
            If srcX(e, j) > srcHiX(e) Then srcHiX(e) = srcX(e, j)
            If srcY(e, j) < srcLoY(e) Then srcLoY(e) = srcY(e, j)
            If srcY(e, j) > srcHiY(e) Then srcHiY(e) = srcY(e, j)
        Next j
        If sourceIndicator(e) > 0# Then
            For ix = CLng(Int(srcLoX(e) / sourceCell)) To CLng(Int(srcHiX(e) / sourceCell))
                For iy = CLng(Int(srcLoY(e) / sourceCell)) To CLng(Int(srcHiY(e) / sourceCell))
                    key = CStr(ix) & ":" & CStr(iy)
                    If Not sourceGrid.Exists(key) Then Set bucket = New Collection: sourceGrid.Add key, bucket
                    sourceGrid(key).Add e
                Next iy
            Next ix
        End If
    Next e
End Sub
Private Sub SortIds(ByRef ids() As Long, ByVal first As Long, ByVal last As Long)
    Dim lo As Long, hi As Long, pivot As Long, swap As Long
    lo = first: hi = last: pivot = ids((first + last) \ 2)
    Do While lo <= hi
        Do While ids(lo) < pivot: lo = lo + 1: Loop
        Do While ids(hi) > pivot: hi = hi - 1: Loop
        If lo <= hi Then swap = ids(lo): ids(lo) = ids(hi): ids(hi) = swap: lo = lo + 1: hi = hi - 1
    Loop
    If first < hi Then SortIds ids, first, hi
    If lo < last Then SortIds ids, lo, last
End Sub

' Segment tree reproduces selection-sort swaps, including its original tie ordering.
Private Function RankWinner(ByVal a As Long, ByVal b As Long, ByRef ranked() As Long, ByRef scores() As Double) As Long
    If a = 0 Then RankWinner = b: Exit Function
    If b = 0 Then RankWinner = a: Exit Function
    If scores(ranked(a)) > scores(ranked(b)) Then
        RankWinner = a
    ElseIf scores(ranked(a)) < scores(ranked(b)) Then
        RankWinner = b
    ElseIf a < b Then
        RankWinner = a
    Else
        RankWinner = b
    End If
End Function
Private Sub RankScores(ByRef scores() As Double, ByRef ranked() As Long, ByVal count As Long)
    Dim size As Long, tree() As Long, i As Long, j As Long, left As Long, right As Long, best As Long, swap As Long, p As Long
    size = 1: Do While size < count: size = size * 2: Loop
    ReDim tree(1 To size * 2)
    For i = 1 To count: tree(size + i - 1) = i: Next i
    For i = size - 1 To 1 Step -1: tree(i) = RankWinner(tree(i * 2), tree(i * 2 + 1), ranked, scores): Next i
    For i = 1 To count - 1
        left = size + i - 1: right = size + count - 1: best = 0
        Do While left <= right
            If left Mod 2 = 1 Then best = RankWinner(best, tree(left), ranked, scores): left = left + 1
            If right Mod 2 = 0 Then best = RankWinner(best, tree(right), ranked, scores): right = right - 1
            left = left \ 2: right = right \ 2
        Loop
        swap = ranked(i): ranked(i) = ranked(best): ranked(best) = swap
        For j = 0 To 1
            If j = 0 Then p = (size + i - 1) \ 2 Else p = (size + best - 1) \ 2
            Do While p > 0
                tree(p) = RankWinner(tree(p * 2), tree(p * 2 + 1), ranked, scores): p = p \ 2
            Loop
        Next j
    Next i
End Sub

' Dirty-element scan in the same ascending sweep order as the original full scans.
' Round/e ordering is essential: different closure order can over-refine a mesh.
Private Sub MarkQueueAdd(ByVal e As Long, ByVal afterKey As Double)
    Dim key As Double, p As Long, parent As Long
    If e = 0 Then Exit Sub
    key = Fix(afterKey / ne) * ne + e
    If key <= afterKey Then key = key + ne
    If markDue(e) > 0# Then If markDue(e) <= key Then Exit Sub
    markDue(e) = key: markHeapN = markHeapN + 1
    If markHeapN > UBound(markHeap) Then ReDim Preserve markHeap(1 To UBound(markHeap) * 2)
    p = markHeapN
    Do While p > 1
        parent = p \ 2: If markHeap(parent) <= key Then Exit Do
        markHeap(p) = markHeap(parent): p = parent
    Loop
    markHeap(p) = key
End Sub
Private Function MarkQueuePop() As Double
    Dim tail As Double, p As Long, child As Long
    MarkQueuePop = markHeap(1): tail = markHeap(markHeapN): markHeapN = markHeapN - 1
    If markHeapN = 0 Then Exit Function
    p = 1
    Do While p * 2 <= markHeapN
        child = p * 2
        If child < markHeapN Then If markHeap(child + 1) < markHeap(child) Then child = child + 1
        If tail <= markHeap(child) Then Exit Do
        markHeap(p) = markHeap(child): p = child
    Loop
    markHeap(p) = tail
End Function
Private Sub MarkOneEdge(ByVal edge As Long, ByRef marked() As Boolean, ByVal afterKey As Double)
    If marked(edge) Then Exit Sub
    marked(edge) = True: markLogN = markLogN + 1: markLog(markLogN) = edge
    MarkQueueAdd ee1(edge), afterKey: MarkQueueAdd ee2(edge), afterKey
End Sub
Private Function TryMarkElement(ByVal element As Long, ByRef marked() As Boolean, ByVal budget As Long) As Boolean
    Dim e As Long, j As Long, n As Long, last As Long, total As Long, key As Double, i As Long
    ReDim markDue(1 To ne): ReDim markHeap(1 To ne + 8): ReDim markLog(1 To edgeN)
    markHeapN = 0: markLogN = 0
    For j = 0 To 3: MarkOneEdge elementEdge(element, j), marked, 0#: Next j
    Do While markHeapN > 0
        key = MarkQueuePop(): e = CLng(key - Fix((key - 1#) / ne) * ne)
        If markDue(e) = key Then
            markDue(e) = 0#: n = 0: last = 0
            For j = 0 To 3: If marked(elementEdge(e, j)) Then n = n + 1: last = j
            Next j
            If n = 1 Then
                MarkOneEdge elementEdge(e, (last + 2) Mod 4), marked, key
            ElseIf n = 3 Then
                For j = 0 To 3: MarkOneEdge elementEdge(e, j), marked, key: Next j
            End If
        End If
    Loop
    For e = 1 To ne
        n = 0: last = -1
        For j = 0 To 3: If marked(elementEdge(e, j)) Then n = n + 1: If last < 0 Then last = j
        Next j
        Select Case n
            Case 0: total = total + 1
            Case 2
                If marked(elementEdge(e, (last + 2) Mod 4)) Then total = total + 2 Else total = total + 3
            Case 4: total = total + 4
            Case Else: Err.Raise vbObjectError + 3805, , "移行要素の辺分割が不整合です。"
        End Select
    Next e
    TryMarkElement = (total <= budget)
    If Not TryMarkElement Then For i = 1 To markLogN: marked(markLog(i)) = False: Next i
End Function

Private Function HasInteriorBoundary() As Boolean
    Dim nextNode() As Long, seen() As Boolean, k As Long, start As Long, n As Long, m As Long
    Dim twiceArea As Double, ox As Double, oy As Double
    ReDim nextNode(1 To nn): ReDim seen(1 To nn)
    ' Element corners are counterclockwise; a hole loop is clockwise.
    For k = 1 To edgeN
        If ee2(k) = 0 Then nextNode(ea(k)) = eb(k)
    Next k
    For start = 1 To nn
        If nextNode(start) > 0 And Not seen(start) Then
            n = start: twiceArea = 0#: ox = xx(start): oy = yy(start)
            Do
                seen(n) = True: m = nextNode(n)
                If m = 0 Then Err.Raise vbObjectError + 3815, , "メッシュ境界が閉じていません。"
                twiceArea = twiceArea + (xx(n) - ox) * (yy(m) - oy) - (xx(m) - ox) * (yy(n) - oy)
                n = m
            Loop Until seen(n)
            If n <> start Then Err.Raise vbObjectError + 3815, , "メッシュ境界の接続が不正です。"
            If twiceArea < -meshScale * meshScale * 0.000000000001 Then HasInteriorBoundary = True: Exit Function
        End If
    Next start
End Function



