Option Explicit
Public Function QualityRegression() As String
    On Error GoTo Failed
    Dim x() As Double, y() As Double, rows() As Variant, conditions As Variant, cn() As Long
    Dim edges As Object, k As Long, e As Long, j As Long, a As Long, b As Long, mid As Long, key As String
    Dim moved As Long, sweeps As Long, before As Double, after As Double, verts As Variant
    SuppressUserMessages = True
    FEMWriteNumericSetting "MESH_QUALITY_MIN_ANGLE_DEG", 45#
    FEMWriteNumericSetting "MESH_QUALITY_MAX_ANGLE_DEG", 135#
    FEMWriteNumericSetting "MESH_T6_MIN_ANGLE_DEG", 20#
    FEMWriteNumericSetting "MESH_T6_MAX_ANGLE_DEG", 140#
    If Not MeshShapeWithinTargets(True, 30#, 120#, 2#) Then Err.Raise 5, , "T6 uses Q8 angle targets"
    If MeshShapeWithinTargets(False, 30#, 120#, 2#) Then Err.Raise 5, , "Q8 angle target lost"
    If MeshShapeWithinTargets(True, 10#, 100#, 3#) Then Err.Raise 5, , "sharp T6 warning hidden"
    If Not MeshQ8JacobianWithinTarget(0.0001, 0.000025) Or Not MeshQ8JacobianWithinTarget(100#, 25#) Then Err.Raise 5, , "size-dependent Q8 shape decision"
    If MeshQ8JacobianWithinTarget(1#, 0.001) Then Err.Raise 5, , "distorted Q8 not detected"
    ReDim x(1 To 13): ReDim y(1 To 13): ReDim rows(1 To 4, 1 To 11): ReDim cn(1 To 4, 1 To 8)
    ReDim conditions(1 To 13, 1 To 14)
    x(2) = 1#: x(3) = 1#: y(3) = 1#: y(4) = 1#: x(5) = 0.2: y(5) = 0.4
    Set edges = CreateObject("Scripting.Dictionary"): mid = 5: before = 1#
    For e = 1 To 4
        a = e: b = e Mod 4 + 1
        rows(e, 1) = e: rows(e, 2) = a: rows(e, 3) = b: rows(e, 4) = 5: rows(e, 5) = 5: rows(e, 8) = 5: rows(e, 10) = 1: rows(e, 11) = "T6"
        verts = Array(a, b, 5)
        For j = 0 To 2
            a = verts(j): b = verts((j + 1) Mod 3)
            If a < b Then key = CStr(a) & ":" & CStr(b) Else key = CStr(b) & ":" & CStr(a)
            If Not edges.Exists(key) Then
                mid = mid + 1: edges.Add key, mid: x(mid) = (x(a) + x(b)) / 2#: y(mid) = (y(a) + y(b)) / 2#
            End If
            rows(e, IIf(j = 2, 9, j + 6)) = edges(key)
        Next j
        before = Application.Min(before, MeshTriangleScore(x, y, CLng(rows(e, 2)), CLng(rows(e, 3)), 5))
    Next e
    MeshImproveT6 x, y, rows, conditions, moved, sweeps, False
    after = 1#
    For e = 1 To 4
        after = Application.Min(after, MeshTriangleScore(x, y, CLng(rows(e, 2)), CLng(rows(e, 3)), 5))
        For j = 1 To 8: cn(e, j) = CLng(rows(e, j + 1)): Next j
    Next e
    If moved <> 1 Or after <= before Or Abs(x(5) - 0.5) > 0.000001 Or Abs(y(5) - 0.5) > 0.000001 Then Err.Raise 5, , "T6 smoothing did not improve interior"
    If x(1) <> 0# Or y(1) <> 0# Or x(3) <> 1# Or y(3) <> 1# Then Err.Raise 5, , "boundary changed"
    P5ValidateMeshQuality 13, 4, cn, x, y
    If Not P5QualityMeetsSettings() Then Err.Raise 5, , P5QualitySummary()
    QualityRegression = "PASS|T6 mean ratio " & CStr(before) & " -> " & CStr(after) & "|moved=" & CStr(moved) & "|" & P5QualitySummary()
    Exit Function
Failed: QualityRegression = "ERROR|" & Err.Description
End Function
Public Function QualityGenerate() As String
    On Error GoTo Failed
    SuppressUserMessages = True
    FEMUiCalculateSettings
    FEMUiBuildGeometry
    ボタン19_Click
    If Not FEMShapeGenerationOK Or Len(FEMUiMeshPreviewError) > 0 Then Err.Raise 5, , FEMUiMeshPreviewError
    QualityGenerate = "PASS|" & P5QualitySummary()
    Exit Function
Failed: QualityGenerate = "ERROR|" & Err.Description
End Function
Public Function QualityAnalyze() As String
    On Error GoTo Failed
    SuppressUserMessages = True
    P0_RunAnalysis
    QualityAnalyze = ResultStatus & "|" & AnalysisMessage
    Exit Function
Failed: QualityAnalyze = "ERROR|" & Err.Description
End Function
Public Function QualityPreflight() As String
    On Error GoTo Failed
    FEMWriteBuildStamp
    QualityPreflight = "PASS"
    Exit Function
Failed: QualityPreflight = "ERROR|" & Err.Source & "|" & Err.Description
End Function
Public Function QualityRefinementFixture() As String
    On Error GoTo Failed
    Dim ws As Worksheet, nodes(1 To 8, 1 To 14) As Variant, es(1 To 1, 1 To 11) As Variant, j As Long
    Set ws = ThisWorkbook.Worksheets("節点データ")
    ws.Range("A2:N60010").ClearContents
    nodes(2, 2) = 6#: nodes(3, 2) = 6#: nodes(3, 3) = 1#: nodes(4, 3) = 1#
    nodes(5, 2) = 3#: nodes(6, 2) = 6#: nodes(6, 3) = 0.5: nodes(7, 2) = 3#: nodes(7, 3) = 1#: nodes(8, 3) = 0.5
    For j = 1 To 8
        nodes(j, 1) = j
        If IsEmpty(nodes(j, 2)) Then nodes(j, 2) = 0#
        If IsEmpty(nodes(j, 3)) Then nodes(j, 3) = 0#
    Next j
    ws.Range("A2:N9").Value2 = nodes
    Set ws = ThisWorkbook.Worksheets("要素データ"): ws.Range("A2:K60010").ClearContents
    es(1, 1) = 1: es(1, 10) = 1: es(1, 11) = "Q8"
    For j = 1 To 8: es(1, j + 1) = j: Next j
    ws.Range("A2:K2").Value2 = es
    FEMWriteNumericSetting "MESH_QUALITY_MIN_ANGLE_DEG", 45#
    FEMWriteNumericSetting "MESH_QUALITY_MAX_ANGLE_DEG", 135#
    FEMWriteNumericSetting "MESH_GEN_COMPARE", 1#
    FEMWriteNumericSetting "MESH_REPAIR_TERRAIN_COMPARE", 0#
    FEMWriteNumericSetting "MESH_QUALITY_REFINE", 1#
    FEMWriteNumericSetting "MESH_QUALITY_REFINE_MAX_RATIO", 2#
    FEMWriteNumericSetting "MESH_QUALITY_REFINE_FRACTION", 1#
    FEMInvalidateRemeshedModel
    If Not MeshPrepareCandidate("SMOOTH") Then Err.Raise 5, , MeshAdaptLastError
    If InStr(MeshAdaptSelectedMethod(), "局所細分化") = 0 Then Err.Raise 5, , MeshAdaptLastSummary
    If Not MeshApplyCandidate(True) Then Err.Raise 5, , MeshAdaptLastError
    P5EvaluateCurrentSheet "局所細分化テスト"
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row <> 3 Or Not P5QualityMeetsSettings() Then Err.Raise 5, , P5QualitySummary()
    QualityRefinementFixture = "PASS|6:1 Q8 -> two 3:1 Q8|" & MeshAdaptLastSummary
    Exit Function
Failed: QualityRefinementFixture = "ERROR|" & Err.Description
End Function
Public Function QualityValidityAndDiagonal() As String
    On Error GoTo Failed
    Dim x() As Double, y() As Double, rows() As Variant, cn() As Long, bad() As Boolean, result As Variant
    Dim j As Long, a As Long, b As Long, count As Long, converted As Long, outCount As Long, detail As String, caught As Long
    ReDim x(1 To 8): ReDim y(1 To 8): ReDim rows(1 To 1, 1 To 10): ReDim cn(1 To 1, 1 To 8): ReDim bad(1 To 1)
    x(2) = 3#: x(3) = 3#: y(3) = 1#: x(4) = 1#: y(4) = 1#
    For j = 1 To 4
        a = j: b = j Mod 4 + 1: x(j + 4) = (x(a) + x(b)) / 2#: y(j + 4) = (y(a) + y(b)) / 2#
    Next j
    For j = 1 To 8: rows(1, j + 1) = j: cn(1, j) = j: Next j
    rows(1, 1) = 1: rows(1, 10) = 1: bad(1) = True
    FEMWriteNumericSetting "MESH_T6_MIN_ANGLE_DEG", 10#
    result = MeshAdaptConvertBadQ8ToT6(x, y, rows, bad, converted, outCount, detail)
    If converted <> 1 Or outCount <> 2 Or Abs(x(9) - 2#) > 0.000000001 Or Abs(y(9) - 0.5) > 0.000000001 Then Err.Raise 5, , "best diagonal not selected"
    ' A uniformly small/large valid Q8 must not be converted just for size.
    ReDim x(1 To 8): ReDim y(1 To 8)
    x(2) = 1#: x(3) = 1#: y(3) = 1#: y(4) = 1#
    For j = 1 To 4
        a = j: b = j Mod 4 + 1: x(j + 4) = (x(a) + x(b)) / 2#: y(j + 4) = (y(a) + y(b)) / 2#
    Next j
    FEMWriteNumericSetting "MESH_T6_FALLBACK", 1#
    FEMWriteNumericSetting "MESH_QUALITY_MIN_JACOBIAN_RATIO", 100#
    P5BuildT6FallbackMask 8, 1, cn, x, y, bad, count
    If count <> 0 Then Err.Raise 5, , "size warning triggered T6 conversion"
    ' Inversion is still a hard error after separating preferred targets.
    x(2) = -1#
    On Error Resume Next
    P5BuildT6FallbackMask 8, 1, cn, x, y, bad, count
    caught = Err.Number: Err.Clear
    On Error GoTo Failed
    If caught = 0 Then Err.Raise 5, , "invalid geometry was accepted"
    FEMWriteNumericSetting "MESH_QUALITY_MIN_JACOBIAN_RATIO", 0.000001
    FEMWriteNumericSetting "MESH_T6_MIN_ANGLE_DEG", 20#
    QualityValidityAndDiagonal = "PASS|best diagonal selected; size-only conversion avoided; invalid geometry rejected"
    Exit Function
Failed: QualityValidityAndDiagonal = "ERROR|" & Err.Description
End Function
Public Function QualityEdgeFlip() As String
    On Error GoTo Failed
    Dim x() As Double, y() As Double, rs() As Variant, cond As Variant, cn() As Long, j As Long, e As Long, count As Long
    ReDim x(1 To 9): ReDim y(1 To 9): ReDim rs(1 To 2, 1 To 11): ReDim cond(1 To 9, 1 To 14): ReDim cn(1 To 2, 1 To 8)
    x(2) = 3#: x(3) = 3#: y(3) = 1#: x(4) = 1#: y(4) = 1#
    x(5) = 1.5: x(6) = 3#: y(6) = 0.5: x(7) = 2#: y(7) = 1#: x(8) = 0.5: y(8) = 0.5: x(9) = 1.5: y(9) = 0.5
    Dim p As Variant
    p = Array(1, 2, 3, 3, 5, 6, 3, 9)
    For j = 1 To 8: rs(1, j + 1) = p(j - 1): Next j
    p = Array(1, 3, 4, 4, 9, 7, 4, 8)
    For j = 1 To 8: rs(2, j + 1) = p(j - 1): Next j
    For e = 1 To 2: rs(e, 1) = e: rs(e, 10) = 1: rs(e, 11) = "T6": Next e
    MeshFlipT6Edges x, y, rs, cond, 5, count
    If count <> 1 Or Abs(x(9) - 2#) > 0.000000001 Or y(9) <> 0.5 Then Err.Raise 5, , "neighbor T6 edge flip failed"
    For e = 1 To 2: For j = 1 To 8: cn(e, j) = CLng(rs(e, j + 1)): Next j: Next e
    P5ValidateMeshQuality 9, 2, cn, x, y
    QualityEdgeFlip = "PASS|neighbor T6 edge improved; boundary/midpoint conformity preserved"
    Exit Function
Failed: QualityEdgeFlip = "ERROR|" & Err.Description
End Function
