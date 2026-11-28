Option Explicit
' Prepared, conforming T6 mesh for the supplied twelve-region model.
' The snapshot is valid only for the exact geometry/spacing ledger saved with it.
Private Const CacheSheet As String = "SRM準備メッシュ"

Private Sub HashText(ByVal s As String, ByRef a As Double, ByRef b As Double)
    Dim i As Long, c As Long
    For i = 1 To Len(s)
        c = AscW(mid$(s, i, 1)): If c < 0 Then c = c + 65536
        a = a * 131# + c: a = a - Fix(a / 2147483629#) * 2147483629#
        b = b * 137# + c: b = b - Fix(b / 2147483587#) * 2147483587#
    Next i
End Sub

Public Function FEMPreparedShapeSignature() As String
    Dim ws As Worksheet, spec As Variant, data As Variant, r As Long, c As Long
    Dim a As Double, b As Double, last As Long, s As String, v As Variant
    Set ws = ThisWorkbook.Worksheets("形状入力")
    a = 17#: b = 31#
    For Each spec In Array(Array(2, 4), Array(8, 10))
        last = ws.cells(ws.rows.count, CLng(spec(0))).End(xlUp).row
        If last < 8 Then last = 8
        data = ws.range(ws.cells(8, CLng(spec(0))), ws.cells(last, CLng(spec(1)))).value2
        HashText CStr(spec(0)) & ":" & CStr(last), a, b
        For r = 1 To UBound(data, 1)
            For c = 1 To UBound(data, 2)
                v = data(r, c)
                If IsError(v) Then
                    s = "ERROR"
                ElseIf IsNumeric(v) And Not IsEmpty(v) Then
                    s = Format$(CDbl(v), "0.00000000000000000E+00")
                Else
                    s = CStr(v)
                End If
                HashText CStr(Len(s)) & ":" & s & ";", a, b
            Next c
        Next r
    Next spec
    FEMPreparedShapeSignature = CStr(a) & ":" & CStr(b)
End Function

Public Sub FEMPreparedAssertGeometry()
    Dim ws As Worksheet
    If FEMReadSetting("MESH_GEN_COMPARE", 1#) < 2.5 Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(CacheSheet)
    If FEMPreparedShapeSignature() <> CStr(ws.range("Y2").value2) Then
        Err.Raise vbObjectError + 7201, "FEMPreparedMesh", _
            "形状または格子間隔が準備済みメッシュと異なります。添付形状用メッシュのまま解析できません。変更形状に対応するメッシュを再作成してください。"
    End If
End Sub

Public Sub FEMPreparedValidateMesh()
    Dim ws As Worksheet, wn As Worksheet, we As Worksheet, nd As Variant, el As Variant
    Dim n As Long, e As Long, i As Long, j As Long, k As Long, mat As Long, slots As Long
    Dim p(1 To 4) As Long, area As Double, total(1 To 12) As Double, target As Double
    FEMPreparedAssertGeometry
    If FEMReadSetting("MESH_GEN_COMPARE", 1#) < 2.5 Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(CacheSheet)
    Set wn = ThisWorkbook.Worksheets("節点データ"): Set we = ThisWorkbook.Worksheets("要素データ")
    n = wn.cells(wn.rows.count, 1).End(xlUp).row - 1
    e = we.cells(we.rows.count, 1).End(xlUp).row - 1
    If n < 3 Or e < 1 Then GoTo BadMesh
    nd = wn.range("A2:C" & CStr(n + 1)).value2
    el = we.range("A2:K" & CStr(e + 1)).value2
    For i = 1 To e
        mat = CLng(el(i, 10)): If mat < 1 Or mat > 12 Then GoTo BadMesh
        slots = 4: If UCase$(Trim$(CStr(el(i, 11)))) = "T6" Then slots = 3
        For j = 1 To slots
            p(j) = CLng(el(i, j + 1))
            If p(j) < 1 Or p(j) > n Then GoTo BadMesh
            If CLng(nd(p(j), 1)) <> p(j) Then GoTo BadMesh
        Next j
        area = 0#
        For j = 1 To slots
            k = j Mod slots + 1
            area = area + CDbl(nd(p(j), 2)) * CDbl(nd(p(k), 3)) - CDbl(nd(p(k), 2)) * CDbl(nd(p(j), 3))
        Next j
        If area <= 0# Then GoTo BadMesh
        total(mat) = total(mat) + area / 2#
    Next i
    For mat = 1 To 12
        target = CDbl(ws.cells(mat + 5, 26).value2)
        If Abs(total(mat) - target) > 0.000001 * Application.Max(1#, target) Then GoTo BadMesh
    Next mat
    Exit Sub
BadMesh:
    Err.Raise vbObjectError + 7202, "FEMPreparedMesh", _
        "解析メッシュの材料別面積が添付の12領域と一致しません。モデル画面の「形状からメッシュ作成」で準備済みメッシュを復元してください。"
End Sub

Public Sub FEMPreparedRestoreMesh()
    Dim ws As Worksheet, wn As Worksheet, we As Worksheet, n As Long, e As Long, last As Long
    FEMPreparedAssertGeometry
    Set ws = ThisWorkbook.Worksheets(CacheSheet)
    Set wn = ThisWorkbook.Worksheets("節点データ"): Set we = ThisWorkbook.Worksheets("要素データ")
    n = CLng(ws.range("Y3").value2): e = CLng(ws.range("Y4").value2)
    If n < 3 Or e < 1 Then Err.Raise vbObjectError + 7203, , "準備済みメッシュが空です。"
    last = Application.Max(n + 1, wn.cells(wn.rows.count, 1).End(xlUp).row)
    wn.range("A2:N" & CStr(last)).ClearContents
    last = Application.Max(e + 1, we.cells(we.rows.count, 1).End(xlUp).row)
    we.range("A2:K" & CStr(last)).ClearContents
    wn.range("A2").Resize(n, 9).value2 = ws.range("A2").Resize(n, 9).value2
    we.range("A2").Resize(e, 11).value2 = ws.range("L2").Resize(e, 11).value2
    MeshApplySheetBoundaryConditions
    FEMInvalidateInputCache
    FEMUiInvalidateMeshPreview
    FEMViewerMarkDirty
    AnalysisOK = False: ResultRevision = 0: P1ResultReady = False
    FEMPreparedValidateMesh
    P5EvaluateCurrentSheet "添付12層・共有境界T6メッシュ"
    FEMWriteTextSetting "MESH_GEN_LAST_METHOD", "添付形状用・準備済みT6（形状照合あり）"
    FEMShapeGenerationOK = True: FEMUiMeshPreviewError = ""
End Sub
