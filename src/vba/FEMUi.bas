Option Explicit

' Settings bridge for the purpose-specific UI sheets.  Numerical code continues
' to read the legacy 設定 sheet through its existing E/C registry.
Private Const FEM_UI_PANEL As String = "操作パネル"
Private Function FEMUiPages() As Variant
    FEMUiPages = Array("メッシュ設定", "解析設定", "高速化設定", "出力・表示", "詳細設定")
End Function

Public Function FEMUiInstalled() As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FEM_UI_PANEL)
    On Error GoTo 0
    FEMUiInstalled = Not ws Is Nothing
End Function

Public Function FEMUiIsSettingsSheet(ByVal sheetName As String) As Boolean
    Dim p As Variant, pages As Variant
    If Not FEMUiInstalled() Then Exit Function
    pages = FEMUiPages()
    For Each p In pages
        If StrComp(CStr(p), sheetName, vbBinaryCompare) = 0 Then FEMUiIsSettingsSheet = True: Exit Function
    Next p
End Function

Public Sub FEMUiCalculateSettings()
    Dim ws As Worksheet
    If Not FEMUiInstalled() Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("設定")
    On Error GoTo 0
    If Not ws Is Nothing Then ws.Calculate
End Sub

Private Function FEMUiFindCanonical(ByVal keyName As String, ByRef page As Worksheet, ByRef rowNo As Long) As Boolean
    Dim p As Variant, ws As Worksheet, lastRow As Long, r As Long, pages As Variant
    pages = FEMUiPages()
    For Each p In pages
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(p))
        On Error GoTo 0
        If Not ws Is Nothing Then
            lastRow = ws.Cells(ws.rows.count, 6).End(xlUp).row
            For r = 1 To lastRow
                If StrComp(Trim$(CStr(ws.Cells(r, 6).value2)), Trim$(keyName), vbTextCompare) = 0 Then
                    Set page = ws: rowNo = r: FEMUiFindCanonical = True: Exit Function
                End If
            Next r
        End If
    Next p
End Function

Public Function FEMUiTryWriteSetting(ByVal keyName As String, ByVal value As Variant) As Boolean
    Dim ws As Worksheet, rowNo As Long, oldEvents As Boolean
    If Not FEMUiInstalled() Then Exit Function
    If Not FEMUiFindCanonical(keyName, ws, rowNo) Then Exit Function
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    ws.Cells(rowNo, 3).value2 = value
    FEMUiCalculateSettings
    Application.EnableEvents = oldEvents
    FEMUiTryWriteSetting = True
    Exit Function
Failed:
    Application.EnableEvents = oldEvents
    Err.Raise Err.Number, Err.source, Err.Description
End Function

Private Function FEMUiFindLegacyRow(ByVal ws As Worksheet, ByVal keyName As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.rows.count, 5).End(xlUp).row
    For r = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 5).value2)), Trim$(keyName), vbTextCompare) = 0 Then FEMUiFindLegacyRow = r: Exit Function
    Next r
End Function

Public Sub FEMUiBindSettings()
    Dim pages As Variant, p As Variant, ws As Worksheet, legacy As Worksheet
    Dim seen As Object, formulas As Object, key As String, lastRow As Long, r As Long, lr As Long
    Dim oldEvents As Boolean, addr As String
    If Not FEMUiInstalled() Then Exit Sub
    Set seen = CreateObject("Scripting.Dictionary")
    Set formulas = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    formulas.CompareMode = vbTextCompare
    Set legacy = ThisWorkbook.Worksheets("設定")
    pages = FEMUiPages()
    For Each p In pages
        Set ws = ThisWorkbook.Worksheets(CStr(p))
        lastRow = ws.Cells(ws.rows.count, 6).End(xlUp).row
        For r = 1 To lastRow
            key = Trim$(CStr(ws.Cells(r, 6).value2))
            If Len(key) > 0 Then
                If seen.Exists(key) Then Err.Raise vbObjectError + 3801, "FEMUiBindSettings", "重複した設定キー: " & key
                seen.Add key, True
                lr = FEMUiFindLegacyRow(legacy, key)
                If lr = 0 Then Err.Raise vbObjectError + 3802, "FEMUiBindSettings", "旧設定キーがありません: " & key
                addr = "=IF('" & Replace(CStr(p), "'", "''") & "'!C" & CStr(r) & "="""","""",'" & Replace(CStr(p), "'", "''") & "'!C" & CStr(r) & ")"
                formulas.Add CStr(lr), addr
            End If
        Next r
    Next p
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Dim k As Variant
    For Each k In formulas.Keys
        ' Text-formatted cells would keep the formula as literal text.
        legacy.Cells(CLng(k), 3).NumberFormat = "General"
        legacy.Cells(CLng(k), 3).formula = formulas(k)
    Next k
    FEMUiCalculateSettings
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    Application.EnableEvents = oldEvents
    Err.Raise Err.Number, Err.source, Err.Description
End Sub

Public Sub FEMUiSettingsChanged(ByVal Sh As Object, ByVal target As range)
    Dim hit As range, cell As range, keyName As String, invalidateAnalysis As Boolean
    If Not FEMUiInstalled() Then Exit Sub
    If Not FEMUiIsSettingsSheet(Sh.name) Then Exit Sub
    Set hit = Application.Intersect(target, Union(Sh.columns(3), Sh.columns(6)))
    If hit Is Nothing Then Exit Sub
    For Each cell In hit.Cells
        If cell.column = 6 Then
            invalidateAnalysis = True
        ElseIf cell.column = 3 Then
            keyName = UCase$(Trim$(CStr(Sh.Cells(cell.row, 6).value2)))
            If Len(keyName) > 0 Then
                If Not FEMUiIsNonAnalysisKey(keyName) Then invalidateAnalysis = True
            End If
        End If
    Next cell
    On Error Resume Next
    FEMInvalidateSettingCache
    FEMViewerMarkDirty
    If invalidateAnalysis Then
        AnalysisOK = False
        ResultRevision = 0
        P1ResultReady = False
    End If
    On Error GoTo 0
End Sub

Private Function FEMUiIsNonAnalysisKey(ByVal keyName As String) As Boolean
    Dim k As String
    k = UCase$(Trim$(keyName))
    FEMUiIsNonAnalysisKey = (left$(k, 5) = "VIEW_" Or left$(k, 6) = "DEBUG_" Or _
        left$(k, 7) = "EXPORT_" Or left$(k, 4) = "RCM_" Or left$(k, 6) = "ADAPT_" Or _
        left$(k, 12) = "MESH_REPAIR_" Or left$(k, 18) = "MESH_QUALITY_LAST_" Or _
        left$(k, 20) = "MESH_QUALITY_REPAIR_")
End Function










