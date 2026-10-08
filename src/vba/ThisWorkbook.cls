
Private Sub Workbook_Open()
  On Error Resume Next
  FEMWriteBuildStamp False
  FEMUiShowHome
  Err.Clear
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As range)
  Dim sheetName As String
  Dim cell As range
  Dim keyName As String
  Dim invalidateAnalysis As Boolean
  Dim invalidateSettings As Boolean
  Dim rngC As range
  Dim rngE As range
  On Error Resume Next
  If AnalysisRunning Then Exit Sub
  sheetName = Sh.name
  If FEMUiIsSettingsSheet(sheetName) Then
    FEMUiSettingsChanged Sh, target
    Exit Sub
  End If
  If sheetName = "節点データ" Or sheetName = "材料データ" Or sheetName = "要素データ" _
     Or sheetName = "形状入力" Or sheetName = "ステージ" Or sheetName = "載荷" Or sheetName = "図形定義" Or sheetName = "接合" Then
    FEMViewerMarkDirty
    AnalysisOK = False
    ResultRevision = 0
    P1ResultReady = False
    Exit Sub
  End If
  If sheetName <> "設定" Then Exit Sub
  Set rngC = Application.Intersect(target, Sh.columns(3))
  Set rngE = Application.Intersect(target, Sh.columns(5))
  If rngC Is Nothing And rngE Is Nothing Then Exit Sub
  invalidateAnalysis = False
  invalidateSettings = False
  If Not rngE Is Nothing Then
    For Each cell In rngE.Cells
      If cell.row >= 4 Then
        invalidateSettings = True
        invalidateAnalysis = True
        Exit For
      End If
    Next cell
  End If
  If Not rngC Is Nothing Then
    For Each cell In rngC.Cells
      keyName = UCase$(Trim$(CStr(Sh.Cells(cell.row, 5).value2)))
      If Len(keyName) > 0 And keyName <> "設定キー" Then
        invalidateSettings = True
        If left$(keyName, 6) <> "ADAPT_" And left$(keyName, 4) <> "RCM_" And left$(keyName, 12) <> "MESH_REPAIR_" _
           And left$(keyName, 5) <> "VIEW_" And left$(keyName, 6) <> "DEBUG_" And left$(keyName, 7) <> "EXPORT_" _
           And left$(keyName, 18) <> "MESH_QUALITY_LAST_" _
           And left$(keyName, 20) <> "MESH_QUALITY_REPAIR_" Then
          invalidateAnalysis = True
        End If
      End If
    Next cell
  End If
  If invalidateSettings Then FEMInvalidateSettingCache
  If invalidateAnalysis Then
    FEMViewerMarkDirty
    AnalysisOK = False
    ResultRevision = 0
    P1ResultReady = False
  End If
End Sub

