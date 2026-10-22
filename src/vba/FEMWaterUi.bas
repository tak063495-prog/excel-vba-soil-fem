Option Explicit

Public Sub FEMWaterInstallUi()
  Dim ws As Worksheet, r As Long
  Set ws = FEMWaterSheet("水条件", True)
  If Len(CStr(ws.range("A1").value2)) = 0 Then ws.range("A1").value2 = "水条件・定常浸透"
  ws.range("A3").value2 = "モード": If Len(CStr(ws.range("B3").value2)) = 0 Then ws.range("B3").value2 = "NONE"
  ws.range("A4").value2 = "水の単位体積重量 γw [kN/m3]": If Len(CStr(ws.range("B4").value2)) = 0 Then ws.range("B4").value2 = 9.80665
  ws.range("A5").value2 = "外水位線ID": If Len(CStr(ws.range("B5").value2)) = 0 Then ws.range("B5").value2 = "外水"
  ws.range("A6").value2 = "外水圧AUTO": If Len(CStr(ws.range("B6").value2)) = 0 Then ws.range("B6").value2 = "AUTO"
  ws.range("A7").value2 = "入力確認署名": ws.range("A8").value2 = "単位：座標・水位・全水頭はm、Kはm/s。飽和水の流れと外水圧を確認します。既存の自重（全応力）は変更しません。"
  ws.rows("3:7").RowHeight = 24
  FEMWaterHeaders ws
  FEMWaterValidation ws
  FEMWaterButton ws, "GW_UI_INPUT", "入力へ", "FEMWaterShowInputs", 8, 3
  FEMWaterButton ws, "GW_UI_PREVIEW", "水圧プレビュー", "FEMWaterPreview", 8, 5
  FEMWaterButton ws, "GW_UI_RESULTS", "水圧結果", "FEMWaterResults", 8, 7
  ws.columns("A:J").ColumnWidth = 18
  ws.columns("A").ColumnWidth = 30
  ws.range("A8:J9").Merge
  ws.range("A8:J9").WrapText = True
  ws.rows("8:9").RowHeight = 24
  ws.rows(11).RowHeight = 54: ws.rows(215).RowHeight = 54
  ws.range("A11:J11").WrapText = True: ws.range("A215:F215").WrapText = True
  ws.range("A11:J11").Interior.Color = RGB(218, 232, 246)
  ws.range("A215:F215").Interior.Color = RGB(218, 232, 246)
  ws.range("B3:B6").Interior.Color = RGB(255, 249, 210)
  FEMWaterLink "解析設定", "水条件", "A3", "水条件"
  FEMWaterLink "出力・表示", "水条件", "A3", "水圧入力"
End Sub

Private Sub FEMWaterButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, ByVal macroName As String, ByVal col As Long, ByVal row As Long)
  Dim shp As shape
  On Error Resume Next: ws.Shapes(shapeName).Delete: On Error GoTo 0
  Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.cells(row, col).left, ws.cells(row, col).Top, 125, 22)
  shp.name = shapeName: shp.TextFrame2.TextRange.text = caption: shp.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & macroName
End Sub

Public Sub FEMWaterShowInputs()
  FEMWaterInstallUi
  FEMWaterSheet("水条件", False).Activate
End Sub

Public Sub FEMWaterPreview()
  On Error GoTo Failed
  FEMWaterInstallUi
  FEMWaterSheet("水条件", False).range("B7").ClearContents
  GWPreviewCompute
  FEMWaterSheet("水条件", False).range("B7").value2 = GWInputSignature()
  FEMWaterSheet("水圧確認", False).Activate
  Exit Sub
Failed:
  If SuppressUserMessages Then Err.Raise Err.number, "FEMWaterPreview", Err.description
  MsgBox "水圧プレビューに失敗しました：" & Err.description, vbExclamation, "水条件"
End Sub

Public Sub FEMWaterResults()
  Dim ws As Worksheet
  On Error GoTo MissingResult
  If Not GWEnabled() Then Err.Raise vbObjectError + 7912, , "水条件が無効です。"
  Set ws = FEMWaterSheet("水圧結果", False)
  If CStr(ws.range("R5").value2) <> GWInputSignature() Then Err.Raise vbObjectError + 7912, , "入力が変更されています。再解析してください。"
  If ThisWorkbook.names("_FEMViewerResultsDirty").RefersTo = "=TRUE" Then Err.Raise vbObjectError + 7912, , "計算結果が古くなっています。再解析してください。"
  ws.Activate
  Exit Sub
MissingResult:
  If SuppressUserMessages Then Err.Raise Err.number, "FEMWaterResults", Err.description
  MsgBox "受理した水圧結果を表示できません。" & Err.description, vbExclamation
End Sub

Private Function FEMWaterSheet(ByVal sheetName As String, ByVal createIfMissing As Boolean) As Worksheet
  On Error Resume Next: Set FEMWaterSheet = ThisWorkbook.Worksheets(sheetName): On Error GoTo 0
  If FEMWaterSheet Is Nothing And createIfMissing Then
    Set FEMWaterSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    FEMWaterSheet.name = sheetName
  End If
  If FEMWaterSheet Is Nothing Then Err.Raise vbObjectError + 7911, "FEMWaterSheet", "シートがありません：" & sheetName
End Function

Private Sub FEMWaterHeaders(ByVal ws As Worksheet)
  Dim c As Long
  ws.range("A11:D11").value2 = Array("水位線ID", "X[m]", "水位標高[m]", "説明")
  ws.range("F11:J11").value2 = Array("材料番号", "応力基準(TOTAL/EFFECTIVE)", "透水係数K[m/s]", "水位線ID", "説明")
  ws.range("A215:F215").value2 = Array("要素番号", "辺番号(1..4 CCW)", "種別(HEAD/NOFLOW/EXPOSED/DRY)", "水位線ID", "全水頭[m]（任意）", "説明")
  For c = 1 To 10: ws.cells(11, c).Font.bOld = True: Next c
  For c = 1 To 6: ws.cells(215, c).Font.bOld = True: Next c
End Sub

Private Sub FEMWaterValidation(ByVal ws As Worksheet)
  On Error Resume Next
  ws.range("B3").Validation.Delete: ws.range("B3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="NONE,HYDROSTATIC,STEADY"
  ws.range("B6").Validation.Delete: ws.range("B6").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="AUTO,OFF"
  ws.range("G12:G211").Validation.Delete: ws.range("G12:G211").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="TOTAL,EFFECTIVE"
  ws.range("C216:C415").Validation.Delete: ws.range("C216:C415").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="HEAD,NOFLOW,EXPOSED,DRY"
  On Error GoTo 0
End Sub

Private Sub FEMWaterLink(ByVal destination As String, ByVal source As String, ByVal cellAddress As String, ByVal caption As String)
  Dim ws As Worksheet, r As Long, anchor As range, wasProtected As Boolean
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets(destination): On Error GoTo 0
  If ws Is Nothing Then Exit Sub
  For r = 1 To ws.UsedRange.row + ws.UsedRange.rows.count
    If CStr(ws.cells(r, 1).value2) = caption Then Exit Sub
  Next r
  wasProtected = ws.ProtectContents
  If wasProtected Then ws.Unprotect
  r = ws.cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row + 3
  Set anchor = ws.cells(r, 1)
  On Error Resume Next: anchor.Hyperlinks.Delete: On Error GoTo 0
  ws.Hyperlinks.Add anchor:=anchor, address:="", SubAddress:="'" & source & "'!" & cellAddress, TextToDisplay:=caption
  If wasProtected Then ws.Protect UserInterfaceOnly:=True, DrawingObjects:=False
End Sub
