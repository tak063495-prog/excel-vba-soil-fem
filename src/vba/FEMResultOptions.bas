Option Explicit

Public Function FEMResultChoiceLabel(ByVal item As Long) As String
  Dim names As Variant, gp As Long, q As Long
  names = Array("X方向応力", "Y方向応力", "XYせん断応力", "最大主応力", "最小主応力", "最大せん断応力", "主応力角", "相当応力", "偏差応力", "降伏関数", "X方向ひずみ（累積）", "Y方向ひずみ（累積）", "XYせん断ひずみ（累積）", "8面せん断ひずみ（累積）", "体積ひずみ（累積）", "面内最大せん断ひずみ（累積）", "Z方向応力")
  If item >= 1 And item <= 68 Then
    gp = (item - 1) \ 17 + 1: q = (item - 1) Mod 17
    FEMResultChoiceLabel = CStr(names(q)) & "（積分点" & CStr(gp) & "）"
  ElseIf item = 69 Then
    FEMResultChoiceLabel = "最大せん断ひずみ増分（最終収束増分・4点最大）"
  ElseIf item = 70 Then
    FEMResultChoiceLabel = "最大塑性せん断ひずみ増分（最終収束増分・4点最大）"
  ElseIf item = 71 Then
    FEMResultChoiceLabel = "塑性状態（弾性／現在降伏／過去に降伏・要素の4点）"
  ElseIf item = 72 Then
    FEMResultChoiceLabel = "累積塑性せん断ひずみ（4点最大）"
  End If
End Function

Public Function FEMResultChoiceId(ByVal value As Variant) As Long
  Dim i As Long, x As Double
  If IsError(value) Or IsEmpty(value) Then Exit Function
  If IsNumeric(value) Then
    x = CDbl(value)
    If x >= 1# And x <= 72# And x = Fix(x) Then FEMResultChoiceId = CLng(x)
  Else
    For i = 1 To 72
      If CStr(value) = FEMResultChoiceLabel(i) Then FEMResultChoiceId = i: Exit Function
    Next i
  End If
End Function

Public Sub FEMResultOptionsEnsureList()
  Dim ws As Worksheet, i As Long
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("表示量一覧"): On Error GoTo 0
  If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = "表示量一覧"
  ws.cells(2, 6).value2 = "内容で選択する表示名"
  For i = 1 To 72
    ws.cells(i + 2, 1).value2 = i: ws.cells(i + 2, 6).value2 = FEMResultChoiceLabel(i)
  Next i
  On Error Resume Next
  ThisWorkbook.names("FEMViewerSelectionLabels").Delete
  ThisWorkbook.names("FEMViewerSelectionIDs").Delete
  On Error GoTo 0
  ThisWorkbook.names.Add name:="FEMViewerSelectionLabels", RefersTo:="='表示量一覧'!$F$3:$F$74"
  ThisWorkbook.names.Add name:="FEMViewerSelectionIDs", RefersTo:="='表示量一覧'!$A$3:$A$74"
  ws.columns("F").ColumnWidth = 64
  ws.range("A1").value2 = "出力・表示の「表示する結果」で内容名を選びます。増分図には新しい版での再解析が必要です。"
End Sub

Public Sub FEMResultOptionsInstall()
  Dim ws As Worksheet, r As Long, selected As Long
  FEMResultOptionsEnsureList
  If Not FEMUiInstalled() Then Exit Sub
  Set ws = ThisWorkbook.Worksheets("出力・表示")
  For r = 26 To 34
    If Len(CStr(ws.cells(r, 6).value2)) > 0 Then Err.Raise vbObjectError + 3913, , "追加表示の配置範囲F26:F34に独自設定があります。設定を別の行へ移動してください。"
  Next r
  ws.Unprotect
  For r = 8 To ws.cells(ws.rows.count, 6).End(xlUp).row
    If CStr(ws.cells(r, 6).value2) = "VIEW_COLOR_RESULT" Then
      selected = FEMResultChoiceId(ws.cells(r, 3).value2)
      ws.range(ws.cells(r, 3), ws.cells(r, 5)).UnMerge
      ws.cells(r, 5).ClearContents
      ws.range(ws.cells(r, 3), ws.cells(r, 5)).Merge
      ws.cells(r, 2).value2 = "表示する結果（内容で選択）"
      If selected > 0 Then ws.cells(r, 3).value2 = FEMResultChoiceLabel(selected)
      ws.cells(r, 3).WrapText = True: ws.cells(r, 3).MergeArea.locked = False
      ws.cells(r, 3).HorizontalAlignment = xlLeft
      ws.cells(r, 7).value2 = "内容名を選択。累積と増分を区別。"
      With ws.cells(r, 3).Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=FEMViewerSelectionLabels"
        .IgnoreBlank = False: .InCellDropdown = True: .ShowError = True
        .InputTitle = "表示する物理量": .InputMessage = "内容名から選択してください。旧番号1～72も内部読込では対応します。"
      End With
    End If
  Next r
  ws.range("B26:G26").Merge: ws.range("B26").value2 = "せん断帯の追加表示（チェックは個別にON/OFF）"
  ws.range("B27").value2 = "増分の表示方法"
  ws.range("C27:E27").Merge
  If Len(CStr(ws.range("C27").value2)) = 0 Then ws.range("C27").value2 = "A：選択結果へ等値線を追加"
  With ws.range("C27").Validation
    .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="A：選択結果へ等値線を追加,B：増分専用カラー"
    .InCellDropdown = True: .ShowError = True
  End With
  ws.range("C27").MergeArea.locked = False
  ws.range("B28:G29").UnMerge: ws.range("B28:G29").ClearContents
  FEMOverlayCheck ws, "UI_ShearIncrement", "最大せん断ひずみ増分を追加", "B28:G28", "J28"
  FEMOverlayCheck ws, "UI_PlasticIncrement", "塑性主ひずみ増分クロスを追加", "B29:G29", "J29"
  ws.columns("J").Hidden = True
  ws.range("B30").value2 = "クロスの表示倍率"
  If Len(CStr(ws.range("C30").value2)) = 0 Then ws.range("C30").value2 = 1#
  With ws.range("C30").Validation
    .Delete: .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=AND(ISNUMBER(C30),C30>0,C30<=10)"
    .ShowError = True
  End With
  ws.range("C30").locked = False: ws.range("C30").Interior.Color = RGB(255, 242, 204)
  ws.range("E30:G30").Merge: ws.range("E30").value2 = "赤=引張（外向き）/ 青=圧縮（内向き）/ 長さは共通尺度": ws.range("E30").WrapText = True
  ws.range("B32:G34").Merge: ws.range("B32").value2 = "対象は最終収束荷重増分です。SRMでは最後の成功Fs。等値線は要素の4点最大値を節点平均して補間、クロスは要素平均塑性テンソルです。旧結果に増分がない場合は最初から再解析してください。": ws.range("B32").WrapText = True
  ws.rows("26:34").RowHeight = 25
  With ws.Shapes("UI_ResultFigure")
    .left = ws.range("B37:C38").left + 2: .Top = ws.range("B37:C38").Top + 2
    .width = ws.range("B37:C38").width - 4: .Height = ws.range("B37:C38").Height - 4
  End With
  FEMUiBindSettings
  ws.Protect UserInterfaceOnly:=True, DrawingObjects:=False
End Sub

Private Sub FEMOverlayCheck(ByVal ws As Worksheet, ByVal objectName As String, ByVal caption As String, ByVal address As String, ByVal link As String)
  Dim check As Object, target As range, saved As Boolean
  If Not IsEmpty(ws.range(link).value2) Then saved = CBool(ws.range(link).value2)
  On Error Resume Next: Set check = ws.CheckBoxes(objectName): On Error GoTo 0
  Set target = ws.range(address)
  ' Forms checkbox fonts cannot be changed in Excel. Use a readable cell label beside the box.
  target.Merge: target.cells(1, 1).value2 = caption
  target.Font.size = 11: target.IndentLevel = 3
  target.HorizontalAlignment = xlLeft: target.VerticalAlignment = xlCenter
  If check Is Nothing Then Set check = ws.CheckBoxes.Add(target.left + 4, target.Top + 2, 18, target.Height - 4)
  check.name = objectName: check.caption = ""
  check.left = target.left + 4: check.Top = target.Top + 2: check.width = 18: check.Height = target.Height - 4
  check.LinkedCell = "'" & ws.name & "'!" & link
  check.value = IIf(saved, xlOn, xlOff): check.Placement = xlMoveAndSize
  ws.range(link).locked = False
End Sub

Public Function FEMShearOverlayOn() As Boolean
  FEMShearOverlayOn = CBool(ThisWorkbook.Worksheets("出力・表示").range("J28").value2)
End Function

Public Function FEMPlasticCrossOn() As Boolean
  FEMPlasticCrossOn = CBool(ThisWorkbook.Worksheets("出力・表示").range("J29").value2)
End Function

Public Function FEMOverlayModeA() As Boolean
  Dim value As String
  value = CStr(ThisWorkbook.Worksheets("出力・表示").range("C27").value2)
  If value <> "A：選択結果へ等値線を追加" And value <> "B：増分専用カラー" Then Err.Raise vbObjectError + 3910, , "増分の表示方法はAまたはBを選択してください。"
  FEMOverlayModeA = (left$(value, 1) = "A")
End Function

Public Function FEMPlasticCrossScale() As Double
  Dim value As Variant
  value = ThisWorkbook.Worksheets("出力・表示").range("C30").value2
  If Not IsNumeric(value) Then Err.Raise vbObjectError + 3910, , "クロス倍率は0より大きく10以下の数値を指定してください。"
  If CDbl(value) <= 0# Or CDbl(value) > 10# Then Err.Raise vbObjectError + 3910, , "クロス倍率は0より大きく10以下で指定してください。"
  FEMPlasticCrossScale = CDbl(value)
End Function
