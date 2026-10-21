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
  If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)): ws.Name = "表示量一覧"
  ws.Cells(2, 6).Value2 = "内容で選択する表示名"
  For i = 1 To 72
    ws.Cells(i + 2, 1).Value2 = i: ws.Cells(i + 2, 6).Value2 = FEMResultChoiceLabel(i)
  Next i
  On Error Resume Next
  ThisWorkbook.Names("FEMViewerSelectionLabels").Delete
  ThisWorkbook.Names("FEMViewerSelectionIDs").Delete
  On Error GoTo 0
  ThisWorkbook.Names.Add Name:="FEMViewerSelectionLabels", RefersTo:="='表示量一覧'!$F$3:$F$74"
  ThisWorkbook.Names.Add Name:="FEMViewerSelectionIDs", RefersTo:="='表示量一覧'!$A$3:$A$74"
  ws.Columns("F").ColumnWidth = 64
  ws.Range("A1").Value2 = "出力・表示の「表示する結果」で内容名を選びます。増分図には新しい版での再解析が必要です。"
End Sub

Public Sub FEMResultOptionsInstall()
  Dim ws As Worksheet, r As Long, selected As Long
  FEMResultOptionsEnsureList
  If Not FEMUiInstalled() Then Exit Sub
  Set ws = ThisWorkbook.Worksheets("出力・表示")
  For r = 26 To 34
    If Len(CStr(ws.Cells(r, 6).Value2)) > 0 Then Err.Raise vbObjectError + 3913, , "追加表示の配置範囲F26:F34に独自設定があります。設定を別の行へ移動してください。"
  Next r
  ws.Unprotect
  For r = 8 To ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    If CStr(ws.Cells(r, 6).Value2) = "VIEW_COLOR_RESULT" Then
      selected = FEMResultChoiceId(ws.Cells(r, 3).Value2)
      ws.Range(ws.Cells(r, 3), ws.Cells(r, 5)).UnMerge
      ws.Cells(r, 5).ClearContents
      ws.Range(ws.Cells(r, 3), ws.Cells(r, 5)).Merge
      ws.Cells(r, 2).Value2 = "表示する結果（内容で選択）"
      If selected > 0 Then ws.Cells(r, 3).Value2 = FEMResultChoiceLabel(selected)
      ws.Cells(r, 3).WrapText = True: ws.Cells(r, 3).MergeArea.Locked = False
      ws.Cells(r, 3).HorizontalAlignment = xlLeft
      ws.Cells(r, 7).Value2 = "内容名を選択。累積と増分を区別。"
      With ws.Cells(r, 3).Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=FEMViewerSelectionLabels"
        .IgnoreBlank = False: .InCellDropdown = True: .ShowError = True
        .InputTitle = "表示する物理量": .InputMessage = "内容名から選択してください。旧番号1～72も内部読込では対応します。"
      End With
    End If
  Next r
  ws.Range("B26:G26").Merge: ws.Range("B26").Value2 = "せん断帯の追加表示（チェックは個別にON/OFF）"
  ws.Range("B27").Value2 = "増分の表示方法"
  ws.Range("C27:E27").Merge
  If Len(CStr(ws.Range("C27").Value2)) = 0 Then ws.Range("C27").Value2 = "A：選択結果へ等値線を追加"
  With ws.Range("C27").Validation
    .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="A：選択結果へ等値線を追加,B：増分専用カラー"
    .InCellDropdown = True: .ShowError = True
  End With
  ws.Range("C27").MergeArea.Locked = False
  ws.Range("B28:G29").UnMerge: ws.Range("B28:G29").ClearContents
  FEMOverlayCheck ws, "UI_ShearIncrement", "最大せん断ひずみ増分を追加", "B28:G28", "J28"
  FEMOverlayCheck ws, "UI_PlasticIncrement", "塑性主ひずみ増分クロスを追加", "B29:G29", "J29"
  ws.Columns("J").Hidden = True
  ws.Range("B30").Value2 = "クロスの表示倍率"
  If Len(CStr(ws.Range("C30").Value2)) = 0 Then ws.Range("C30").Value2 = 1#
  With ws.Range("C30").Validation
    .Delete: .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=AND(ISNUMBER(C30),C30>0,C30<=10)"
    .ShowError = True
  End With
  ws.Range("C30").Locked = False: ws.Range("C30").Interior.Color = RGB(255, 242, 204)
  ws.Range("E30:G30").Merge: ws.Range("E30").Value2 = "赤=引張（外向き）/ 青=圧縮（内向き）/ 長さは共通尺度": ws.Range("E30").WrapText = True
  ws.Range("B32:G34").Merge: ws.Range("B32").Value2 = "対象は最終収束荷重増分です。SRMでは最後の成功Fs。等値線は要素の4点最大値を節点平均して補間、クロスは要素平均塑性テンソルです。旧結果に増分がない場合は最初から再解析してください。": ws.Range("B32").WrapText = True
  ws.Rows("26:34").RowHeight = 25
  With ws.Shapes("UI_ResultFigure")
    .Left = ws.Range("B37:C38").Left + 2: .Top = ws.Range("B37:C38").Top + 2
    .Width = ws.Range("B37:C38").Width - 4: .Height = ws.Range("B37:C38").Height - 4
  End With
  FEMUiBindSettings
  ws.Protect UserInterfaceOnly:=True, DrawingObjects:=False
End Sub

Private Sub FEMOverlayCheck(ByVal ws As Worksheet, ByVal objectName As String, ByVal caption As String, ByVal address As String, ByVal link As String)
  Dim check As Object, target As Range, saved As Boolean
  If Not IsEmpty(ws.Range(link).Value2) Then saved = CBool(ws.Range(link).Value2)
  On Error Resume Next: Set check = ws.CheckBoxes(objectName): On Error GoTo 0
  Set target = ws.Range(address)
  ' Forms checkbox fonts cannot be changed in Excel. Use a readable cell label beside the box.
  target.Merge: target.Cells(1, 1).Value2 = caption
  target.Font.Size = 11: target.IndentLevel = 3
  target.HorizontalAlignment = xlLeft: target.VerticalAlignment = xlCenter
  If check Is Nothing Then Set check = ws.CheckBoxes.Add(target.Left + 4, target.Top + 2, 18, target.Height - 4)
  check.Name = objectName: check.Caption = ""
  check.Left = target.Left + 4: check.Top = target.Top + 2: check.Width = 18: check.Height = target.Height - 4
  check.LinkedCell = "'" & ws.Name & "'!" & link
  check.Value = IIf(saved, xlOn, xlOff): check.Placement = xlMoveAndSize
  ws.Range(link).Locked = False
End Sub

Public Function FEMShearOverlayOn() As Boolean
  FEMShearOverlayOn = CBool(ThisWorkbook.Worksheets("出力・表示").Range("J28").Value2)
End Function

Public Function FEMPlasticCrossOn() As Boolean
  FEMPlasticCrossOn = CBool(ThisWorkbook.Worksheets("出力・表示").Range("J29").Value2)
End Function

Public Function FEMOverlayModeA() As Boolean
  Dim value As String
  value = CStr(ThisWorkbook.Worksheets("出力・表示").Range("C27").Value2)
  If value <> "A：選択結果へ等値線を追加" And value <> "B：増分専用カラー" Then Err.Raise vbObjectError + 3910, , "増分の表示方法はAまたはBを選択してください。"
  FEMOverlayModeA = (Left$(value, 1) = "A")
End Function

Public Function FEMPlasticCrossScale() As Double
  Dim value As Variant
  value = ThisWorkbook.Worksheets("出力・表示").Range("C30").Value2
  If Not IsNumeric(value) Then Err.Raise vbObjectError + 3910, , "クロス倍率は0より大きく10以下の数値を指定してください。"
  If CDbl(value) <= 0# Or CDbl(value) > 10# Then Err.Raise vbObjectError + 3910, , "クロス倍率は0より大きく10以下で指定してください。"
  FEMPlasticCrossScale = CDbl(value)
End Function
