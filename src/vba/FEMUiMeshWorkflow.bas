Option Explicit

Public FEMShapeGenerationOK As Boolean
Public FEMUiMeshPreviewError As String
Public FEMUiMeshQualityNote As String
Private previewReady As Boolean
Private previewSignature As String

Public Function FEMUiRetiredMeshKey(ByVal key As String) As Boolean
  FEMUiRetiredMeshKey = (left$(key, 10) = "MESH_POINT" Or key = "MESH_NX" Or key = "MESH_NY")
End Function

Public Sub FEMUiInvalidateMeshPreview()
  previewReady = False: previewSignature = "": FEMUiMeshQualityNote = ""
End Sub

' Keep old simple-mesh values in hidden rows for file compatibility only.
Public Sub FEMUiShapeOnlyPage()
  Dim ws As Worksheet, item As Variant, r As Long
  Set ws = ThisWorkbook.Worksheets("メッシュ設定")
  For r = 35 To 49
    If Len(CStr(ws.cells(r, 6).value2)) > 0 Then Err.Raise vbObjectError + 3896, , "メッシュ設定の操作領域にカスタムKEYがあります。値は変更していません。行=" & CStr(r)
  Next r
  ws.rows("16:27").Hidden = True
  ws.range("B2").value2 = "メッシュ設定（形状メッシュ）"
  ws.range("B5").value2 = "材料・外周拘束を確認 → 形状入力 → 作成 → メッシュ・拘束・品質を確認します。"
  ws.range("B28").value2 = "形状メッシュ：座標・境界・格子間隔は「形状入力」で指定"
  ws.range("E9").value2 = "形状メッシュの生成要素に付ける材料番号"
  ws.range("E29").value2 = "0=品質比較。1=デローニ優先。2=地形格子を直接生成（適用不可は停止）。"
  ws.range("B31:E31").UnMerge: ws.range("B31:E31").Merge
  ws.range("B31").value2 = "作成すると節点・要素を更新します。既存モデルを使う場合は作成せず、確認表示へ進んでください。"
  ws.range("B31").WrapText = True: ws.rows(31).RowHeight = 34
  ws.range("B35:G49").UnMerge: ws.range("B35:G49").ClearContents
  ws.rows("35:36").RowHeight = 22: ws.rows("40:41").RowHeight = 22
  ws.range("B38:G38").Merge
  ws.range("B38").value2 = "作成後は図を自動表示します。X/Yの拘束方向と品質を確認し、必要なら品質修正候補を作ります。"
  ws.range("B38").WrapText = True: ws.rows(38).RowHeight = 36
  ws.range("B43:G43").Merge
  ws.range("B43").value2 = "候補は採用するまで元のメッシュを変更しません。採用後の図も確認してから解析します。"
  ws.range("B43").WrapText = True: ws.rows(43).RowHeight = 36
  On Error Resume Next: ws.Shapes("UI_Mesh1").Delete: On Error GoTo 0
  ' Remove duplicate cell captions underneath action buttons; explanations stay outside.
  Set ws = ThisWorkbook.Worksheets("操作パネル")
  For Each item In Array("B18:D19", "B31:C31", "D31", "F31:G31", "H31", "B33:C33", "D33")
    ws.range(CStr(item)).ClearContents
  Next item
  ws.range("B18:D19").ClearContents
  ws.PageSetup.PrintArea = "$B$2:$H$44"
End Sub

Private Sub FingerprintText(ByVal s As String, ByRef a As Double, ByRef b As Double)
  Dim j As Long, c As Long
  For j = 1 To Len(s)
    c = AscW(mid$(s, j, 1)): If c < 0 Then c = c + 65536
    a = a * 131# + c: a = a - Fix(a / 2147483629#) * 2147483629#
    b = b * 137# + c: b = b - Fix(b / 2147483587#) * 2147483587#
  Next j
End Sub

Private Function MeshFingerprint() As String
  Dim spec As Variant, ws As Worksheet, data As Variant, v As Variant, r As Long, c As Long, last As Long
  Dim a As Double, b As Double, s As String
  a = 17#: b = 31#
  For Each spec In Array(Array("節点データ", 9), Array("要素データ", 10), Array("材料データ", 13), Array("ステージ", 6), Array("接合", 4))
    Set ws = ThisWorkbook.Worksheets(CStr(spec(0)))
    last = ws.cells(ws.rows.count, 1).End(xlUp).row
    data = ws.range(ws.cells(1, 1), ws.cells(Application.Max(2, last), CLng(spec(1)))).value2
    FingerprintText CStr(spec(0)) & ":" & CStr(last), a, b
    For r = 1 To UBound(data, 1)
      For c = 1 To UBound(data, 2)
        v = data(r, c)
        If IsError(v) Then
          s = "ERROR"
        ElseIf VarType(v) = vbDouble Then
          s = Format$(v, "0.00000000000000000E+00")
        Else
          s = CStr(v)
        End If
        FingerprintText CStr(VarType(v)) & ":" & CStr(Len(s)) & ":" & s, a, b
      Next c
    Next r
  Next spec
  MeshFingerprint = CStr(a) & ":" & CStr(b)
End Function

Private Sub PreviewButton(ByVal ws As Worksheet, ByVal name As String, ByVal caption As String, ByVal macro As String, ByVal x As Double, ByVal y As Double)
  Dim shp As shape
  On Error Resume Next: Set shp = ws.Shapes(name): On Error GoTo 0
  If shp Is Nothing Then Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, 245, 32)
  shp.name = name: shp.left = x: shp.Top = y: shp.width = 245: shp.Height = 32
  shp.Visible = msoTrue
  shp.TextFrame2.TextRange.text = caption: shp.TextFrame2.TextRange.Font.size = 11
  shp.TextFrame2.TextRange.Font.fill.ForeColor.RGB = vbWhite
  shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
  shp.fill.ForeColor.RGB = RGB(49, 91, 79): shp.line.Visible = msoFalse
  shp.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & macro
End Sub

Public Sub FEMUiPreviewMesh()
  Dim ws As Worksheet, pic As shape, qualityError As String, eventSaved As Boolean
  If AnalysisRunning Then Exit Sub
  Call FEMUiInvalidateMeshPreview
  FEMUiMeshPreviewError = ""
  eventSaved = Application.EnableEvents
  On Error GoTo Failed
  Application.EnableEvents = False
  On Error Resume Next
  P5EvaluateCurrentSheet "解析前メッシュ確認"
  qualityError = Err.description: Err.Clear
  On Error GoTo Failed
  If Len(qualityError) > 0 Then
    FEMUiMeshQualityNote = "品質評価を確認してください：" & left$(qualityError, 55)
  Else
    FEMUiMeshQualityNote = "品質目標=" & IIf(P5QualityMeetsSettings(), "PASS", "WARN") & " / " & P5QualitySummary()
  End If
  FEMViewAll
  If Len(FEMViewerLastError) > 0 Then Err.Raise vbObjectError + 3895, , FEMViewerLastError
  Set ws = ThisWorkbook.Worksheets("図"): Set pic = ws.Shapes("P1Agg_ViewSVG")
  PreviewButton ws, "UI_MeshStart", "このメッシュで解析を開始", "FEMUiStartFromMesh", 12, pic.Top + pic.Height + 12
  PreviewButton ws, "UI_MeshRepair", "品質修正候補を確認", "ボタン20_Click", 275, pic.Top + pic.Height + 12
  previewSignature = MeshFingerprint(): previewReady = True
Finish:
  Application.EnableEvents = eventSaved
  Exit Sub
Failed:
  FEMUiMeshPreviewError = CStr(Err.number) & " " & Err.description
  If Not SuppressUserMessages Then MsgBox "メッシュ確認表示に失敗しました。" & vbCrLf & FEMUiMeshPreviewError, vbExclamation
  Resume Finish
End Sub

Public Sub FEMUiStartFromMesh()
  On Error GoTo Failed
  If AnalysisRunning Then Exit Sub
  If Not previewReady Then GoTo ReviewAgain
  If previewSignature <> MeshFingerprint() Then GoTo ReviewAgain
  previewReady = False
  Call ボタン5_Click
  Exit Sub
ReviewAgain:
  FEMUiPreviewMesh
  If Not SuppressUserMessages And Len(FEMUiMeshPreviewError) = 0 Then MsgBox "メッシュ・拘束・品質を確認し、図の下の「このメッシュで解析を開始」を押してください。", vbInformation
  Exit Sub
Failed:
  If Not SuppressUserMessages Then MsgBox Err.description, vbExclamation
End Sub
