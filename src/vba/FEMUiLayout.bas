Option Explicit
Public Const FEM_UI_STAMP As String = "20261008_UI_01"

Public Sub FEMUiAfterLayout()
  Dim eventsSaved As Boolean, screenSaved As Boolean, activeName As String
  Dim ws As Worksheet, pageNames As Variant, p As Variant, r As Long, maxRow As Long
  Dim errorNo As Long, errorSource As String, errorText As String
  Dim uiPhase As String
  If Not FEMUiInstalled() Then Exit Sub
  eventsSaved = Application.EnableEvents: screenSaved = Application.ScreenUpdating
  activeName = ThisWorkbook.ActiveSheet.name
  On Error GoTo Failed
  Application.EnableEvents = False: Application.ScreenUpdating = False
  pageNames = Array("メッシュ設定", "解析設定", "高速化設定", "出力・表示", "詳細設定")
  For Each p In pageNames
    ThisWorkbook.Worksheets(CStr(p)).Unprotect
  Next p
  uiPhase = "設定連携": FEMUiBindSettings
  For Each p In pageNames
    Set ws = ThisWorkbook.Worksheets(CStr(p))
    ws.Unprotect
    ws.columns(6).Hidden = True
    ws.range("A1:F100").locked = True
    maxRow = ws.Cells(ws.rows.count, 6).End(xlUp).row
    For r = 1 To maxRow
      If Len(CStr(ws.Cells(r, 6).value2)) > 0 Then ws.Cells(r, 3).MergeArea.locked = False
    Next r
  Next p
  uiPhase = "入力説明": FEMUiRestoreGuides
  uiPhase = "リンク": FEMUiMakeLinks
  uiPhase = "ボタン": FEMUiMakeButtons
  uiPhase = "シート順": FEMUiOrderTabs
  uiPhase = "表示設定": FEMUiSetInputViews
  For Each p In pageNames
    ThisWorkbook.Worksheets(CStr(p)).Protect UserInterfaceOnly:=True, DrawingObjects:=False
  Next p
  If FEMUiWorksheetExists(activeName) Then
    If ThisWorkbook.Worksheets(activeName).Visible = xlSheetVisible Then ThisWorkbook.Worksheets(activeName).Activate
  End If
  Application.EnableEvents = eventsSaved: Application.ScreenUpdating = screenSaved
  FEMPracticalAfterLayout
  FEMInvalidateSettingCache
  Exit Sub
Failed:
  errorNo = Err.Number: errorSource = Err.source: errorText = Err.Description
  Application.EnableEvents = eventsSaved: Application.ScreenUpdating = screenSaved
  Err.Raise errorNo, "FEMUiAfterLayout:" & uiPhase, uiPhase & ": " & errorText
End Sub

Private Function FEMUiWorksheetExists(ByVal sheetName As String) As Boolean
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets(sheetName)
  On Error GoTo 0
  FEMUiWorksheetExists = Not ws Is Nothing
End Function

Private Sub FEMUiList(ByVal ws As Worksheet, ByVal address As String, ByVal choices As String)
  With ws.range(address).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=choices
    .IgnoreBlank = True: .InCellDropdown = True
  End With
End Sub

Private Sub FEMUiRestoreGuides()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets("ステージ")
  If CStr(ws.range("F2").value2) = "自重。底面Y固定・左右X固定は要素作成②後に設定シートの拘束を適用" Then
    ws.range("F2").value2 = "自重。外周の拘束は「メッシュ設定」、個別の支持・変位は「節点データ」で設定。"
  End If
  ws.range("H1:I24").ClearContents
  ws.Cells(1, 8).value2 = "施工ステージの入力"
  ws.Cells(1, 9).value2 = "上から順に実行。有効1=実行、0=スキップ。"
  ws.Cells(2, 8).value2 = "GRAVITY　自重"
  ws.Cells(2, 9).value2 = "自重を計算。支持・変位は節点データで設定。"
  ws.Cells(3, 8).value2 = "LOAD　載荷"
  ws.Cells(3, 9).value2 = "C列に対象材料番号。「載荷」シートの該当行を使用。"
  ws.Cells(4, 8).value2 = "UNLOAD　除荷"
  ws.Cells(4, 9).value2 = "対象材料の荷重・変位を外す。要素は残る。"
  ws.Cells(5, 8).value2 = "BIRTH　追加"
  ws.Cells(5, 9).value2 = "C列の材料の要素を有効化。材料の初期状態をOFFにしておく。"
  ws.Cells(6, 8).value2 = "DEATH　撤去"
  ws.Cells(6, 9).value2 = "C列の材料の要素を無効化。材料単位で指定する。"
  ws.Cells(7, 8).value2 = "SRM　強度低減"
  ws.Cells(7, 9).value2 = "最後の有効行ならFs探索。途中の行はD列の必要Fsを確認する。"
  ws.Cells(8, 8).value2 = "RESET_U"
  ws.Cells(8, 9).value2 = "変位と反力をリセット。応力履歴は保持。"
  ws.Cells(9, 8).value2 = "RESET_STRESS"
  ws.Cells(9, 9).value2 = "応力履歴をリセット。挙動は仕様シートで確認。"
  ws.Cells(10, 8).value2 = "MATSET"
  ws.Cells(10, 9).value2 = "材料状態を変更。D列の書式は仕様シートで確認。"
  ws.Cells(11, 8).value2 = "BIRTHの準備"
  ws.Cells(11, 9).value2 = "追加材料を登録 → 要素へ材料番号を付与 → 初期状態OFF → BIRTH行。"
  ws.Cells(12, 8).value2 = "DEATHの準備"
  ws.Cells(12, 9).value2 = "撤去する範囲を別の材料番号へ分ける → DEATH行へその材料番号。"
  ws.Cells(13, 8).value2 = "注意"
  ws.Cells(13, 9).value2 = "A列の番号を替えても並び順は替わらない。行を並べ替えて順番を変更。"
  FEMUiList ws, "B2:B2000", "GRAVITY,LOAD,UNLOAD,BIRTH,DEATH,SRM,RESET_U,RESET_STRESS,MATSET"
  FEMUiList ws, "E2:E2000", "1,0"
  Set ws = ThisWorkbook.Worksheets("接合")
  ws.range("G1:H20").ClearContents
  ws.Cells(1, 7).value2 = "ジョイントの設定"
  ws.Cells(1, 8).value2 = "異なる材料の共有辺に接合要素を作ります。"
  ws.Cells(2, 7).value2 = "1　材料を登録"
  ws.Cells(2, 8).value2 = "「材料データ」にJOINT材料を追加。種別=JOINT、B列=Kn、L列=Ks。"
  ws.Cells(3, 7).value2 = "2　接合する組合せ"
  ws.Cells(3, 8).value2 = "A列=節点を残す側、B列=節点を増やす側、C列=JOINT材料の番号。"
  ws.Cells(4, 7).value2 = "3　生成"
  ws.Cells(4, 8).value2 = "有効=1にして、メッシュ設定の「形状メッシュを作成②」を実行。"
  ws.Cells(5, 7).value2 = "生成のタイミング"
  ws.Cells(5, 8).value2 = "要素作成②の後と解析開始時に、既存の接合処理を使用。"
  ws.Cells(6, 7).value2 = "入力の意味"
  ws.Cells(6, 8).value2 = "材料A/Bは材料番号。接合材料にも番号を入力。説明は任意。"
  ws.Cells(7, 7).value2 = "剛性と引張"
  ws.Cells(7, 8).value2 = "Kn/Ksは正。M列の引張=0で開口可、1で引張を保持。"
  ws.Cells(8, 7).value2 = "高速化の扱い"
  ws.Cells(8, 8).value2 = "ジョイントを含むモデルでは対応していない高速化方式を見送る。"
  FEMUiList ws, "D2:D2000", "1,0"
  Set ws = ThisWorkbook.Worksheets("載荷")
  FEMUiList ws, "B2:B2000", "TOP,BOTTOM,LEFT,RIGHT,BOX"
  FEMUiList ws, "C2:C2000", "X,Y"
  FEMUiList ws, "D2:D2000", "DISP,FORCE,FIX"
  Set ws = ThisWorkbook.Worksheets("材料データ")
  FEMUiList ws, "I2:I2000", "1,0"
  FEMUiList ws, "J2:J2000", "ON,OFF"
  FEMUiList ws, "K2:K2000", "SOIL,STRUCT,JOINT"
  FEMUiList ws, "M2:M2000", "0,1"
End Sub

Private Sub FEMUiLink(ByVal sheetName As String, ByVal address As String, ByVal caption As String, ByVal destination As String, ByVal targetCell As String)
  Dim ws As Worksheet, cell As range
  Set ws = ThisWorkbook.Worksheets(sheetName): Set cell = ws.range(address)
  cell.Hyperlinks.Delete
  ws.Hyperlinks.Add Anchor:=cell, address:="", SubAddress:="'" & Replace(destination, "'", "''") & "'!" & targetCell, TextToDisplay:=caption
  cell.Font.Color = RGB(38, 94, 146): cell.Font.size = 11
End Sub
Private Sub FEMUiMakeLinks()
  FEMUiLink "メッシュ設定", "B4", "操作パネルへ", "操作パネル", "B2"
  FEMUiLink "メッシュ設定", "B33", "形状の座標・境界を入力", "形状入力", "A1"
  FEMUiLink "メッシュ設定", "E33", "材料データを入力", "材料データ", "A1"
  FEMUiLink "解析設定", "B4", "操作パネルへ", "操作パネル", "B2"
  FEMUiLink "高速化設定", "B4", "操作パネルへ", "操作パネル", "B2"
  FEMUiLink "出力・表示", "B4", "操作パネルへ", "操作パネル", "B2"
  FEMUiLink "詳細設定", "B4", "操作パネルへ", "操作パネル", "B2"
  FEMUiLink "形状入力", "B4", "操作パネルへ", "操作パネル", "B2"
  FEMUiLink "形状入力", "G4", "メッシュ条件・作成へ", "メッシュ設定", "A1"
  FEMUiLink "操作パネル", "B8", "材料：地盤・構造・JOINT", "材料データ", "A1"
  FEMUiLink "操作パネル", "B10", "メッシュ条件・生成", "メッシュ設定", "A1"
  FEMUiLink "操作パネル", "B12", "ジョイント（接合する場合）", "接合", "A1"
  FEMUiLink "操作パネル", "F8", "ステージ：BIRTH / DEATH / SRM", "ステージ", "A1"
  FEMUiLink "操作パネル", "F10", "載荷：荷重・変位の位置と値", "載荷", "A1"
  FEMUiLink "操作パネル", "F12", "解析・SRMの設定", "解析設定", "A1"
  FEMUiLink "操作パネル", "B20", "結果図", "図", "A1"
  FEMUiLink "操作パネル", "B22", "ステージごとの結果", "ステージ結果", "A1"
  FEMUiLink "操作パネル", "F18", "高速化の比較設定", "高速化設定", "A1"
  FEMUiLink "操作パネル", "F20", "出力・再開・図の表示", "出力・表示", "A1"
  FEMUiLink "操作パネル", "F22", "メッシュ品質・ソルバーの詳細", "詳細設定", "A1"
End Sub

Private Sub FEMUiButton(ByVal sheetName As String, ByVal objectName As String, ByVal address As String, ByVal caption As String, ByVal macroName As String, Optional ByVal primary As Boolean = False)
  Dim ws As Worksheet, target As range, shp As shape
  Set ws = ThisWorkbook.Worksheets(sheetName): Set target = ws.range(address)
  On Error Resume Next
  Set shp = ws.Shapes(objectName)
  On Error GoTo 0
  On Error GoTo Failed
  If shp Is Nothing Then
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, target.left + 2, target.Top + 2, target.width - 4, target.Height - 4)
    shp.name = objectName
  End If
  shp.left = target.left + 2: shp.Top = target.Top + 2
  shp.width = target.width - 4: shp.Height = target.Height - 4
  shp.TextFrame2.TextRange.text = caption
  shp.TextFrame2.TextRange.Font.name = "Arial": shp.TextFrame2.TextRange.Font.size = 11
  shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
  shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
  shp.line.Visible = msoFalse
  If primary Then
    shp.fill.ForeColor.RGB = RGB(51, 92, 80): shp.TextFrame2.TextRange.Font.fill.ForeColor.RGB = RGB(255, 255, 255)
  Else
    shp.fill.ForeColor.RGB = RGB(231, 237, 244): shp.TextFrame2.TextRange.Font.fill.ForeColor.RGB = RGB(36, 52, 71)
  End If
  shp.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & macroName
  shp.Placement = xlMoveAndSize
  Exit Sub
Failed:
  Err.Raise Err.Number, "FEMUiButton", sheetName & "/" & objectName & ": " & Err.Description
End Sub
Private Sub FEMUiMakeButtons()
  Dim ws As Worksheet, shp As shape, actionText As String, marker As Long
  On Error GoTo Failed
  FEMUiButton "操作パネル", "UI_Run", "B18:D19", "解析を実行", "FEMUiRun", True
  FEMUiButton "操作パネル", "UI_Nodes", "B31:C31", "節点・拘束", "FEMUiShowNodes", False
  FEMUiButton "操作パネル", "UI_Elements", "D31", "要素", "FEMUiShowElements", False
  FEMUiButton "操作パネル", "UI_Results", "F31:G31", "数値結果", "FEMUiShowResults", False
  FEMUiButton "操作パネル", "UI_Diagnostic", "H31", "診断", "FEMUiShowDiagnostic", False
  FEMUiButton "操作パネル", "UI_Quantities", "B33:C33", "表示項目一覧", "FEMUiShowQuantities", False
  FEMUiButton "操作パネル", "UI_Spec", "D33", "仕様・制約", "FEMUiShowSpecs", False
  FEMUiButton "メッシュ設定", "UI_Mesh1", "B35:C35", "簡易メッシュを作成①", "ボタン1_Click", False
  FEMUiButton "メッシュ設定", "UI_Mesh2", "E35", "形状メッシュを作成②", "FEMUiGenerateShape", False
  FEMUiButton "メッシュ設定", "UI_Repair", "B37:C37", "メッシュ品質を修正", "ボタン20_Click", False
  FEMUiButton "メッシュ設定", "UI_ViewMesh", "E37", "現在の形状を表示", "FEMViewInitial", False
  FEMUiButton "メッシュ設定", "UI_Remesh", "B39:C39", "塑性域から再メッシュ候補", "MeshPlasticRemesh", False
  FEMUiButton "出力・表示", "UI_ResultFigure", "B37:C38", "変形・応力図を表示", "FEMViewResult", False
  ' Qualify all retained action buttons against this copy, including remesh adoption.
  For Each ws In ThisWorkbook.Worksheets
    For Each shp In ws.Shapes
      ' Filter dropdowns expose Shapes but do not support OnAction.
      actionText = ""
      On Error Resume Next
      actionText = shp.OnAction
      On Error GoTo Failed
      If Len(actionText) > 0 Then
        marker = InStrRev(actionText, "!")
        If marker > 0 Then actionText = mid$(actionText, marker + 1)
        shp.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & actionText
      End If
    Next shp
  Next ws
  Exit Sub
Failed:
  Err.Raise Err.Number, "FEMUiMakeButtons", ws.name & "/" & shp.name & ": " & Err.Description
End Sub

Private Sub FEMUiOrderTabs()
  Dim names As Variant, p As Variant, previous As Worksheet, ws As Worksheet
  names = Array("操作パネル", "材料データ", "メッシュ設定", "形状入力", "接合", "ステージ", "載荷", "解析設定", "図", "ステージ結果", "高速化設定", "出力・表示", "詳細設定")
  For Each p In names
    Set ws = ThisWorkbook.Worksheets(CStr(p)): ws.Visible = xlSheetVisible
    If previous Is Nothing Then ws.Move Before:=ThisWorkbook.Worksheets(1) Else ws.Move After:=previous
    Set previous = ws
  Next p
  ThisWorkbook.Worksheets("設定").Visible = xlSheetVeryHidden
  For Each p In Array("図形定義", "節点データ", "要素データ", "結果節点", "結果要素", "診断", "仕様")
    ThisWorkbook.Worksheets(CStr(p)).Visible = xlSheetHidden
  Next p
End Sub

Private Sub FEMUiSetInputViews()
  Dim names As Variant, p As Variant, ws As Worksheet
  names = Array("操作パネル", "メッシュ設定", "形状入力", "解析設定", "高速化設定", "出力・表示", "詳細設定", "材料データ", "ステージ", "載荷", "接合")
  For Each p In names
    Set ws = ThisWorkbook.Worksheets(CStr(p)): ws.Activate
    ActiveWindow.DisplayGridlines = False: ActiveWindow.Zoom = 85
    ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
    ActiveWindow.FreezePanes = False
    If FEMUiIsSettingsSheet(ws.name) Or ws.name = "形状入力" Then
      ws.range("B8").Select: ActiveWindow.FreezePanes = True
    ElseIf ws.name <> "操作パネル" Then
      ws.range("B2").Select: ActiveWindow.FreezePanes = True
    End If
  Next p
  With ThisWorkbook.Worksheets("材料データ")
    .rows("2:200").RowHeight = 29
    .columns("A").ColumnWidth = 11
    .columns("B:H").ColumnWidth = 15
    .columns("I:M").ColumnWidth = 13
    .columns("O").ColumnWidth = 20: .columns("P").ColumnWidth = 68
    .range("O1:P7").WrapText = True
  End With
  With ThisWorkbook.Worksheets("載荷")
    .rows("2:200").RowHeight = 29
    .columns("A").ColumnWidth = 13: .columns("B:D").ColumnWidth = 15
    .columns("E:I").ColumnWidth = 12: .columns("J").ColumnWidth = 40
    .columns("L").ColumnWidth = 18: .columns("M").ColumnWidth = 60
    .range("L1:M13").WrapText = True
  End With
  With ThisWorkbook.Worksheets("図形定義")
    .columns("A").ColumnWidth = 8: .columns("B:C").ColumnWidth = 14
    .columns("D:F").ColumnWidth = 14
    .columns("G:N").ColumnWidth = 13
  End With
End Sub

Public Sub FEMUiShowHome()
  If Not FEMUiInstalled() Then Exit Sub
  FEMUiOpen "操作パネル"
End Sub
Private Sub FEMUiOpen(ByVal name As String)
  With ThisWorkbook.Worksheets(name)
    .Visible = xlSheetVisible: .Activate
    Application.Goto .range("A1"), True
  End With
End Sub
Public Sub FEMUiShowNodes(): FEMUiOpen "節点データ": End Sub
Public Sub FEMUiShowElements(): FEMUiOpen "要素データ": End Sub
Public Sub FEMUiShowResults(): FEMUiOpen "結果節点": End Sub
Public Sub FEMUiShowDiagnostic(): FEMUiOpen "診断": End Sub
Public Sub FEMUiShowQuantities(): FEMUiOpen "表示量一覧": End Sub
Public Sub FEMUiShowSpecs(): FEMUiOpen "仕様": End Sub
Public Sub FEMUiRun(): Call ボタン5_Click: End Sub
