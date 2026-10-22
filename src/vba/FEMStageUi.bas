Option Explicit

' Presentation only. All editable cells write through to the original input tables.
Public Const FEM24_STAMP As String = "20261024_STAGE_01"
Public FEM24Error As String
Private mBusy As Boolean
Private mSnapshots As Object

Public Function FEM24CaptureCells(ByVal area As range) As Variant
  Dim saved(0 To 1) As Variant, formulas As Collection, found As range, c As range
  Set formulas = New Collection
  saved(0) = area.value2
  On Error Resume Next
  Set found = area.SpecialCells(xlCellTypeFormulas)
  Err.Clear
  On Error GoTo 0
  If Not found Is Nothing Then
    For Each c In found.cells
      If Not Intersect(c, area) Is Nothing Then formulas.Add Array(c.row - area.row + 1, c.column - area.column + 1, c.formula)
    Next c
  End If
  Set saved(1) = formulas
  FEM24CaptureCells = saved
End Function

Public Sub FEM24RestoreCells(ByVal area As range, ByVal saved As Variant)
  Dim formulas As Collection, item As Variant
  area.value2 = saved(0)
  Set formulas = saved(1)
  For Each item In formulas
    area.cells(CLng(item(0)), CLng(item(1))).formula = item(2)
  Next item
End Sub

Private Sub CacheReset(ByVal page As String)
  Dim records As Collection
  If mSnapshots Is Nothing Then Set mSnapshots = CreateObject("Scripting.Dictionary")
  If mSnapshots.Exists(page) Then mSnapshots.Remove page
  Set records = New Collection
  mSnapshots.Add page, records
End Sub

Private Sub CacheBlock(ByVal page As String, ByVal source As String, ByVal address As String)
  Dim records As Collection
  Set records = mSnapshots(page)
  records.Add Array(source, address, ThisWorkbook.Worksheets(source).range(address).value2)
End Sub

Private Function CacheState(ByVal kind As String) As String
  CacheState = ThisWorkbook.names("_FEMViewerResultsDirty").RefersTo
  If kind = "HOME" Then CacheState = CacheState & "|" & CStr(ThisWorkbook.Worksheets("節点データ").cells(rows.count, 1).End(xlUp).row) & "|" & CStr(ThisWorkbook.Worksheets("要素データ").cells(rows.count, 1).End(xlUp).row)
End Function

Private Sub CacheMeta(ByVal page As String, ByVal kind As String)
  Dim records As Collection
  Set records = mSnapshots(page)
  records.Add Array("#", kind, CacheState(kind))
End Sub

Private Function SameValues(ByVal a As Variant, ByVal b As Variant) As Boolean
  Dim r As Long, c As Long
  If IsArray(a) <> IsArray(b) Then Exit Function
  If Not IsArray(a) Then
    If VarType(a) <> VarType(b) Then Exit Function
    SameValues = (CStr(a) = CStr(b))
    Exit Function
  End If
  If UBound(a, 1) <> UBound(b, 1) Or UBound(a, 2) <> UBound(b, 2) Then Exit Function
  For r = 1 To UBound(a, 1)
    For c = 1 To UBound(a, 2)
      If VarType(a(r, c)) <> VarType(b(r, c)) Then Exit Function
      If IsError(a(r, c)) Then
        If CStr(a(r, c)) <> CStr(b(r, c)) Then Exit Function
      ElseIf a(r, c) <> b(r, c) Then
        Exit Function
      End If
    Next c
  Next r
  SameValues = True
End Function

Public Function FEM24PageCurrent(ByVal page As String) As Boolean
  Dim records As Collection, item As Variant, current As Variant
  On Error GoTo Stale
  If mSnapshots Is Nothing Then Exit Function
  If Not mSnapshots.Exists(page) Then Exit Function
  Set records = mSnapshots(page)
  If records.count = 0 Then Exit Function
  For Each item In records
    If CStr(item(0)) = "#" Then
      current = CacheState(CStr(item(1)))
    Else
      current = ThisWorkbook.Worksheets(CStr(item(0))).range(CStr(item(1))).value2
    End If
    If Not SameValues(item(2), current) Then Exit Function
  Next item
  FEM24PageCurrent = True
Stale:
End Function

Private Sub CacheRenew(ByVal page As String)
  Dim old As Collection, item As Variant
  If mSnapshots Is Nothing Then Exit Sub
  If Not mSnapshots.Exists(page) Then Exit Sub
  Set old = mSnapshots(page)
  CacheReset page
  For Each item In old
    If CStr(item(0)) = "#" Then
      CacheMeta page, CStr(item(1))
    Else
      CacheBlock page, CStr(item(0)), CStr(item(1))
    End If
  Next item
End Sub

Private Function PageNames() As Variant
  PageNames = Array("ホーム", "モデル", "材料・接合", "境界・水条件", "施工工程", "解析・保存", "結果")
End Function

Public Function FEM24Installed() As Boolean
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("ホーム")
  FEM24Installed = Not ws Is Nothing
End Function

Public Function FEM24IsPage(ByVal page As String) As Boolean
  Dim p As Variant
  For Each p In PageNames()
    If CStr(p) = page Then FEM24IsPage = True: Exit Function
  Next p
End Function

Private Function Sections(ByVal page As String) As String
  Select Case page
    Case "ホーム": Sections = "作業の流れ"
    Case "モデル": Sections = "形状の座標,形状の領域,メッシュ生成・支持条件"
    Case "材料・接合": Sections = "材料特性,接合条件,材料の水理特性"
    Case "境界・水条件": Sections = "水条件の基本,水位線,水理境界,載荷条件,節点の支持・変位"
    Case "施工工程": Sections = "工程の編集"
    Case "解析・保存": Sections = "解析・SRM,保存・再開,高速化,詳細設定"
    Case "結果": Sections = "図と概要,表示設定,工程別数値,節点結果,要素結果,水圧結果,接合結果,診断"
  End Select
End Function

Private Sub Button(ByVal ws As Worksheet, ByVal id As String, ByVal label As String, ByVal action As String, ByVal x As Double, ByVal y As Double, ByVal w As Double)
  Dim s As shape
  Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, 24)
  s.name = "UI24_" & id: s.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & action
  s.fill.ForeColor.RGB = RGB(31, 78, 121): s.line.Visible = msoFalse
  s.TextFrame2.TextRange.text = label
  s.TextFrame2.TextRange.Font.size = 10: s.TextFrame2.TextRange.Font.fill.ForeColor.RGB = vbWhite
  s.TextFrame2.VerticalAnchor = msoAnchorMiddle
  s.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
  s.Placement = xlFreeFloating
End Sub

Public Sub FEM24Install()
  Dim p As Variant, ws As Worksheet, oldEvents As Boolean, oldScreen As Boolean
  If AnalysisRunning Then Exit Sub
  oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
  On Error GoTo Failed
  Application.EnableEvents = False: Application.ScreenUpdating = False: mBusy = True
  With ThisWorkbook.Worksheets("解析設定")
    .Unprotect
    .range("E9").value2 = "TOTAL：入力γによる自重。水圧・外水圧は水条件で別途計算。水中重量へ自動変換しません。"
    .Protect UserInterfaceOnly:=True, DrawingObjects:=False
  End With
  For Each p In PageNames()
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CStr(p))
    On Error GoTo Failed
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = CStr(p)
    Frame ws
    If Len(CStr(ws.range("C4").value2)) = 0 Then ws.range("C4").value2 = Split(Sections(ws.name), ",")(0)
    Render ws
  Next p
  FEM24ApplyTabs
  mBusy = False: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
  Exit Sub
Failed:
  FEM24Error = "画面の準備: " & Err.description
  mBusy = False: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
  Err.Raise vbObjectError + 2440, "FEM24Install", FEM24Error
End Sub

Private Sub Frame(ByVal ws As Worksheet)
  Dim i As Long, s As shape, p As Variant, actions As Variant
  ws.Unprotect
  For i = ws.Shapes.count To 1 Step -1
    If left$(ws.Shapes(i).name, 5) = "UI24_" Then ws.Shapes(i).Delete
  Next i
  ws.range("A1:M6").UnMerge
  ws.range("A1:M6").Clear
  ws.range("A1:O2007").Font.name = "Meiryo UI": ws.range("A1:O2007").Font.size = 10
  ws.columns("A:M").ColumnWidth = 12
  ws.columns("N:O").Hidden = True
  ws.rows(1).RowHeight = 30: ws.rows(2).RowHeight = 29: ws.rows(3).RowHeight = 29
  ws.rows(4).RowHeight = 25: ws.rows("5:6").RowHeight = 22
  ws.range("A1:M1").Merge: ws.range("A1").value2 = "2D Soil FEM  |  " & ws.name & "  |  1024"
  ws.range("A1:M1").Interior.Color = RGB(20, 43, 66): ws.range("A1").Font.Color = vbWhite: ws.range("A1").Font.size = 17
  actions = Array("FEM24Home", "FEM24Model", "FEM24Materials", "FEM24Water", "FEM24Stages", "FEM24Analysis", "FEM24Results")
  i = 0
  For Each p In PageNames()
    Button ws, "Nav" & CStr(i), CStr(i + 1) & " " & CStr(p), CStr(actions(i)), 5 + i * 131, ws.rows(2).Top + 2, 126
    i = i + 1
  Next p
  ws.range("A4:B4").Merge: ws.range("A4").value2 = "表示する項目 →"
  ws.range("C4:F4").Merge: ws.range("C4:F4").Interior.Color = RGB(219, 234, 251)
  ws.range("C4:F4").locked = False
  With ws.range("C4").Validation
    .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, Sections(ws.name)
    .InCellDropdown = True: .ShowError = True
  End With
  ws.range("G4:J4").Merge: ws.range("G4").value2 = "青いセル：入力  /  灰色：参照"
  If ws.name = "結果" Then
    ws.range("K4").value2 = "参照ページ"
    ws.range("L4").value2 = 1: ws.range("L4").locked = False
  End If
  ws.range("A5:M6").Merge: ws.range("A5").WrapText = True
  Select Case ws.name
    Case "モデル"
      Button ws, "Generate", "形状からメッシュ作成", "FEM24Generate", 5, ws.rows(3).Top + 2, 175
      Button ws, "Preview", "解析前メッシュを確認", "FEM24Preview", 187, ws.rows(3).Top + 2, 175
      Button ws, "Raw", "詳細データを開く", "FEM24Detail", 370, ws.rows(3).Top + 2, 150
    Case "境界・水条件"
      Button ws, "WaterPreview", "水圧を確認（指定・定常）", "FEM24WaterPreview", 5, ws.rows(3).Top + 2, 195
      Button ws, "Preview", "解析前メッシュを確認", "FEM24Preview", 207, ws.rows(3).Top + 2, 175
      Button ws, "Detail", "元の入力表を開く", "FEM24Detail", 390, ws.rows(3).Top + 2, 150
    Case "施工工程"
      actions = Array("FEM24StageAdd", "FEM24StageCopy", "FEM24StageDelete", "FEM24StageUp", "FEM24StageDown", "FEM24StageUndo")
      p = Array("追加", "複製", "削除", "上へ", "下へ", "元に戻す")
      For i = 0 To 5
        Button ws, "Edit" & CStr(i), CStr(p(i)), CStr(actions(i)), 5 + i * 100, ws.rows(3).Top + 2, 94
      Next i
      Button ws, "Detail", "元の工程表", "FEM24Detail", 613, ws.rows(3).Top + 2, 120
    Case "解析・保存"
      Button ws, "Preview", "解析前メッシュを確認", "FEM24Preview", 5, ws.rows(3).Top + 2, 175
      Button ws, "Run", "確認したメッシュで解析", "FEM24Run", 187, ws.rows(3).Top + 2, 185
      Button ws, "Save", "ブックを保存", "FEM24Save", 380, ws.rows(3).Top + 2, 125
    Case "結果"
      Button ws, "Plot", "最終保存状態の変形・応力図", "FEM24Plot", 5, ws.rows(3).Top + 2, 215
      Button ws, "Water", "水位・水頭図（詳細）", "FEM24WaterFigure", 227, ws.rows(3).Top + 2, 170
      Button ws, "Preview", "解析前メッシュを確認", "FEM24Preview", 405, ws.rows(3).Top + 2, 175
      Button ws, "Detail", "元の結果表を開く", "FEM24Detail", 588, ws.rows(3).Top + 2, 150
    Case Else
      Button ws, "Detail", "元の入力表・説明を開く", "FEM24Detail", 5, ws.rows(3).Top + 2, 190
  End Select
End Sub

Public Sub FEM24ApplyTabs()
  Dim p As Variant, ws As Worksheet, prev As Worksheet, e As Boolean, screen As Boolean, errorNo As Long, errorText As String
  If Not FEM24Installed() Then Exit Sub
  e = Application.EnableEvents: screen = Application.ScreenUpdating
  On Error GoTo Failed
  Application.EnableEvents = False: Application.ScreenUpdating = False
  For Each p In PageNames()
    Set ws = ThisWorkbook.Worksheets(CStr(p))
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    If prev Is Nothing Then
      If ws.Index <> 1 Then ws.Move before:=ThisWorkbook.Worksheets(1)
    Else
      If ws.Index <> prev.Index + 1 Then ws.Move After:=prev
    End If
    Set prev = ws
  Next p
  If Not FEM24IsPage(ThisWorkbook.ActiveSheet.name) Then ThisWorkbook.Worksheets("ホーム").Activate
  For Each ws In ThisWorkbook.Worksheets
    If Not FEM24IsPage(ws.name) And ws.Visible = xlSheetVisible Then
      If ws.name = "設定" Or ws.name = "表示量一覧" Or ws.name = "メッシュ引継" Or ws.name = "再メッシュ候補" Then
        ws.Visible = xlSheetVeryHidden
      Else
        ws.Visible = xlSheetHidden
      End If
    End If
  Next ws
  Application.EnableEvents = e: Application.ScreenUpdating = screen
  Exit Sub
Failed:
  errorNo = Err.number: errorText = Err.description
  Application.EnableEvents = e: Application.ScreenUpdating = screen
  Err.Raise errorNo, "FEM24ApplyTabs", errorText
End Sub

' Copy list and constant numeric rules without overwriting the user's clipboard.
' Row-dependent custom rules remain enforced by the existing analysis preflight.
Private Sub CloneValidation(ByVal src As range, ByVal dest As range)
  Dim t As Long, op As Long, f1 As String, f2 As String
  On Error Resume Next
  t = src.Validation.Type: f1 = src.Validation.Formula1
  op = src.Validation.Operator: f2 = src.Validation.Formula2
  On Error GoTo 0
  If t = xlValidateList Then
    If InStr(f1, "!") > 0 Or InStr(f1, "$") > 0 Or InStr(f1, "(") > 0 Then Exit Sub
  ElseIf t >= xlValidateWholeNumber And t <= xlValidateTextLength And t <> xlValidateList Then
    If Not IsNumeric(f1) Then Exit Sub
    If Len(f2) > 0 Then
      If Not IsNumeric(f2) Then Exit Sub
    End If
  Else
    Exit Sub
  End If
  With dest.Validation
    .Delete
    If t = xlValidateList Then
      .Add Type:=t, AlertStyle:=xlValidAlertStop, Formula1:=f1
    Else
      .Add Type:=t, AlertStyle:=xlValidAlertStop, Operator:=op, Formula1:=f1, Formula2:=f2
    End If
    .IgnoreBlank = src.Validation.IgnoreBlank
    .ShowInput = src.Validation.ShowInput: .ShowError = src.Validation.ShowError
    .InputTitle = src.Validation.InputTitle: .InputMessage = src.Validation.InputMessage
    .ErrorTitle = src.Validation.ErrorTitle: .ErrorMessage = src.Validation.ErrorMessage
    If t = xlValidateList Then .InCellDropdown = True
  End With
End Sub

Private Sub Table(ByVal ws As Worksheet, ByVal source As String, ByVal address As String, ByVal headerRow As Long, ByVal note As String, Optional ByVal editable As Boolean = True)
  Dim src As Worksheet, a As range, dest As range, col As Long, r As Long
  Set src = ThisWorkbook.Worksheets(source): Set a = src.range(address)
  Set dest = ws.cells(8, 1).Resize(a.rows.count, a.columns.count)
  ws.cells(7, 1).Resize(1, a.columns.count).value2 = src.cells(headerRow, a.column).Resize(1, a.columns.count).value2
  dest.value2 = a.value2
  If editable Then
    dest.locked = False: dest.Interior.Color = RGB(235, 245, 255): dest.Font.Color = RGB(0, 66, 130)
    For col = 1 To a.columns.count
      If source = "水条件" And address = "A3:B6" Then
        For r = 1 To a.rows.count
          CloneValidation a.cells(r, col), dest.cells(r, col)
        Next r
      Else
        CloneValidation a.cells(1, col), dest.columns(col)
      End If
    Next col
    ws.range("N1").value2 = source: ws.range("N2").value2 = address
    ws.range("N3").value2 = dest.address
  End If
  ws.range("A5").value2 = note
  CacheBlock ws.name, source, address
End Sub

Private Sub Settings(ByVal ws As Worksheet, ByVal source As String, Optional ByVal mode As String = "")
  Dim src As Worksheet, r As Long, d As Long, key As String, include As Boolean
  Set src = ThisWorkbook.Worksheets(source): d = 8
  ws.range("A7:C7").Merge: ws.range("A7").value2 = "項目"
  ws.range("D7:E7").Merge: ws.range("D7").value2 = "現在の設定値（入力）"
  ws.range("F7:M7").Merge: ws.range("F7").value2 = "意味・注意"
  For r = 1 To src.cells(src.rows.count, 6).End(xlUp).row
    key = CStr(src.cells(r, 6).value2): include = (Len(key) > 0)
    If mode = "VIEW" Then include = (left$(key, 5) = "VIEW_")
    If mode = "SAVE" Then include = (Len(key) > 0 And left$(key, 5) <> "VIEW_")
    If mode = "MESH" Then include = (left$(key, 8) = "MESH_BC_" Or key = "MESH_MATERIAL" Or key = "MESH_GEN_COMPARE")
    If include Then
      ws.range(ws.cells(d, 1), ws.cells(d, 3)).Merge: ws.cells(d, 1).value2 = src.cells(r, 2).value2
      ws.range(ws.cells(d, 4), ws.cells(d, 5)).Merge: ws.cells(d, 4).value2 = src.cells(r, 3).value2
      ws.range(ws.cells(d, 4), ws.cells(d, 5)).WrapText = True
      ws.range(ws.cells(d, 6), ws.cells(d, 13)).Merge: ws.cells(d, 6).value2 = src.cells(r, 5).value2
      ws.cells(d, 6).WrapText = True: ws.rows(d).RowHeight = 38
      ws.cells(d, 14).value2 = source: ws.cells(d, 15).value2 = r
      If left$(key, 14) <> "MESH_GEN_LAST_" Then
        ws.range(ws.cells(d, 4), ws.cells(d, 5)).locked = False
        ws.range(ws.cells(d, 4), ws.cells(d, 5)).Interior.Color = RGB(235, 245, 255)
        CloneValidation src.cells(r, 3), ws.cells(d, 4).MergeArea
      End If
      d = d + 1
    End If
  Next r
  ws.range("A5").value2 = "値は従来の計算設定へ直接反映します。材料定数・許容誤差・Fs条件は解析目的に合わせて確認してください。"
  CacheBlock ws.name, source, "C1:J200"
End Sub

Private Sub Render(ByVal ws As Worksheet)
  Dim section As String, src As Worksheet, a As range, row0 As Long, name As String, i As Long
  ws.Unprotect
  CacheReset ws.name
  ws.columns("A:M").ColumnWidth = 12
  For i = ws.CheckBoxes.count To 1 Step -1
    ws.CheckBoxes(i).Delete
  Next i
  If ws.name = "結果" And CStr(ws.range("C4").value2) <> "図と概要" Then
    On Error Resume Next
    ws.Shapes("UI24Plot").Delete
    On Error GoTo 0
  End If
  ws.range("A7:O2007").UnMerge: ws.range("A7:O2007").Clear
  ws.range("A7:O2007").Validation.Delete
  ws.range("N1:N3").ClearContents
  ws.range("A7:M2007").locked = True
  ws.rows("7:2007").RowHeight = 21
  ws.range("A7:M7").Interior.Color = RGB(55, 77, 94): ws.range("A7:M7").Font.Color = vbWhite
  section = CStr(ws.range("C4").value2)
  Select Case ws.name
    Case "ホーム"
      ws.range("A5").value2 = "施工工程を中心に、条件を登録 → メッシュ確認 → 解析 → 結果を確認します。項目欄で同じ画面内の表を切り替えます。"
      ws.range("A8:M9").Merge: ws.range("A8").value2 = "1 モデル → 2 材料・接合 → 3 境界・水条件 → 4 施工工程 → 5 解析・保存 → 6 結果"
      ws.range("A11:M12").Merge: ws.range("A11").value2 = "工程は上から実行。追加・複製・削除・上下移動が可能です。途中のSRMは必要Fsの照査、最後の有効行のSRMはFs探索です。"
      ws.range("A14:M15").Merge: ws.range("A14").value2 = "入力変更後は結果を無効化し、メッシュ確認をやり直します。詳細表を開いた後は上部の「7画面へ戻る」で戻れます。"
      ws.range("A17:M18").Merge: ws.range("A17").value2 = "水条件：指定水位の静水圧／定常浸透→力学の片方向連携。時間圧密・FIXED_STRESSは現版では実行できません。"
      ws.range("A20:M21").Merge: ws.range("A20").value2 = "節点 " & CStr(ThisWorkbook.Worksheets("節点データ").cells(rows.count, 1).End(xlUp).row - 1) & " / 要素 " & CStr(ThisWorkbook.Worksheets("要素データ").cells(rows.count, 1).End(xlUp).row - 1) & " / " & ResultState()
      ws.range("A8:M21").WrapText = True
      CacheMeta ws.name, "HOME"
    Case "モデル"
      Select Case section
        Case "形状の座標": Table ws, "形状入力", "B8:D2000", 7, "輪郭点を登録します。点番号は領域の節点順で参照。座標[m]。数式で座標を作る場合は「詳細データを開く」で元の形状入力表を編集します。"
        Case "形状の領域": Table ws, "形状入力", "G8:J2000", 7, "外側・材料領域・穴の節点順を一周順に指定。格子間隔は正の長さ。入力後にメッシュ作成を実行します。"
        Case Else: Settings ws, "メッシュ設定", "MESH": ws.range("A5").value2 = "ここは次回生成時の支持条件です。生成済み節点の実際の拘束は「境界・水条件 → 節点の支持・変位」で確認。再メッシュ前に節点番号依存の荷重・水理境界・接合を見直します。"
      End Select
    Case "材料・接合"
      Select Case section
        Case "材料特性": Table ws, "材料データ", "A2:M2000", 1, "SOIL/STRUCT/JOINTを材料ごとに指定。BIRTH待ちは初期状態OFF。SRM対象は強度低減=1。物性・重量は解析条件に合わせます。"
        Case "接合条件": Table ws, "接合", "A2:G2000", 1, "材料A/Bの共有辺へ接合を生成。水圧方式SOIL/CONTINUOUS/PRESCRIBED。接合を含む場合はメッシュ生成方式の適用範囲を確認。"
        Case Else: Table ws, "水条件", "F12:J211", 11, "材料番号に対応させ、TOTAL/EFFECTIVE・透水係数K[m/s]・水位線IDを指定。水頭0は非透水節点では未計算値です。"
      End Select
    Case "境界・水条件"
      Select Case section
        Case "水条件の基本"
          Table ws, "水条件", "A3:B6", 2, "NONE=水なし / HYDROSTATIC=指定水位から静水圧 / STEADY=飽和定常浸透。未知の自由水面・時間圧密は解きません。"
          ws.range("A7").value2 = "項目": ws.range("B7").value2 = "設定値"
          ws.range("A8:A11").locked = True: ws.columns("A").ColumnWidth = 36
        Case "水位線": Table ws, "水条件", "A12:D211", 11, "同じIDで左→右へXと水位標高を複数行登録。折れ点の間は直線補間。地表より高い指定は外水圧の判定にも使います。"
        Case "水理境界": Table ws, "水条件", "A216:F415", 215, "STEADYの水頭・流量・不透水境界等を指定。節点/辺の番号は再メッシュで変わるため再設定が必要です。"
        Case "載荷条件": Table ws, "載荷", "A2:J2000", 1, "工程のLOAD/UNLOAD対象材料に対応。DISP=強制変位、FORCE=節点力、FIX=固定。値0も有効な入力です。"
        Case Else: Table ws, "節点データ", "A2:I2001", 1, "生成済みメッシュの実際の拘束。D/E:0自由・1固定。座標/番号を変える場合は要素との整合が必要です。2000行を超えるモデルは詳細表で編集してください。"
      End Select
    Case "施工工程"
      ws.columns("A:I").ColumnWidth = 14: ws.columns("F").ColumnWidth = 23
      FEM24StagesRender ws
      CacheBlock ws.name, "ステージ", "A2:F2000"
      CacheBlock ws.name, "水条件", "L12:T211"
      CacheBlock ws.name, "水条件", "B3:B6"
      ws.range("B8:H2006").locked = False: ws.range("B8:H2006").Interior.Color = RGB(235, 245, 255)
      With ws.range("B8:B2006").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "自重,載荷,除荷,盛土・追加,掘削・撤去,水位変更,SRM照査,変位リセット,応力リセット,材料変更"
        .ShowError = True
      End With
      ws.range("A5").value2 = "上から順に施工。有効1=実行/0=スキップ。D列:倍率・途中SRMの必要Fs・材料変更式。水位線はWATER行のH列で選択（N→従来C→既定）。IDは移動しても維持。"
    Case "解析・保存"
      Select Case section
        Case "保存・再開": Settings ws, "出力・表示", "SAVE"
        Case "高速化": Settings ws, "高速化設定"
        Case "詳細設定": Settings ws, "詳細設定"
        Case Else: Settings ws, "解析設定"
      End Select
    Case "結果"
      Select Case section
        Case "表示設定": Settings ws, "出力・表示", "VIEW": AddOverlay ws
        Case "図と概要"
          ws.range("A5").value2 = "変形・応力図は最終保存状態です。途中工程の値は「工程別数値」などで確認。ALL保存を指定すると各工程の節点・要素値を残せます。水理図は入力水位と計算水頭を区別します。"
          ws.range("A8:M8").Merge: ws.range("A8").value2 = "上部ボタンで描画します。表示する量・倍率・増分図のON/OFFは「表示設定」で選びます。"
        Case Else
          Select Case section
            Case "工程別数値": name = "ステージ結果"
            Case "節点結果": name = "結果節点"
            Case "要素結果": name = "結果要素"
            Case "水圧結果": name = "水圧結果"
            Case "接合結果": name = "接合結果"
            Case Else: name = "診断"
          End Select
          Set src = Nothing
          On Error Resume Next
          Set src = ThisWorkbook.Worksheets(name)
          On Error GoTo 0
          If src Is Nothing Then
            ws.range("A8").value2 = "この結果表はまだ作成されていません。"
          Else
            row0 = 2 + (PageNumber(ws) - 1) * 200
            Set a = src.cells(row0, 1).Resize(200, 13)
            Table ws, name, a.address, 1, "参照専用。L4のページ番号で200行ずつ表示（元表 " & CStr(row0) & "行～）。13列以降は「元の結果表を開く」で確認。", False
          End If
      End Select
  End Select
  If ws.name = "結果" Then
    CacheMeta ws.name, "DIRTY"
    If ThisWorkbook.names("_FEMViewerResultsDirty").RefersTo = "=TRUE" Then ws.range("A5").value2 = "【再解析が必要】保存された数値表は現在の入力に対応していません。 " & CStr(ws.range("A5").value2)
  End If
  ws.range("A7:M7").WrapText = True: ws.rows(7).RowHeight = 34
  ws.Protect UserInterfaceOnly:=True, DrawingObjects:=False
End Sub

Private Function ResultState() As String
  If ThisWorkbook.names("_FEMViewerResultsDirty").RefersTo = "=TRUE" Then
    ResultState = "未解析または入力変更あり：再解析が必要"
  Else
    ResultState = "保存結果あり"
  End If
End Function

Private Function PageNumber(ByVal ws As Worksheet) As Long
  PageNumber = 1
  If IsNumeric(ws.range("L4").value2) Then
    If CDbl(ws.range("L4").value2) >= 1 And CDbl(ws.range("L4").value2) <= 5000 Then PageNumber = CLng(ws.range("L4").value2)
  End If
End Function

Private Sub AddOverlay(ByVal ws As Worksheet)
  Dim i As Long, c As Object, r As Long, src As Worksheet, anchorRow As Long, modeRow As Long, crossRow As Long
  Set src = ThisWorkbook.Worksheets("出力・表示")
  For i = ws.CheckBoxes.count To 1 Step -1
    ws.CheckBoxes(i).Delete
  Next i
  anchorRow = ws.cells(ws.rows.count, 14).End(xlUp).row + 1
  For i = 0 To 1
    r = 28 + i
    Set c = ws.CheckBoxes.Add(ws.cells(anchorRow + i, 1).left, ws.cells(anchorRow + i, 1).Top, 360, 20)
    c.caption = Choose(i + 1, "最大せん断ひずみ増分を追加描画", "塑性主ひずみ増分クロスを追加描画")
    c.name = "UI24Overlay" & CStr(i): c.value = IIf(CBool(src.cells(r, 10).value2), 1, -4146)
    c.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!FEM24Overlay"
  Next i
  modeRow = anchorRow + 2: crossRow = anchorRow + 3
  ws.cells(modeRow, 1).Resize(1, 3).Merge: ws.cells(modeRow, 1).value2 = "増分の描き方 A/B"
  ws.cells(modeRow, 4).Resize(1, 2).Merge: ws.cells(modeRow, 4).value2 = src.range("C27").value2
  With ws.cells(modeRow, 4).Validation
    .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "A：選択結果へ等値線を追加,B：増分専用カラー"
    .InCellDropdown = True: .ShowError = True
  End With
  ws.cells(modeRow, 4).WrapText = True: ws.rows(modeRow).RowHeight = 40
  ws.cells(modeRow, 14).value2 = "出力・表示": ws.cells(modeRow, 15).value2 = 27
  ws.cells(crossRow, 1).Resize(1, 3).Merge: ws.cells(crossRow, 1).value2 = "クロス倍率"
  ws.cells(crossRow, 4).Resize(1, 2).Merge: ws.cells(crossRow, 4).value2 = src.range("C30").value2
  ws.cells(crossRow, 14).value2 = "出力・表示": ws.cells(crossRow, 15).value2 = 30
  ws.cells(modeRow, 4).Resize(2, 2).locked = False
  ws.cells(modeRow, 4).Resize(2, 2).Interior.Color = RGB(235, 245, 255)
  ws.cells(modeRow, 6).Resize(1, 8).Merge: ws.cells(modeRow, 6).value2 = "A=等値線を追加 / B=増分カラーへ切替"
End Sub

Public Sub FEM24Overlay()
  Dim ws As Worksheet, i As Long, e As Boolean
  e = Application.EnableEvents: Application.EnableEvents = False
  On Error GoTo Finish
  Set ws = ThisWorkbook.Worksheets("結果")
  For i = 0 To 1
    ThisWorkbook.Worksheets("出力・表示").cells(28 + i, 10).value2 = (ws.CheckBoxes("UI24Overlay" & CStr(i)).value = 1)
  Next i
  FEMInvalidateSettingCache
  CacheRenew "結果"
Finish:
  Application.EnableEvents = e
End Sub

Public Sub FEM24Refresh(ByVal sh As Object)
  Dim e As Boolean, screen As Boolean
  If mBusy Or AnalysisRunning Or Not FEM24IsPage(sh.name) Then Exit Sub
  If Application.CutCopyMode <> False Then
    If Not FEM24PageCurrent(sh.name) Then
      FEM24Error = "元の入力が更新されています。Escでコピーを解除し、この画面を開き直してください。"
      If Not SuppressUserMessages Then MsgBox FEM24Error, vbInformation
    End If
    Exit Sub
  End If
  e = Application.EnableEvents: screen = Application.ScreenUpdating
  mBusy = True: Application.EnableEvents = False: Application.ScreenUpdating = False
  FEM24Error = ""
  On Error GoTo Failed
  Render sh
Finish:
  mBusy = False: Application.EnableEvents = e: Application.ScreenUpdating = screen
  Exit Sub
Failed:
  FEM24Error = Err.description
  Resume Finish
End Sub

Public Sub FEM24RefreshStages()
  FEM24Refresh ThisWorkbook.Worksheets("施工工程")
End Sub

Public Sub FEM24Changed(ByVal sh As Object, ByVal target As range)
  Dim e As Boolean, hit As range, src As Worksheet, a As range, old As Variant, c As range, r As Long, backups As Object, key As Variant, saved As Variant, screen As Boolean
  If mBusy Or AnalysisRunning Then Exit Sub
  e = Application.EnableEvents: screen = Application.ScreenUpdating
  Application.EnableEvents = False: Application.ScreenUpdating = False: mBusy = True: FEM24Error = ""
  On Error GoTo Failed
  If Not Intersect(target, sh.range("C4:F4,L4")) Is Nothing Then
    Render sh: GoTo Finish
  End If
  If Not FEM24PageCurrent(sh.name) Then Err.Raise vbObjectError + 2447, , "元の入力が更新されたため、画面を更新しました。内容を確認して入力し直してください。"
  If sh.name = "施工工程" Then
    FEM24StagesChanged sh, target
    FEM24Error = FEM24StageError: GoTo Finish
  End If
  If Len(CStr(sh.range("N1").value2)) > 0 Then
    Set a = sh.range(CStr(sh.range("N3").value2))
    Set hit = Intersect(target, a)
    If hit Is Nothing Then GoTo Finish
    Set src = ThisWorkbook.Worksheets(CStr(sh.range("N1").value2))
    Set a = src.range(CStr(sh.range("N2").value2)): old = FEM24CaptureCells(a)
    For Each c In hit.cells
      If c.HasFormula Then Err.Raise vbObjectError + 2441, , "入力には数値または文字を指定してください。計算式は使用できません。"
      src.cells(a.row + c.row - 8, a.column + c.column - 1).value2 = c.value2
    Next c
    FEM24Invalidate
  Else
    Set hit = Intersect(target, sh.range("D8:D2007"))
    If Not hit Is Nothing Then
      Set backups = CreateObject("Scripting.Dictionary")
      For Each c In hit.cells
        r = CLng(val(CStr(sh.cells(c.row, 15).value2)))
        If r > 0 Then
          Set src = ThisWorkbook.Worksheets(CStr(sh.cells(c.row, 14).value2))
          key = src.name & "|" & CStr(r)
          If Not backups.Exists(key) Then backups.Add key, Array(src.name, r, FEM24CaptureCells(src.cells(r, 3)))
          If c.HasFormula Then Err.Raise vbObjectError + 2442, , "設定値に計算式は使用できません。"
          If src.name = "出力・表示" And r = 27 Then
            Select Case CStr(c.value2)
              Case "A", "A：選択結果へ等値線を追加": src.cells(r, 3).value2 = "A：選択結果へ等値線を追加"
              Case "B", "B：増分専用カラー": src.cells(r, 3).value2 = "B：増分専用カラー"
              Case Else: Err.Raise vbObjectError + 2448, , "増分の描き方はAまたはBを選択してください。"
            End Select
          Else
            src.cells(r, 3).value2 = c.value2
          End If
          FEMUiSettingsChanged src, src.cells(r, 3)
        End If
      Next c
      FEMUiCalculateSettings
      FEM24ClearPlot
    End If
  End If
  GoTo Finish
Failed:
  FEM24Error = Err.description
  If Not a Is Nothing Then
    If Not IsEmpty(old) Then Call FEM24RestoreCells(a, old)
  End If
  If Not backups Is Nothing Then
    For Each key In backups.keys
      saved = backups(key)
      Call FEM24RestoreCells(ThisWorkbook.Worksheets(CStr(saved(0))).cells(CLng(saved(1)), 3), saved(2))
    Next key
    FEMUiCalculateSettings
  End If
  Render sh
  If Not SuppressUserMessages Then MsgBox FEM24Error, vbExclamation
Finish:
  CacheRenew sh.name
  mBusy = False: Application.EnableEvents = e: Application.ScreenUpdating = screen
End Sub

Public Sub FEM24Invalidate()
  Call FEMViewerMarkDirty
  Call FEMUiInvalidateMeshPreview
  AnalysisOK = False: ResultRevision = 0: P1ResultReady = False
  Call FEMInvalidateSettingCache
  FEM24ClearPlot
End Sub

Public Sub FEM24ClearPlot()
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("結果")
  ws.Unprotect: ws.Shapes("UI24Plot").Delete
  ws.Protect UserInterfaceOnly:=True, DrawingObjects:=False
End Sub

Private Sub GoPage(ByVal page As String)
  Dim e As Boolean, screen As Boolean
  If AnalysisRunning Then Exit Sub
  If Application.CutCopyMode <> False Then
    If Not FEM24PageCurrent(page) Then
      FEM24Error = "元の入力が更新されています。Escでコピーを解除し、この画面を開き直してください。"
      If Not SuppressUserMessages Then MsgBox FEM24Error, vbInformation
      Exit Sub
    End If
    If ThisWorkbook.Worksheets(page).Visible <> xlSheetVisible Then ThisWorkbook.Worksheets(page).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(page).Activate
    Exit Sub
  End If
  e = Application.EnableEvents: screen = Application.ScreenUpdating
  On Error GoTo Failed
  Application.EnableEvents = False: Application.ScreenUpdating = False
  FEM24ApplyTabs
  With ThisWorkbook.Worksheets(page)
    .Activate: FEM24Refresh ThisWorkbook.Worksheets(page)
    Application.Goto .range("A1"), True
  End With
  ActiveWindow.Zoom = 85: ActiveWindow.DisplayGridlines = False
  ActiveWindow.FreezePanes = False: ActiveWindow.SplitColumn = 0: ActiveWindow.SplitRow = 7: ActiveWindow.FreezePanes = True
  Application.EnableEvents = e: Application.ScreenUpdating = screen
  Exit Sub
Failed:
  FEM24Error = Err.description
  Application.EnableEvents = e: Application.ScreenUpdating = screen
  If Not SuppressUserMessages Then MsgBox FEM24Error, vbExclamation
End Sub
Public Sub FEM24Home(): GoPage "ホーム": End Sub
Public Sub FEM24Model(): GoPage "モデル": End Sub
Public Sub FEM24Materials(): GoPage "材料・接合": End Sub
Public Sub FEM24Water(): GoPage "境界・水条件": End Sub
Public Sub FEM24Stages(): GoPage "施工工程": End Sub
Public Sub FEM24Analysis(): GoPage "解析・保存": End Sub
Public Sub FEM24Results(): GoPage "結果": End Sub
Public Sub FEM24Save(): ThisWorkbook.Save: End Sub

Public Sub FEM24Generate()
  If AnalysisRunning Then Exit Sub
  On Error GoTo Failed
  FEMUiGenerateShape
  If Not FEMShapeGenerationOK Or Len(FEMUiMeshPreviewError) > 0 Then Err.Raise vbObjectError + 2446, , "メッシュを作成できませんでした。" & FEMUiMeshPreviewError
  PlotFront True
  Exit Sub
Failed:
  FEM24Error = Err.description
  If Not SuppressUserMessages Then MsgBox FEM24Error, vbExclamation
End Sub

Private Sub PlotFront(ByVal preview As Boolean)
  Dim ws As Worksheet, s As shape, e As Boolean, screen As Boolean, errorNo As Long, errorText As String
  e = Application.EnableEvents: screen = Application.ScreenUpdating
  On Error GoTo Failed
  Application.EnableEvents = False: Application.ScreenUpdating = False
  If Len(FEMViewerLastError) > 0 Then Err.Raise vbObjectError + 2443, , FEMViewerLastError
  If Len(FEMViewerLastSvg) = 0 Then Err.Raise vbObjectError + 2444, , "描画データがありません。"
  FEM24Results
  Set ws = ThisWorkbook.Worksheets("結果")
  ws.Unprotect: ws.range("C4").value2 = "図と概要": Render ws: ws.Unprotect
  Call FEM24ClearPlot
  ws.Unprotect
  Set s = FEMNativeFigureCreate(ws, FEMViewerLastSvg, FEMViewerCanvasW, FEMViewerCanvasH)
  s.name = "UI24Plot": s.LockAspectRatio = msoTrue
  s.Placement = xlFreeFloating
  If s.width > 905 Then s.width = 905
  s.left = 5: s.Top = ws.rows(9).Top + 2
  ws.range("A8").value2 = IIf(preview, "解析前メッシュ（確認済み）", "最終保存状態の変形・応力図")
  ws.Protect UserInterfaceOnly:=True, DrawingObjects:=False
  ActiveWindow.Zoom = 70
  Application.EnableEvents = e: Application.ScreenUpdating = screen
  Exit Sub
Failed:
  errorNo = Err.number: errorText = Err.description
  Application.EnableEvents = e: Application.ScreenUpdating = screen
  Err.Raise errorNo, "PlotFront", errorText
End Sub

Public Sub FEM24Preview()
  If AnalysisRunning Then Exit Sub
  On Error GoTo Failed
  FEMViewerLastSvg = ""
  FEMUiPreviewMesh
  If Len(FEMUiMeshPreviewError) > 0 Then Err.Raise vbObjectError + 2445, , FEMUiMeshPreviewError
  PlotFront True
  Exit Sub
Failed:
  FEM24Error = Err.description: FEM24ClearPlot
  If Not SuppressUserMessages Then MsgBox FEM24Error, vbExclamation
End Sub
Public Sub FEM24Plot()
  If AnalysisRunning Then Exit Sub
  On Error GoTo Failed
  FEMViewResult
  PlotFront False
  Exit Sub
Failed:
  FEM24Error = Err.description: FEM24ClearPlot
  If Not SuppressUserMessages Then MsgBox FEM24Error, vbExclamation
End Sub
Public Sub FEM24Run()
  If AnalysisRunning Then Exit Sub
  On Error GoTo Failed
  FEMViewerLastSvg = ""
  FEMUiStartFromMesh
  If Len(FEMViewerLastSvg) > 0 Then
    PlotFront True
  Else
    FEM24Results
    ThisWorkbook.Worksheets("結果").Unprotect
    ThisWorkbook.Worksheets("結果").range("A8").value2 = ResultStatus & "  " & AnalysisMessage
    ThisWorkbook.Worksheets("結果").Protect UserInterfaceOnly:=True, DrawingObjects:=False
  End If
  Exit Sub
Failed:
  FEM24Error = Err.description
  If Not SuppressUserMessages Then MsgBox FEM24Error, vbExclamation
End Sub

Public Sub FEM24WaterPreview()
  If AnalysisRunning Then Exit Sub
  On Error GoTo Failed
  FEMWaterPreview
  FEM24ReturnLink ThisWorkbook.ActiveSheet
  Exit Sub
Failed:
  FEM24Error = Err.description
  If Not SuppressUserMessages Then MsgBox FEM24Error, vbExclamation
End Sub

Public Sub FEM24WaterFigure()
  On Error GoTo Failed
  FEMWaterResults
  FEM24ReturnLink ThisWorkbook.ActiveSheet
  Exit Sub
Failed:
  FEM24Error = Err.description
  If Not SuppressUserMessages Then MsgBox FEM24Error, vbExclamation
End Sub

Private Sub FEM24ReturnLink(ByVal ws As Worksheet)
  Dim wasProtected As Boolean
  On Error Resume Next
  wasProtected = ws.ProtectContents
  ws.Unprotect: ws.Shapes("UI24_Return").Delete
  Button ws, "Return", "7画面へ戻る", "FEM24Home", ws.columns("V").left, 5, 140
  If wasProtected Then ws.Protect UserInterfaceOnly:=True, DrawingObjects:=False
End Sub

Public Sub FEM24Detail()
  Dim ws As Worksheet, name As String, section As String
  Set ws = ThisWorkbook.ActiveSheet: section = CStr(ws.range("C4").value2)
  name = CStr(ws.range("N1").value2)
  If name = "" Then
    Select Case ws.name
      Case "ホーム": name = "操作ガイド"
      Case "モデル": name = "メッシュ設定"
      Case "施工工程": name = "ステージ"
      Case "結果"
        Select Case section
          Case "節点結果": name = "結果節点"
          Case "要素結果": name = "結果要素"
          Case "工程別数値": name = "ステージ結果"
          Case "水圧結果": name = "水圧結果"
          Case "接合結果": name = "接合結果"
          Case "診断": name = "診断"
          Case Else: name = "図"
        End Select
      Case Else: name = "操作ガイド"
    End Select
  End If
  On Error GoTo Failed
  Set ws = ThisWorkbook.Worksheets(name)
  ws.Visible = xlSheetVisible: ws.Activate: Application.Goto ws.range("A1"), True
  FEM24ReturnLink ws
  Exit Sub
Failed:
  FEM24Error = Err.description
  If Not SuppressUserMessages Then MsgBox "該当する表はまだありません。" & vbCrLf & FEM24Error, vbInformation
End Sub
