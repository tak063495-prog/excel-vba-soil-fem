Option Explicit

' UI migration is additive and idempotent: existing model/settings values win.
Public Sub FEMUiPrepareInputs()
  Dim ws As Worksheet, legacy As Worksheet, p As Variant, r As Long, k As String
  Dim keys As Variant, item As Variant, i As Long, targetRow As Long, legacyRow As Long, found As Boolean
  Set ws = ThisWorkbook.Worksheets("高速化設定")
  Set legacy = ThisWorkbook.Worksheets("設定")
  keys = Array("NATIVE_DLL", "NATIVE_PRECISION", "NATIVE_STATE_COPY", "NATIVE_ASSEMBLY", "NATIVE_GMRES_SIMD", "NATIVE_GMRES_TUNED", "NATIVE_GMRES_PREPARED", "IO_FAST_PATH", "EVENT_LOG_FAST")
  For Each item In keys
    found = False
    For Each p In Array("メッシュ設定", "解析設定", "高速化設定", "出力・表示", "詳細設定")
      With ThisWorkbook.Worksheets(CStr(p))
        For r = 1 To .Cells(.rows.count, 6).End(xlUp).row
          If UCase$(Trim$(CStr(.Cells(r, 6).value2))) = CStr(item) Then found = True: Exit For
        Next r
      End With
      If found Then Exit For
    Next p
    If Not found Then
      legacyRow = 0
      For r = 5 To legacy.Cells(legacy.rows.count, 5).End(xlUp).row
        If UCase$(Trim$(CStr(legacy.Cells(r, 5).value2))) = CStr(item) Then legacyRow = r: Exit For
      Next r
      If legacyRow = 0 Then Err.Raise vbObjectError + 3830, "FEMUiPrepareInputs", "設定キーがありません: " & CStr(item)
      targetRow = ws.Cells(ws.rows.count, 6).End(xlUp).row + 1
      If targetRow < 33 Then targetRow = 33
      ws.Cells(targetRow, 2).value2 = legacy.Cells(legacyRow, 2).value2
      ws.Cells(targetRow, 3).value2 = legacy.Cells(legacyRow, 3).value2
      ws.Cells(targetRow, 5).value2 = legacy.Cells(legacyRow, 4).value2
      ws.Cells(targetRow, 6).value2 = CStr(item)
    End If
  Next item
  FEMUiOrderRuntimePage ws, keys
  For Each p In Array("メッシュ設定", "解析設定", "高速化設定", "出力・表示", "詳細設定")
    Set ws = ThisWorkbook.Worksheets(CStr(p))
    ws.columns("G").Hidden = False: ws.columns("G").ColumnWidth = 20
    ws.range("G7").value2 = "推奨・目安"
    ws.range("G7").Interior.Color = RGB(51, 69, 91): ws.range("G7").Font.Color = vbWhite
    ws.range("G7").Font.bOld = True
    For r = 8 To ws.Cells(ws.rows.count, 6).End(xlUp).row
      k = UCase$(Trim$(CStr(ws.Cells(r, 6).value2)))
      If Len(k) > 0 Then
        ws.Cells(r, 7).value2 = FEMUiRecommendation(k)
        ws.Cells(r, 7).Font.Color = RGB(45, 100, 85)
        ws.Cells(r, 7).Interior.Color = RGB(238, 245, 242)
        ws.Cells(r, 7).WrapText = True
        ws.Cells(r, 3).Interior.Color = RGB(255, 242, 204)
        ws.Cells(r, 5).WrapText = True
        If ws.rows(r).RowHeight < 48 Then ws.rows(r).RowHeight = 48
        If left$(k, 14) = "MESH_GEN_LAST_" Then
          ws.Cells(r, 7).value2 = "自動出力"
          ws.Cells(r, 3).Interior.Color = RGB(231, 237, 244)
        End If
      End If
    Next r
    ws.range("B7:G7").RowHeight = 28
    ws.PageSetup.PrintArea = "$B$2:$G$" & CStr(Application.Max(24, ws.Cells(ws.rows.count, 6).End(xlUp).row))
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False: ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.FitToPagesTall = False
  Next p
  With ThisWorkbook.Worksheets("解析設定")
    .range("E10").value2 = "INCONSISTENT=非関連流れ・非対称整合接線・全Newton / DAVIS=等価関連c*,φ*。JOINTには掛けない。"
    .range("E16").value2 = "Fs区間の停止幅。現値を維持。参考：通常0.025 / 詳細検証0.0125。小さくすると試行が増えます。"
    .range("E17").value2 = "材料のψの扱い。CAP=φ'以下 / REDUCE=Fsで低減 / KEEP=入力のまま。物理条件なので一括変更しません。"
    .range("B5").value2 = "黄色=C列の入力、緑=推奨・目安。SRM許容幅・ψは解析条件に合わせ、現値を確認してください。"
  End With
  With ThisWorkbook.Worksheets("高速化設定")
    .range("B5").value2 = "通常はDLL・I/OをAUTO/1。下段V方式は比較用。未対応経路は自動復帰。収束条件は変更しません。"
  End With
End Sub

Private Sub FEMUiOrderRuntimePage(ByVal ws As Worksheet, ByVal runtimeKeys As Variant)
  Dim entries As Object, k As Variant, r As Long, rowNo As Long, lastRow As Long, record As Variant
  Set entries = CreateObject("Scripting.Dictionary"): entries.CompareMode = vbTextCompare
  lastRow = ws.Cells(ws.rows.count, 6).End(xlUp).row
  For r = 8 To lastRow
    k = Trim$(CStr(ws.Cells(r, 6).value2))
    If Len(k) > 0 Then
      If entries.Exists(CStr(k)) Then Err.Raise vbObjectError + 3831, "FEMUiOrderRuntimePage", "設定キーが重複しています: " & CStr(k)
      entries.Add CStr(k), Array(ws.Cells(r, 2).value2, ws.Cells(r, 3).Formula, ws.Cells(r, 5).value2)
    End If
  Next r
  ws.range("B8:G" & CStr(Application.Max(43, lastRow))).UnMerge
  ws.range("B8:G" & CStr(Application.Max(43, lastRow))).ClearContents
  ws.range("B8:G" & CStr(Application.Max(43, lastRow))).Interior.Color = vbWhite
  ws.range("B8:G" & CStr(Application.Max(43, lastRow))).Font.bOld = False
  ws.range("B8:G" & CStr(Application.Max(43, lastRow))).Font.Color = RGB(36, 52, 71)
  ws.range("B8:E8").Merge: ws.range("B8").value2 = "通常使用：DLL・I/O（倍精度・収束条件を維持）"
  ws.range("B8:E8").Interior.Color = RGB(231, 237, 244)
  rowNo = 9
  For Each k In runtimeKeys
    ' A customized older workbook may keep a key on another canonical page.
    ' Preserve that location and value; the settings bridge still finds it.
    If entries.Exists(CStr(k)) Then
      record = entries(CStr(k))
      FEMUiRuntimeRow ws, rowNo, CStr(k), record
      entries.Remove CStr(k)
    End If
  Next k
  rowNo = rowNo + 1
  ws.range(ws.Cells(rowNo, 2), ws.Cells(rowNo, 5)).Merge
  ws.Cells(rowNo, 2).value2 = "詳細・比較用：V方式（通常は現値を維持）"
  ws.Cells(rowNo, 2).Interior.Color = RGB(231, 237, 244)
  rowNo = rowNo + 1
  For Each k In entries.Keys
    record = entries(CStr(k))
    FEMUiRuntimeRow ws, rowNo, CStr(k), record
  Next k
End Sub

Private Sub FEMUiRuntimeRow(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal k As String, ByVal record As Variant)
  Dim legacy As Worksheet, r As Long
  Set legacy = ThisWorkbook.Worksheets("設定")
  ws.Cells(rowNo, 2).value2 = record(0)
  ws.Cells(rowNo, 3).NumberFormat = "General": ws.Cells(rowNo, 3).Formula = record(1)
  ws.Cells(rowNo, 5).value2 = record(2): ws.Cells(rowNo, 6).value2 = k
  For r = 5 To legacy.Cells(legacy.rows.count, 5).End(xlUp).row
    If CStr(legacy.Cells(r, 5).value2) = k Then
      ws.Cells(rowNo, 5).value2 = legacy.Cells(r, 4).value2: Exit For
    End If
  Next r
  rowNo = rowNo + 1
End Sub

Private Function FEMUiRecommendation(ByVal keyName As String) As Variant
  Select Case keyName
    Case "NATIVE_DLL": FEMUiRecommendation = "AUTO"
    Case "NATIVE_PRECISION": FEMUiRecommendation = "AUTO"
    Case "NATIVE_STATE_COPY": FEMUiRecommendation = 1
    Case "NATIVE_ASSEMBLY": FEMUiRecommendation = 1
    Case "NATIVE_GMRES_SIMD": FEMUiRecommendation = 1
    Case "NATIVE_GMRES_TUNED": FEMUiRecommendation = 1
    Case "NATIVE_GMRES_PREPARED": FEMUiRecommendation = 1
    Case "IO_FAST_PATH": FEMUiRecommendation = 1
    Case "EVENT_LOG_FAST": FEMUiRecommendation = 1
    Case "FLOW_POLICY": FEMUiRecommendation = "INCONSISTENT"
    Case "WEIGHT_MODE": FEMUiRecommendation = "TOTAL"
    Case "Q8_HOURGLASS_FACTOR": FEMUiRecommendation = 0.05
    Case "RCM_POLICY": FEMUiRecommendation = "AUTO"
    Case "SRM_FMAX": FEMUiRecommendation = "モデル依存"
    Case "SRM_TOL": FEMUiRecommendation = "現値を維持"
    Case "SRM_PSI_POLICY": FEMUiRecommendation = "モデル依存"
    Case "SRM_MODE": FEMUiRecommendation = "REAPPLY"
    Case "SRM_FIXED_FS": FEMUiRecommendation = 0
    Case "OUTPUT_STAGE_MODE": FEMUiRecommendation = "FINAL"
    Case "OUTPUT_INACTIVE_ELEMENTS": FEMUiRecommendation = "SKIP"
    Case "DEBUG_MODE": FEMUiRecommendation = "STAGE"
    Case "DEBUG_FLUSH": FEMUiRecommendation = 5
    Case "EXPORT_MODE": FEMUiRecommendation = "STAGE"
    Case "EXPORT_LOAD": FEMUiRecommendation = "空欄（新規）"
    Case "SOLVER": FEMUiRecommendation = "BAND_LU"
    Case "BAND_LU_KERNEL": FEMUiRecommendation = "NEW"
    Case "PHYSICAL_FAILURE_MODE": FEMUiRecommendation = "固定・未使用"
    Case "VIEW_DETAIL_MODE": FEMUiRecommendation = "SIMPLE"
    Case "VIEW_RESULT_SOURCE": FEMUiRecommendation = "SMOOTHED"
    Case "VIEW_SCALE_MODE": FEMUiRecommendation = "SYMMETRIC"
    Case "ACCEL_STEP_RECOVERY": FEMUiRecommendation = 0
    Case "ACCEL_STEP_FRESH_LU": FEMUiRecommendation = 0
    Case "ACCEL_TRACE": FEMUiRecommendation = 0
    Case "ACCEL_V1_PREDICTOR": FEMUiRecommendation = 1
    Case "ACCEL_V2A_REUSE": FEMUiRecommendation = 1
    Case "ACCEL_V2B_COST": FEMUiRecommendation = 0
    Case "ACCEL_V3_COST_SEARCH": FEMUiRecommendation = 0
    Case "ACCEL_V4_ANDERSON": FEMUiRecommendation = 1
    Case "ACCEL_V5_GMRES_LU": FEMUiRecommendation = 0
    Case "ACCEL_ADAPTIVE": FEMUiRecommendation = 1
    Case "ACCEL_PREDICT_BETA": FEMUiRecommendation = 1
    Case "ACCEL_PREDICT_MAX_RATIO": FEMUiRecommendation = 1.5
    Case "ACCEL_TANGENT_CHANGE": FEMUiRecommendation = 0.05
    Case "MESH_QUALITY_MIN_AREA_RATIO": FEMUiRecommendation = 0.000001
    Case "MESH_QUALITY_MIN_JACOBIAN_RATIO": FEMUiRecommendation = 0.000001
    Case "MESH_QUALITY_MAX_ASPECT_RATIO": FEMUiRecommendation = 5#
    Case "MESH_REPAIR_MAX_ITERATIONS": FEMUiRecommendation = 5#
    Case "MESH_REPAIR_DAMPING": FEMUiRecommendation = 0.5
    Case "MESH_REPAIR_IMPROVEMENT_TOLERANCE": FEMUiRecommendation = 0.0001
    Case "MESH_REPAIR_ANGLE_ITERATIONS": FEMUiRecommendation = 2000#
    Case "MESH_REPAIR_TERRAIN_COMPARE": FEMUiRecommendation = 1#
    Case "MESH_TERRAIN_COUNT_MIN": FEMUiRecommendation = 0.55
    Case "MESH_TERRAIN_COUNT_MAX": FEMUiRecommendation = 1.8
    Case "ADAPT_MAX_ELEMENT_RATIO": FEMUiRecommendation = 2#
    Case "ADAPT_MARK_FRACTION": FEMUiRecommendation = 0.15
    Case "MESH_QUALITY_MIN_ANGLE_DEG": FEMUiRecommendation = "品質目標による"
    Case "MESH_QUALITY_MAX_ANGLE_DEG": FEMUiRecommendation = "品質目標による"
    Case Else: FEMUiRecommendation = "モデル・表示による"
  End Select
End Function

' Only this named, user-invoked action changes configuration values.
' Never change geometry, supports, material law, SRM or convergence controls.
Public Sub FEMUiApplyRecommendedRuntime()
  Dim k As Variant
  FEMWriteTextSetting "NATIVE_DLL", "AUTO"
  FEMWriteTextSetting "NATIVE_PRECISION", "AUTO"
  For Each k In Array("NATIVE_STATE_COPY", "NATIVE_ASSEMBLY", "NATIVE_GMRES_SIMD", "NATIVE_GMRES_TUNED", "NATIVE_GMRES_PREPARED", "IO_FAST_PATH", "EVENT_LOG_FAST")
    FEMWriteNumericSetting CStr(k), 1#
  Next k
  AnalysisOK = False: ResultRevision = 0: P1ResultReady = False
  FEMViewerMarkDirty
End Sub

Public Sub FEMUiCheckSettings()
  On Error GoTo Failed
  FEMUiValidateSettings
  MsgBox "設定の形式・範囲は正常です。材料・メッシュ・ステージの整合性は解析開始時に検査します。", vbInformation
  Exit Sub
Failed:
  MsgBox Err.description, vbExclamation, "設定を確認してください"
End Sub

Public Sub FEMUiFinishInputs()
  Dim ws As Worksheet, shp As shape, target As range, item As Variant, rowNo As Long
  Set ws = ThisWorkbook.Worksheets("操作パネル")
  ws.range("B35").value2 = "計算版 " & FEM_BUILD_STAMP & " / 入出力画面 " & FEM_UI_STAMP
  ws.Hyperlinks.Add Anchor:=ws.range("B14"), address:="", SubAddress:="'操作ガイド'!B2", TextToDisplay:="初めに：対応機能・手順・結果の読み方"
  ws.range("B14").Font.size = 11
  ws.range("B37:H38").Merge: ws.range("B37").value2 = "結果を見る前に：入力を変えた場合は必ず再解析。配布時に残っている数値・図は前回計算の例で、現在の入力に対する結果ではありません。"
  ws.range("B37:H38").WrapText = True: ws.range("B37").Font.Color = RGB(140, 80, 30)
  ws.range("B40:H41").Merge: ws.range("B40").value2 = "単位は統一して入力（例：m・kN・kPa）。材料値・荷重・支持・Fs条件はモデルごとに確認。黄色が入力、灰色は結果・自動出力、緑は推奨・目安です。"
  ws.range("B40:H41").WrapText = True
  Set ws = ThisWorkbook.Worksheets("高速化設定")
  Set target = ws.range("B44:D45")
  Set shp = Nothing
  On Error Resume Next: Set shp = ws.Shapes("UI_RecommendedRuntime"): On Error GoTo 0
  If shp Is Nothing Then Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, target.left, target.Top, target.width, target.Height)
  shp.name = "UI_RecommendedRuntime": shp.TextFrame2.TextRange.text = "DLL・I/Oだけ推奨値に戻す"
  shp.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!FEMUiApplyRecommendedRuntime"
  shp.fill.ForeColor.RGB = RGB(49, 91, 79): shp.line.Visible = msoFalse
  shp.TextFrame2.TextRange.Font.size = 11: shp.TextFrame2.TextRange.Font.fill.ForeColor.RGB = vbWhite
  Set ws = ThisWorkbook.Worksheets("解析設定")
  Set target = ws.range("B22:D23"): Set shp = Nothing
  On Error Resume Next: Set shp = ws.Shapes("UI_CheckSettings"): On Error GoTo 0
  If shp Is Nothing Then Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, target.left, target.Top, target.width, target.Height)
  shp.name = "UI_CheckSettings": shp.TextFrame2.TextRange.text = "設定の形式・範囲をチェック"
  shp.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!FEMUiCheckSettings"
  shp.fill.ForeColor.RGB = RGB(49, 91, 79): shp.line.Visible = msoFalse
  shp.TextFrame2.TextRange.Font.size = 11: shp.TextFrame2.TextRange.Font.fill.ForeColor.RGB = vbWhite
  Set ws = ThisWorkbook.Worksheets("操作パネル")
  Set target = ws.range("B43:D44"): Set shp = Nothing
  On Error Resume Next: Set shp = ws.Shapes("UI_ElementResults"): On Error GoTo 0
  If shp Is Nothing Then Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, target.left, target.Top, target.width, target.Height)
  shp.name = "UI_ElementResults": shp.TextFrame2.TextRange.text = "要素の応力・ひずみ・塑性"
  shp.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!FEMUiShowElementResults"
  shp.fill.ForeColor.RGB = RGB(49, 91, 79): shp.line.Visible = msoFalse
  shp.TextFrame2.TextRange.Font.size = 11: shp.TextFrame2.TextRange.Font.fill.ForeColor.RGB = vbWhite
  For Each item In Array("材料データ", "形状入力", "接合", "ステージ", "載荷")
    ThisWorkbook.Worksheets(CStr(item)).Tab.Color = RGB(217, 170, 80)
  Next item
  For Each item In Array("メッシュ設定", "解析設定", "高速化設定", "出力・表示", "詳細設定")
    ThisWorkbook.Worksheets(CStr(item)).Tab.Color = RGB(81, 135, 115)
  Next item
  For Each item In Array("図", "ステージ結果", "接合結果", "結果節点", "結果要素", "診断")
    ThisWorkbook.Worksheets(CStr(item)).Tab.Color = RGB(85, 125, 170)
  Next item
  FEMUiWriteGuide
End Sub

Public Sub FEMUiShowElementResults()
  With ThisWorkbook.Worksheets("結果要素")
    .Visible = xlSheetVisible: .Activate
    Application.Goto .range("A1"), True
  End With
End Sub

Private Sub FEMUiWriteGuide()
  Dim ws As Worksheet, sections As Variant, i As Long, activeName As String
  activeName = ThisWorkbook.ActiveSheet.name
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("操作ガイド"): On Error GoTo 0
  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("操作パネル"))
    ws.name = "操作ガイド"
  End If
  ws.Unprotect
  ws.Cells.Font.name = "Yu Gothic": ws.Cells.Font.size = 11
  ws.range("B2:E42").UnMerge: ws.range("B2:E42").ClearContents
  ws.columns("A").ColumnWidth = 3: ws.columns("B").ColumnWidth = 26: ws.columns("C:E").ColumnWidth = 28
  ws.range("B2:E2").Merge: ws.range("B2").value2 = "操作ガイド / できること・入出力"
  ws.range("B2").Font.size = 20: ws.range("B2").Font.bOld = True
  ws.Hyperlinks.Add Anchor:=ws.range("B4"), address:="", SubAddress:="'操作パネル'!B2", TextToDisplay:="操作パネルへ戻る"
  sections = Array( _
    Array("1 材料・単位", "材料データ：E、ν、γ、c、φ、ψとSOIL/STRUCT/JOINTを確認。全シートで単位を統一（例 m・kN・kPa、角度deg）。物性・荷重に汎用の推奨値はありません。"), _
    Array("2 メッシュ・支持", "メッシュ設定：簡易作成①は3/4点＋nx/ny。形状作成②は形状入力の座標・境界を使用。底面・側面支持はモデル依存。節点/要素を直接入力する操作も可能。"), _
    Array("3 接合（任意）", "接合：材料A/Bの共有辺にJOINTを生成。材料種別JOINT、Kn/Ks、psi=0。履歴なし応力上限モデルで、塑性すべり・除荷履歴は未対応。"), _
    Array("4 ステージ・載荷", "ステージは上から有効=1だけ実行。GRAVITY/LOAD/UNLOAD/BIRTH/DEATH/SRM/RESET_U/RESET_STRESS/MATSET。LOAD対象・BOX範囲・符号は載荷で指定。"), _
    Array("5 解析・Fs", "2D平面ひずみ・小ひずみ・全応力。INCONSISTENTまたはDAVIS。最終SRMはFs区間探索、途中SRMは必要Fs判定。Fs許容幅・ψ方針は物理/評価条件として現値を維持。"), _
    Array("6 設定・実行", "黄色C列が正規入力、緑G列が目安。設定チェック後、操作パネルの解析を実行。入力規則を回避した貼り付けも解析入口で検査。材料・メッシュの整合性も既存検証で確認。"), _
    Array("推奨設定の範囲", "配布版はDLL=AUTO・演算=AUTO・DLL/I/O高速化=1を入力済み。すべて倍精度。ボタンはこの9項目だけを復元し、モデル・Fs・収束条件は変更しません。既存ブックの設定は再表示で上書きしません。"), _
    Array("結果の判定", "ステージ結果：成功/失敗・Fs PASS/FAIL区間を先に確認。点推定だけで安全率を断定しない。診断：残差・収束・停止理由。入力変更後は古い数値や図を現在の結果として使わず再解析。"), _
    Array("変位・応力・図", "結果節点：全節点変位。結果要素：応力・ひずみ・塑性。接合結果：開口・すべり状態。図：表示量一覧の番号で選択。変位倍率と平滑化の有無を併記して読みます。"), _
    Array("保存・再開", "出力・表示：保存先はブック横の「ブック名_out」。run.log・solver_events.csv・ステージcheckpoint。EXPORT_LOAD空欄=新規、AUTO/フォルダ名=再開。入力署名と必要ファイルを照合。"), _
    Array("品質修正・再メッシュ", "メッシュ品質修正、塑性域を使った再メッシュ候補の生成と採用が可能。採用後は再解析。品質角度などは形状に合わせる目標値で、材料の収束条件とは別です。"), _
    Array("対象外・注意", "浮力・水位・連成圧密は未対応（有効化は拒否）。JOINTのpsi≠0や履歴塑性すべりも未対応。既定値は設計の妥当性や収束を保証しません。"))
  For i = LBound(sections) To UBound(sections)
    ws.Cells(6 + i * 2, 2).value2 = sections(i)(0)
    ws.range(ws.Cells(6 + i * 2, 3), ws.Cells(7 + i * 2, 5)).Merge
    ws.Cells(6 + i * 2, 3).value2 = sections(i)(1)
    ws.rows(CStr(6 + i * 2) & ":" & CStr(7 + i * 2)).RowHeight = 40
    ws.Cells(6 + i * 2, 2).Font.bOld = True
    ws.Cells(6 + i * 2, 2).Interior.Color = RGB(231, 237, 244)
  Next i
  ws.range("B6:E29").WrapText = True: ws.range("B6:E29").VerticalAlignment = xlTop
  ws.range("B6:E29").Font.size = 11
  ws.Visible = xlSheetVisible: ws.Move After:=ThisWorkbook.Worksheets("操作パネル")
  ws.Protect UserInterfaceOnly:=True
  ws.PageSetup.PrintArea = "$B$2:$E$29": ws.PageSetup.Zoom = False
  ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.FitToPagesTall = 1
  ws.Activate: ActiveWindow.DisplayGridlines = False: ActiveWindow.Zoom = 85
  ActiveWindow.FreezePanes = False: ws.range("C6").Select: ActiveWindow.FreezePanes = True
  ThisWorkbook.Worksheets(activeName).Activate
End Sub


Public Function FEMUiRefreshInputs() As String
  On Error GoTo Failed
  FEMUiInvalidateLayout
  FEMWriteBuildStamp False
  FEMUiRefreshInputs = "PASS": Exit Function
Failed:
  FEMUiRefreshInputs = "ERROR|" & CStr(Err.Number) & "|" & Err.description
End Function
