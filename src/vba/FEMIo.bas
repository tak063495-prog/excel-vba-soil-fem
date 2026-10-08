Option Explicit
' Workbook layout, run.log, checkpoint export/resume, V0 performance log.
' Limit-state shadow for SRM only. Does not stop the analysis and does not change numerical PASS/FAIL.
Private Const PHYS_REF_N As Long = 3
Private Const PHYS_WARN_RATIO As Double = 5#
Private Const PHYS_TRIP_RATIO As Double = 10#
Private Const PHYS_TRIP_CONSEC As Long = 2
Private Const LIMIT_WARN_RATIO As Double = 5#
Private Const LIMIT_TRIP_RATIO As Double = 10#
Private Const LIMIT_TRIP_COUNT As Long = 2
Private Const PHYS_DLAM_EPS As Double = 0.000001
Private Const PHYS_MIN_DU20_H As Double = 0.0005
Private mPhysRefVals(1 To PHYS_REF_N) As Double
Private mPhysRefCount As Long
Private mPhysRefReady As Boolean
Private mComplianceRef As Double
Private mPhysHitCount As Long
Private mPhysCandidate As Boolean
Private mPhysWarnCount As Long
Private mPhysRcMax As Double
Private mPhysRcMaxOk As Boolean
Private mPhysFailInc As Long
Private mPhysFailLambda As Double
Private mPhysFailUmax As Double
Private mPhysFailUmaxH As Double
Private mPhysFailHave As Boolean
Private mPhysMode As String
Private mPhysH As Double
Private mPhysRatioNow As Double
Public Sub FEMWriteBuildStamp(Optional ByVal applyLayout As Boolean = True)
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets("設定")
  ws.range("F1").NumberFormat = "@"
  ws.range("G1").NumberFormat = "@"
  ws.range("F1").value2 = "Ver"
  ws.range("G1").value2 = FEM_BUILD_STAMP
  ws.Cells(1, 1).value2 = "2DFEM 設定"
  ws.Cells(2, 1).value2 = "値はC列。選択肢はF列と「仕様」シート。手順はステージ。LOADは載荷。再開は EXPORT_LOAD。ログはブック横 *_out。"
  ws.Cells(2, 2).value2 = ""
  FEMUiCalculateSettings
  FEMRebuildSettingsLayout ws
  FEMEnsureStageSheets
  If applyLayout Then FEMEnsureWorkbookLayout
  FEMUiAfterLayout
End Sub

Private Sub FEMRebuildSettingsLayout(ByVal ws As Worksheet)
  Dim cacheKey() As String, cacheVal() As Variant
  Dim cacheN As Long, lastRow As Long, rowNo As Long, keyName As String
  Dim writeRow As Long, i As Long, extraFlag As Boolean, foundFlag As Boolean
  Dim clearLast As Long, cacheCap As Long
  lastRow = ws.Cells(ws.rows.count, 5).End(xlUp).row
  If lastRow < 5 Then lastRow = 5
  cacheCap = lastRow
  If cacheCap < 32 Then cacheCap = 32
  ReDim cacheKey(1 To cacheCap)
  ReDim cacheVal(1 To cacheCap)
  On Error Resume Next
  ws.range(ws.Cells(4, 1), ws.Cells(lastRow, 5)).UnMerge
  If lastRow < 200 Then ws.range("A4:E200").UnMerge
  On Error GoTo 0
  For rowNo = 1 To lastRow
    keyName = UCase$(Trim$(CStr(ws.Cells(rowNo, 5).value2)))
    If Len(keyName) > 0 And keyName <> "SRM_ENABLE" And keyName <> "設定キー" Then
      cacheN = cacheN + 1
      If cacheN > cacheCap Then
        cacheCap = cacheCap * 2
        ReDim Preserve cacheKey(1 To cacheCap)
        ReDim Preserve cacheVal(1 To cacheCap)
      End If
      cacheKey(cacheN) = keyName
      cacheVal(cacheN) = ws.Cells(rowNo, 3).value2
    End If
  Next rowNo
  clearLast = lastRow
  If clearLast < 200 Then clearLast = 200
  ws.range(ws.Cells(4, 1), ws.Cells(clearLast, 5)).ClearContents
  ' Clear row-bound rules before rebuilding the key-bound settings layout.
  ws.range(ws.Cells(4, 3), ws.Cells(clearLast, 3)).Validation.Delete
  ws.Cells(4, 1).value2 = "カテゴリ"
  ws.Cells(4, 2).value2 = "項目"
  ws.Cells(4, 3).value2 = "値"
  ws.Cells(4, 4).value2 = "単位・説明"
  ws.Cells(4, 5).value2 = "設定キー"
  writeRow = 5
  FEMPutSection ws, writeRow, "1. メッシュ（拘束は要素作成②後も適用。点座標は要素作成①）"
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_POINT1_X", "簡易", "点1 X", 0#, "左下X（要素作成①）", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_POINT1_Y", "メッシュ", "点1 Y", 0#, "左下Y", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_POINT2_X", "メッシュ", "点2 X", 0#, "左上X", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_POINT2_Y", "メッシュ", "点2 Y", 4#, "左上Y", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_POINT3_X", "メッシュ", "点3 X", 3#, "右上X", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_POINT3_Y", "メッシュ", "点3 Y", 4#, "右上Y", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_POINT4_X", "メッシュ", "点4 X", 3#, "右下X", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_POINT4_Y", "メッシュ", "点4 Y", 0#, "右下Y", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_NX", "メッシュ", "nx", 25#, "X方向分割", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_NY", "メッシュ", "ny", 25#, "Y方向分割", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_MATERIAL", "メッシュ", "材料番号", 1#, "要素作成①②で付ける", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_BC_BOTTOM", "メッシュ", "底面拘束", "ROLLER", "Ymin面。NONE=自由 / ROLLER=Y固定 / PINNED=XY固定", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_BC_LEFT", "メッシュ", "左面拘束", "ROLLER", "Xmin面。NONE=自由 / ROLLER=X固定 / PINNED=XY固定", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_BC_RIGHT", "メッシュ", "右面拘束", "NONE", "Xmax面。NONE=自由 / ROLLER=X固定 / PINNED=XY固定", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_BC_TOP", "メッシュ", "上面拘束", "NONE", "Ymax面。NONE=自由 / ROLLER=Y固定 / PINNED=XY固定", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_BC_PIN_CORNER", "メッシュ", "剛体モード止め", "NONE", "AUTO=底面の最小X節点をXY固定 / NONE=しない", True
  writeRow = writeRow + 1
  FEMPutSection ws, writeRow, "2. 解析（共通）"
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "WEIGHT_MODE", "解析", "自重", "TOTAL", "TOTALのみ。全応力のγ。浮力・水位は未実装", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "Q8_HOURGLASS_FACTOR", "解析", "アワーグラス係数", 0.05, "0～1", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "RCM_POLICY", "解析", "RCM", "AUTO", "AUTO=帯幅が減るときだけ並び替え / ON=必ずRCM / OFF=入力順", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "FLOW_POLICY", "解析", "非関連の扱い", "INCONSISTENT", "INCONSISTENT=流れ則は非関連・接線は対称近似 / DAVIS=等価関連c*,φ*（JOINTには掛けない）", True
  writeRow = writeRow + 1
  FEMPutSection ws, writeRow, "3. SRMパラメータ（実行の有無はステージシート）"
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "SRM_FMAX", "SRM", "上限Fs", 3#, "探索の上限", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "SRM_TOL", "SRM", "許容差", 0.025, "通常0.025 / 検証0.0125。PASSとFAILの幅がこれ以下で終了", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "SRM_PSI_POLICY", "SRM", "ψの扱い", "CAP", "CAP=ψをφ'以下に制限 / REDUCE=ψもFsで低減 / KEEP=入力のまま", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "SRM_MODE", "SRM", "モード", "REAPPLY", "REAPPLYのみ。各Fs試行で先行ステージを再実行", True
  writeRow = writeRow + 1
  FEMPutSection ws, writeRow, "高速化比較（0=従来経路、各方式を単独評価）"
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_V1_PREDICTOR", "高速化", "変位予測子", 0, "0=基準 / 1=同一ステージの確定変位増分から予測", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_V2A_REUSE", "高速化", "増分間接線再利用", 0, "0=基準 / 1=同一塑性集合・小接線変化で最初の補正に旧LU", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_V2B_COST", "高速化", "費用に応じた接線更新", 0, "0=基準 / 1=実測費用と改善実績で限定再利用。上限3は維持", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_V3_COST_SEARCH", "高速化", "費用に応じたFs探索", 0, "0=二分 / 1=保護付き探索。最終実幅は基準二分以下", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_V4_ANDERSON", "高速化", "Anderson深さ1", 0, "0=基準 / 1=凍結接線・未減衰区間のみ。悪化候補は棄却", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_V5_GMRES_LU", "高速化", "旧LU前処理GMRES", 0, "0=Band LU / 1=最新帯行列＋旧LU。真の残差を確認しLUへ復帰", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_TRACE", "高速化", "高速化詳細ログ", 0, "0=追加詳細OFF / 1=正確な更新理由・全増分・候補/復帰をsolver_events.csvへ", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_PREDICT_BETA", "高速化", "予測子係数beta", 1, "検証用パラメータ 0～1。収束条件は変更しない", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_PREDICT_MAX_RATIO", "高速化", "予測子刻み比上限", 1.5, "検証用パラメータ 0超～2。上限超は予測しない", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_TANGENT_CHANGE", "高速化", "接線変化の上限", 0.05, "検証用パラメータ 0超～0.1。要素ごとの相対Frobenius差", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "SRM_FIXED_FS", "高速化", "固定Fs試験", 0, "0=全探索 / 正=そのFsだけ独立再載荷。F<1も可", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_ADAPTIVE", "高速化", "進捗・費用による自動選別", 0, "0=固定フラグ / 1=ONのV1/V2a/V2b/V4/V5を候補に費用・進捗で休止/再試用", False
  writeRow = writeRow + 1
  FEMPutSection ws, writeRow, "4. 出力・外部ファイル"
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "OUTPUT_STAGE_MODE", "出力", "結果の残し方", "FINAL", "FINAL=最後だけ結果シートへ / ALL=各ステージを追記", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "OUTPUT_INACTIVE_ELEMENTS", "出力", "無効要素", "SKIP", "SKIP=Death中は出さない / INCLUDE=行は残しステージ番号だけ書く", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "DEBUG_MODE", "外部", "デバッグログ", "STAGE", "OFF=書かない / STAGE=ステージ境界 / ITER=反復も。*_out/run.log", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "SOLVER", "外部", "帯域ソルバ", "BAND_LU", "BAND_LU=標準 / BAND_LDLT_EXPERIMENTAL=影のLDLTをサンプル比較（本番はLUのまま）", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "BAND_LU_KERNEL", "外部", "Band LUカーネル", "NEW", "NEW=パックド1D（本番） / OLD=従来2D（回帰）。ログは BandLUKernel=", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "PHYSICAL_FAILURE_MODE", "外部", "物理破壊監視", "SHADOW", "未使用。GRAVITYは監視なし、SRMはSHADOWを解析入口で固定", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "DEBUG_FLUSH", "外部", "ログ書き出し間隔", 5#, "秒。長時間計算の監視用", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "EXPORT_MODE", "外部", "途中データの保存", "STAGE", "OFF=保存しない / STAGE=成功ステージ / ON_FAIL=失敗時も。*_out へ", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "EXPORT_LOAD", "外部", "再開フォルダ", "", "空=最初から / AUTO=最後のSUCCESS / stg01_GRAVITY などフォルダ名", True
  writeRow = writeRow + 1
  FEMPutSection ws, writeRow, "5. 図"
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "VIEW_MIN_X", "表示", "原点X", 1000#, "作図原点X", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "VIEW_MAX_Y", "表示", "原点Y", 1000#, "作図原点Y", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "VIEW_SCALE", "表示", "スケール", 50#, "座標→図形", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "VIEW_DISP_SCALE", "表示", "変位倍率", 100#, "変形図", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "VIEW_COLOR_RESULT", "表示", "色付け", 18#, "P1結果ID", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "VIEW_SCALE_MODE", "表示", "色スケール", "SYMMETRIC", "AUTO=最大最小 / SYMMETRIC=正負対称 / USER=手入力", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "VIEW_DETAIL_MODE", "表示", "表示モード", "SIMPLE", "SIMPLE=簡易 / DETAIL=節点番号など詳細", True
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "VIEW_RESULT_SOURCE", "表示", "色の元", "SMOOTHED", "SMOOTHED=節点平滑 / ELEMENT=要素平均 / RAW_GAUSS=Gauss点の生値", True
  writeRow = writeRow + 1
  FEMPutSection ws, writeRow, "6. メッシュ品質（閾値）"
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_QUALITY_MIN_AREA_RATIO", "品質", "最小面積比", 0.000001, "代表寸法^2に対する比", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_QUALITY_MIN_JACOBIAN_RATIO", "品質", "最小Jacobian比", 0.000001, "代表寸法^2に対する比", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_QUALITY_MIN_ANGLE_DEG", "品質", "最小角度", 10#, "deg", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_QUALITY_MAX_ANGLE_DEG", "品質", "最大角度", 170#, "deg", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_QUALITY_MAX_ASPECT_RATIO", "品質", "最大アスペクト", 5#, "長辺/短辺", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_REPAIR_MAX_ITERATIONS", "修正", "最大反復", 5#, "1～50", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_REPAIR_DAMPING", "修正", "移動係数", 0.5, "0～1", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_REPAIR_IMPROVEMENT_TOLERANCE", "修正", "改善判定", 0.0001, "相対", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_REPAIR_ANGLE_ITERATIONS", "修正", "角度最適化反復", 2000#, "1～5000", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_REPAIR_TERRAIN_COMPARE", "修正", "地形追従比較", 1#, "1=地形追従と角度最適化を比較 / 0=しない", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_GEN_COMPARE", "修正", "生成時比較", 1#, "0=両方合格なら品質の良い方 / 1=デローニ修正優先 / 2=地形格子（生格子。角度最適化しない）", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_TERRAIN_COUNT_MIN", "修正", "格子要素数下限比", 0.55, "デローニ要素数に対する比", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "MESH_TERRAIN_COUNT_MAX", "修正", "格子要素数上限比", 1.8, "デローニ要素数に対する比", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ADAPT_MAX_ELEMENT_RATIO", "適応", "要素数倍率上限", 2#, "再メッシュの増加上限", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ADAPT_MARK_FRACTION", "適応", "細分化割合", 0.15, "0～1", False
  writeRow = writeRow + 1
  FEMPutSection ws, writeRow, "高速化追加（増分幅回復）"
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_STEP_RECOVERY", "高速化", "増分幅の緩やかな回復", 0, "0=従来幅 / 1=1.1倍試用後8成功増分の採算を確認。C47=1時のみ", False
  FEMPutCached ws, writeRow, cacheKey, cacheVal, cacheN, "ACCEL_STEP_FRESH_LU", "高速化", "回復直後のfresh LU試験", 0, "0=既存V2a条件 / 1=回復直後1回だけfresh LU。回復ON時のみ", False
  extraFlag = False
  For i = 1 To cacheN
    foundFlag = False
    If left$(cacheKey(i), 18) = "MESH_QUALITY_LAST_" Then foundFlag = True
    If left$(cacheKey(i), 20) = "MESH_QUALITY_REPAIR_" Then foundFlag = True
    If cacheKey(i) = "SOLVER_POLICY" Or cacheKey(i) = "CONSOL_ENABLE" Or cacheKey(i) = "RCM_SEARCH_LIMIT" Then foundFlag = True
    If Not foundFlag Then
      For rowNo = 5 To writeRow - 1
        If UCase$(Trim$(CStr(ws.Cells(rowNo, 5).value2))) = cacheKey(i) Then
          foundFlag = True
          Exit For
        End If
      Next rowNo
    End If
    If Not foundFlag And Len(cacheKey(i)) > 0 Then
      If Not extraFlag Then
        writeRow = writeRow + 1
        FEMPutSection ws, writeRow, "7. 追加設定（保持）"
        extraFlag = True
      End If
      ws.Cells(writeRow, 1).value2 = "追加"
      ws.Cells(writeRow, 2).value2 = cacheKey(i)
      ws.Cells(writeRow, 3).value2 = cacheVal(i)
      ws.Cells(writeRow, 4).value2 = "再構築時に保持"
      ws.Cells(writeRow, 5).value2 = cacheKey(i)
      writeRow = writeRow + 1
    End If
  Next i
  FEMWriteSettingsChoiceGuide ws
  FEMApplySettingValidation ws
End Sub

Private Sub FEMPutSection(ByVal ws As Worksheet, ByRef writeRow As Long, ByVal titleText As String)
  ws.Cells(writeRow, 1).value2 = titleText
  writeRow = writeRow + 1
End Sub

Private Sub FEMPutCached(ByVal ws As Worksheet, ByRef writeRow As Long, ByRef cacheKey() As String, ByRef cacheVal() As Variant, ByVal cacheN As Long, ByVal keyName As String, ByVal categoryText As String, ByVal itemText As String, ByVal defaultValue As Variant, ByVal noteText As String, ByVal isText As Boolean)
  Dim i As Long, foundValue As Variant, foundFlag As Boolean
  foundValue = defaultValue
  For i = 1 To cacheN
    If cacheKey(i) = UCase$(keyName) Then
      foundValue = cacheVal(i)
      foundFlag = True
      Exit For
    End If
  Next i
  If foundFlag Then
    If isText Then
      If Len(Trim$(CStr(foundValue))) = 0 Then foundValue = defaultValue
    Else
      If Not IsNumeric(foundValue) Then foundValue = defaultValue
    End If
  End If
  ws.Cells(writeRow, 1).value2 = categoryText
  ws.Cells(writeRow, 2).value2 = itemText
  If isText Then ws.Cells(writeRow, 3).NumberFormat = "@"
  ws.Cells(writeRow, 3).value2 = foundValue
  ws.Cells(writeRow, 4).value2 = noteText
  ws.Cells(writeRow, 5).value2 = keyName
  writeRow = writeRow + 1
End Sub

Public Sub FEMEnsureStageSheets()
  Dim wsStage As Worksheet, wsLoad As Worksheet, wsOut As Worksheet, wsSet As Worksheet
  Dim wsMat As Worksheet
  On Error Resume Next
  Set wsStage = ThisWorkbook.Worksheets("ステージ")
  Set wsLoad = ThisWorkbook.Worksheets("載荷")
  Set wsSet = ThisWorkbook.Worksheets("要素セット")
  Set wsOut = ThisWorkbook.Worksheets("ステージ結果")
  Set wsMat = ThisWorkbook.Worksheets("材料データ")
  On Error GoTo 0
  If wsStage Is Nothing Then
    Set wsStage = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("設定"))
    wsStage.name = "ステージ"
  End If
  wsStage.Cells(1, 1).value2 = "番号"
  wsStage.Cells(1, 2).value2 = "種別"
  wsStage.Cells(1, 3).value2 = "材料番号"
  wsStage.Cells(1, 4).value2 = "パラメータ"
  wsStage.Cells(1, 5).value2 = "有効"
  wsStage.Cells(1, 6).value2 = "説明"
  If Len(Trim$(CStr(wsStage.Cells(2, 2).value2))) = 0 Then
    wsStage.Cells(2, 1).value2 = 1
    wsStage.Cells(2, 2).value2 = "GRAVITY"
    wsStage.Cells(2, 5).value2 = 1
    wsStage.Cells(2, 6).value2 = "自重。支持は節点データ（拘束かつ変位0）"
    wsStage.Cells(3, 1).value2 = 2
    wsStage.Cells(3, 2).value2 = "LOAD"
    wsStage.Cells(3, 3).value2 = 1
    wsStage.Cells(3, 5).value2 = 1
    wsStage.Cells(3, 6).value2 = "載荷シートの材料1。値0なら実質なし"
    wsStage.Cells(4, 1).value2 = 3
    wsStage.Cells(4, 2).value2 = "SRM"
    wsStage.Cells(4, 4).value2 = 1
    wsStage.Cells(4, 5).value2 = 0
    wsStage.Cells(4, 6).value2 = "途中ゲート。必要Fs=パラメータ。最後の行ならFs探索"
    wsStage.Cells(5, 1).value2 = 4
    wsStage.Cells(5, 2).value2 = "UNLOAD"
    wsStage.Cells(5, 3).value2 = 1
    wsStage.Cells(5, 5).value2 = 0
    wsStage.Cells(5, 6).value2 = "材料1の載荷を外す"
    wsStage.Cells(6, 1).value2 = 5
    wsStage.Cells(6, 2).value2 = "DEATH"
    wsStage.Cells(6, 3).value2 = 1
    wsStage.Cells(6, 5).value2 = 0
    wsStage.Cells(6, 6).value2 = "材料1を撤去（無効化）"
    wsStage.Cells(7, 1).value2 = 6
    wsStage.Cells(7, 2).value2 = "BIRTH"
    wsStage.Cells(7, 3).value2 = 2
    wsStage.Cells(7, 5).value2 = 0
    wsStage.Cells(7, 6).value2 = "材料2を追加（初期状態OFFから有効化）"
  End If
  FEMWriteStageLegend wsStage
  FEMApplyStageValidation wsStage
  If wsLoad Is Nothing Then
    Set wsLoad = ThisWorkbook.Worksheets.Add(After:=wsStage)
    wsLoad.name = "載荷"
  End If
  FEMEnsureLoadSheetLayout wsLoad
  If Not wsSet Is Nothing Then
    On Error Resume Next
    wsSet.Visible = 0
    On Error GoTo 0
  End If
  FEMEnsureMaterialSheetLayout
  FEMEnsureJointSheet
  FEMEnsureSpecSheet
  FEMDeleteCommonSheet
  If wsOut Is Nothing Then
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("要素データ"))
    wsOut.name = "ステージ結果"
  End If
  wsOut.Cells(1, 1).value2 = "番号"
  wsOut.Cells(1, 2).value2 = "種別"
  wsOut.Cells(1, 3).value2 = "材料番号"
  wsOut.Cells(1, 4).value2 = "成否"
  wsOut.Cells(1, 5).value2 = "Fsまたはλ"
  wsOut.Cells(1, 6).value2 = "最大変位"
  wsOut.Cells(1, 7).value2 = "反力X"
  wsOut.Cells(1, 8).value2 = "反力Y"
  wsOut.Cells(1, 9).value2 = "試行"
  wsOut.Cells(1, 10).value2 = "メッセージ"
  On Error Resume Next
  wsStage.Move After:=ThisWorkbook.Worksheets("設定")
  wsLoad.Move After:=ThisWorkbook.Worksheets("ステージ")
  If Not wsMat Is Nothing Then wsMat.Move After:=ThisWorkbook.Worksheets("設定")
  On Error Resume Next
  ThisWorkbook.Worksheets("接合").Move After:=ThisWorkbook.Worksheets("載荷")
  ThisWorkbook.Worksheets("ステージ結果").Move After:=ThisWorkbook.Worksheets("要素データ")
  On Error GoTo 0
End Sub

Private Sub FEMDeleteCommonSheet()
  Dim ws As Worksheet
  Dim previousAlerts As Boolean
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("共通データ")
  On Error GoTo 0
  If ws Is Nothing Then Exit Sub
  previousAlerts = Application.DisplayAlerts
  Application.DisplayAlerts = False
  On Error Resume Next
  ws.Delete
  On Error GoTo 0
  Application.DisplayAlerts = previousAlerts
End Sub

Public Sub FEMEnsureWorkbookLayout()
  Dim previousAlerts As Boolean
  previousAlerts = Application.DisplayAlerts
  Application.DisplayAlerts = False
  On Error Resume Next
  FEMIoDeleteSheet "Sheet1"
  FEMIoDeleteSheet "P1結果"
  FEMIoDeleteSheet "P2材料点結果"
  FEMIoDeleteSheet "P2材料点試験"
  FEMIoRenameSheet "解析結果", FEM_INTEGRATED_RESULT_SHEET
  FEMIoHideSheet "表示量一覧"
  FEMIoHideSheet "メッシュ引継"
  FEMIoHideSheet "再メッシュ候補"
  FEMIoSetNodeHeaders
  FEMIoOrderSheets
  FEMIoTrimDiagnosticSheet
  On Error GoTo 0
  Application.DisplayAlerts = previousAlerts
End Sub

Private Sub FEMIoTrimDiagnosticSheet()
  Dim ws As Worksheet
  Dim lastRow As Long
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets(FEM_INTEGRATED_RESULT_SHEET)
  On Error GoTo 0
  If ws Is Nothing Then Exit Sub
  lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
  If lastRow < ws.Cells(ws.rows.count, 4).End(xlUp).row Then lastRow = ws.Cells(ws.rows.count, 4).End(xlUp).row
  If lastRow > FEM_DIAGNOSTIC_LAST_ROW Then
    ws.range(ws.Cells(FEM_INTEGRATED_P1_START_ROW, 1), ws.Cells(lastRow, 36)).ClearContents
  End If
  lastRow = ws.Cells(ws.rows.count, 4).End(xlUp).row
  If lastRow > 80 Then
    ws.range(ws.Cells(81, 4), ws.Cells(lastRow, 14)).ClearContents
  End If
End Sub

Private Sub FEMIoDeleteSheet(ByVal sheetName As String)
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets(sheetName)
  If Not ws Is Nothing Then ws.Delete
  Err.Clear
End Sub

Private Sub FEMIoRenameSheet(ByVal oldName As String, ByVal newName As String)
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets(newName)
  If Not ws Is Nothing Then Exit Sub
  Set ws = ThisWorkbook.Worksheets(oldName)
  If Not ws Is Nothing Then ws.name = newName
  Err.Clear
End Sub

Private Sub FEMIoHideSheet(ByVal sheetName As String)
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets(sheetName)
  If Not ws Is Nothing Then ws.Visible = xlSheetVeryHidden
  Err.Clear
End Sub

Private Sub FEMIoSetNodeHeaders()
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("節点データ")
  On Error GoTo 0
  If ws Is Nothing Then Exit Sub
  ws.Cells(1, 1).value2 = "節点番号"
  ws.Cells(1, 2).value2 = "X座標"
  ws.Cells(1, 3).value2 = "Y座標"
  ws.Cells(1, 4).value2 = "X拘束条件"
  ws.Cells(1, 5).value2 = "Y拘束条件"
  ws.Cells(1, 6).value2 = "X変位(互換)"
  ws.Cells(1, 7).value2 = "Y変位(互換)"
  ws.Cells(1, 8).value2 = "X力(互換)"
  ws.Cells(1, 9).value2 = "Y力(互換)"
  ws.Cells(1, 11).value2 = "凡例（入力しない）"
  ws.Cells(2, 11).value2 = "X/Y拘束条件"
  ws.Cells(2, 12).value2 = "0=自由 / 1=固定。載荷のFIXでも拘束できる"
  ws.Cells(3, 11).value2 = "変位・力列"
  ws.Cells(3, 12).value2 = "互換用。通常の載荷は載荷シートへ"
End Sub

Private Sub FEMIoOrderSheets()
  Dim names As Variant
  Dim i As Long, ws As Worksheet
  names = Array("設定", "仕様", "材料データ", "図形定義", "ステージ", "載荷", "接合", "節点データ", "要素データ", "図", "ステージ結果", "結果節点", "結果要素", FEM_INTEGRATED_RESULT_SHEET)
  On Error Resume Next
  For i = LBound(names) To UBound(names)
    Set ws = ThisWorkbook.Worksheets(CStr(names(i)))
    If Not ws Is Nothing Then
      If i = LBound(names) Then
        ws.Move Before:=ThisWorkbook.Worksheets(1)
      Else
        ws.Move After:=ThisWorkbook.Worksheets(CStr(names(i - 1)))
      End If
    End If
    Set ws = Nothing
  Next i
  Err.Clear
End Sub

Public Sub FEMIoBegin()
  Dim folderPath As String
  P6ResetBandBench
  FEMIoClose
  femIoMode = UCase$(Trim$(FEMReadTextSetting("DEBUG_MODE", "STAGE")))
  femIoExportMode = UCase$(Trim$(FEMReadTextSetting("EXPORT_MODE", "STAGE")))
  femIoFlushSec = FEMReadSetting("DEBUG_FLUSH", 5#)
  If femIoFlushSec < 1# Then femIoFlushSec = 1#
  femIoResumeStage = 0
  femIoResumeFolder = vbNullString
  femIoResumeFs = 1#
  femIoResumeFss = 1#
  femIoResumeGravity = False
  femIoResumeHasLock = False
  femIoResumePrefix = 0
  femIoStartedAt = Timer
  If femIoMode = "OFF" And femIoExportMode = "OFF" Then Exit Sub
  folderPath = FEMIoOutDir()
  If Len(folderPath) = 0 Then Exit Sub
  femIoPath = folderPath & Application.PathSeparator & "run.log"
  On Error Resume Next
  femIoFileNum = FreeFile
  Open femIoPath For Output As #femIoFileNum
  If Err.Number <> 0 Then
    femIoFileNum = 0
    Err.Clear
    Exit Sub
  End If
  Print #femIoFileNum, "time,elapsed_s,place,stage,kind,Fs,inc,iter,relres,status,message"
  Print #femIoFileNum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & ",0,BEGIN,0,,1,0,0,0,RUN,Ver=" & FEM_BUILD_STAMP
  femIoLastFlush = Timer
  On Error GoTo 0
  FEMAppendRunLog "SOLVER", "KERNEL", "BandLUKernel=" & P6BandKernelName()
  If femIoMode <> "OFF" Then FEMIoResetMeasureFiles folderPath
End Sub

Private Sub FEMIoDeleteIfExists(ByVal pathText As String)
  On Error Resume Next
  If Len(pathText) = 0 Then Exit Sub
  If Dir$(pathText) <> vbNullString Then Kill pathText
  Err.Clear
End Sub

Private Sub FEMIoResetMeasureFiles(ByVal folderPath As String)
  Dim sep As String
  If Len(folderPath) = 0 Then Exit Sub
  sep = Application.PathSeparator
  FEMIoWriteNewCsv folderPath, "perf_summary.csv", P6CsvHeaderSummary()
  FEMIoWriteNewCsv folderPath, "perf_trace.csv", P6CsvHeaderTrace()
  FEMIoWriteNewCsv folderPath, "solver_events.csv", P6CsvHeaderEvent()
  FEMIoWriteNewCsv folderPath, "adaptive_v2a.csv", PolicyLogSampleHeader()
  FEMIoWriteNewCsv folderPath, "adaptive_v2a_gates.csv", PolicyLogGateHeader()
  FEMIoWriteNewCsv folderPath, "increment_summary.csv", IncrementLogHeader()
  If P6ReadSetting("BAND_LU_BENCH", 0#) >= 1# Then
    FEMIoWriteNewCsv folderPath, "band_lu_bench.csv", P6CsvHeaderBench()
    P6CsvBenchHeader = True
  Else
    FEMIoDeleteIfExists folderPath & sep & "band_lu_bench.csv"
    P6CsvBenchHeader = False
  End If
  FEMIoDeleteIfExists folderPath & sep & "physical_failure_trace.csv"
  FEMIoDeleteIfExists folderPath & sep & "perf_summary.txt"
  P6CsvSummaryHeader = True
  P6CsvTraceHeader = True
  P6CsvEventHeader = True
  P6CsvPhysHeader = True
  P6PerfSummaryWritten = False
End Sub

Private Function P6CsvHeaderSummary() As String
  Dim s As String
  s = "ver,fs,numerical_status,mechanical_status,elapsed_s,inc,newton,factor,reuse,umax,umax_over_H,limit_candidate,first_candidate_inc,first_candidate_umax,max_compliance_ratio,relres,energy"
  s = s & ",newton_policy,ls50_enabled,ls50_tried,ls50_ok,ls50_saved"
  s = s & ",ls50_gate_policy_reject,ls50_gate_disabled,ls50_gate_alpha_reject,ls50_gate_qls_reject,ls50_gate_cutback_reject,ls50_gate_active_reject,ls50_gate_residual_reject,ls50_gate_factor_invalid"
  s = s & ",band_kernel,factor_avg_ms"
  P6CsvHeaderSummary = s
End Function

Private Function P6CsvHeaderTrace() As String
  P6CsvHeaderTrace = "ver,fs,inc,lambda,dlambda20,dlambda80,umax,du20,du80,compliance20,compliance_ref,compliance_ratio,plastic_gp_ratio,new_plastic_gp20,factor_ratio20,reuse_ratio20,damped_ls_ratio20"
End Function

Private Function P6CsvHeaderEvent() As String
  P6CsvHeaderEvent = "ver,stage,kind,fs,inc,event,detail"
End Function

Private Function P6CsvHeaderBench() As String
  P6CsvHeaderBench = "ver,n,bw,reps,old_factor_ms,new_factor_ms,old_solve_ms,new_solve_ms,pack_ms,speedup_factor,speedup_net,x_error,residual_old,residual_new,gate,reason"
End Function

Public Sub FEMIoAppendBenchRow(ByVal lineText As String)
  FEMIoAppendCsv "band_lu_bench.csv", P6CsvBenchHeader, P6CsvHeaderBench(), lineText
End Sub

Private Function P6CsvHeaderPhys() As String
  P6CsvHeaderPhys = "ver,reason,stage,kind,fs,inc,newton_total,lambda,remaining_lambda,step_size,step_ratio,umax,umax_over_H,du_max,du_max_over_H,du20,du80,du_growth_ratio20,dlambda20,dlambda80,compliance20,compliance80,compliance_ref,compliance_ratio,du20_ref,du_ratio,newton_per_inc20,newton_per_inc_ref,plastic_gp,plastic_gp_ratio,new_plastic_gp20,new_plastic_gp80,max_eq_plastic_strain,mean_eq_plastic_strain,plastic_element_count,largest_plastic_component,plastic_band_connected,relres,factor_total,reuse_total,factor_ratio20,reuse_ratio20,damped_ls_ratio20,factor_ratio80,reuse_ratio80,damped_ls_ratio80,cutback_total,damped_ls_total,q_ls,q_next,ls_alpha,tangent_age,would_physical_fail,du20_over_H,du20_h_ge_min"
End Function

Private Sub FEMIoWriteNewCsv(ByVal folderPath As String, ByVal fileName As String, ByVal headerText As String)
  Dim pathText As String, fileNum As Integer
  On Error Resume Next
  pathText = folderPath & Application.PathSeparator & fileName
  fileNum = FreeFile
  Open pathText For Output As #fileNum
  Print #fileNum, headerText
  Close #fileNum
  Err.Clear
End Sub

Public Sub FEMIoLog(ByVal placeText As String, ByVal statusText As String, ByVal messageText As String)
  Dim stageNo As Long, kindText As String
  If femIoFileNum = 0 Then Exit Sub
  If femIoMode = "OFF" Then Exit Sub
  If femIoMode = "STAGE" Then
    If InStr(1, placeText, "反復", vbTextCompare) > 0 Then Exit Sub
    If InStr(1, UCase$(placeText), "ITER") > 0 Then Exit Sub
    If InStr(1, UCase$(placeText), "CORRECTION") > 0 Then Exit Sub
  End If
  stageNo = P3RunLogStageNo
  kindText = P3RunLogKind
  If stageNo <= 0 And P3ActivePrefix >= 1 And P3ActivePrefix <= P3StageN Then
    stageNo = P3StageId(P3ActivePrefix)
    kindText = P3StageKind(P3ActivePrefix)
  End If
  FEMIoWriteLine placeText, statusText, messageText, stageNo, kindText
End Sub

Private Sub FEMIoWriteLine(ByVal placeText As String, ByVal statusText As String, ByVal messageText As String, ByVal stageNo As Long, ByVal kindText As String)
  Dim elapsed As Double, lineText As String
  If femIoFileNum = 0 Then Exit Sub
  elapsed = Timer - femIoStartedAt
  If elapsed < 0# Then elapsed = elapsed + 86400#
  lineText = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "," & Format$(elapsed, "0.000") & "," & FEMIoCsv(placeText) & "," & CStr(stageNo) & "," & FEMIoCsv(kindText) & "," & Format$(P3CurrentStrengthFactor, "0.000") & "," & CStr(CurrentIncrement) & "," & CStr(CurrentIteration) & "," & Format$(RelativeResidualFree, "0.000E+00") & "," & FEMIoCsv(statusText) & "," & FEMIoCsv(messageText)
  On Error Resume Next
  Print #femIoFileNum, lineText
  If Timer - femIoLastFlush >= femIoFlushSec Or Timer - femIoLastFlush < 0# Then
    Close #femIoFileNum
    femIoFileNum = FreeFile
    Open femIoPath For Append As #femIoFileNum
    femIoLastFlush = Timer
  End If
  Err.Clear
End Sub

Public Sub FEMIoClose()
  On Error Resume Next
  P6FlushBandBench
  P6PerfWriteSummary
  If femIoFileNum <> 0 Then Close #femIoFileNum
  femIoFileNum = 0
  Err.Clear
End Sub

Private Function FEMIoCsv(ByVal textValue As String) As String
  FEMIoCsv = Replace$(Replace$(Replace$(CStr(textValue), ",", ";"), vbCr, " "), vbLf, " ")
End Function

Public Sub P6PerfMark()
  PolicyLogMark
  P6PerfMarkNewton = P6PerfNewtonCount
  P6PerfMarkIncrement = P6PerfIncrementCount
  P6PerfMarkCutback = P6PerfCutbackCount
  P6PerfMarkFactor = P6FactorizationCount
  P6PerfMarkReuse = P6FactorizationReuseCount
  P6PerfMarkBand = P6BandSolveCallCount
  P6PerfMarkModNewton = P6PerfModNewtonSkips
  P6PerfMarkRebuildFirst = P6PerfRebuildFirst
  P6PerfMarkRebuildForce = P6PerfRebuildForce
  P6PerfMarkRebuildForceStart = P6PerfRebuildForceStart
  P6PerfMarkRebuildForceCutback = P6PerfRebuildForceCutback
  P6PerfMarkRebuildForceLsCut = P6PerfRebuildForceLsCut
  P6PerfMarkRebuildForceLsOk = P6PerfRebuildForceLsOk
  P6PerfMarkRebuildForceLsRetry = P6PerfRebuildForceLsRetry
  P6PerfMarkRebuildForceOther = P6PerfRebuildForceOther
  P6PerfMarkRebuildActiveSet = P6PerfRebuildActiveSet
  P6PerfMarkRebuildResidual = P6PerfRebuildResidual
  P6PerfMarkRebuildStale = P6PerfRebuildStale
  P6PerfMarkRebuildFactorInvalid = P6PerfRebuildFactorInvalid
  P6PerfMarkTangent = P6ProfTangentCount
  P6PerfMarkFactorMs = P6ProfFactorMs
  P6PerfMarkSolveMs = P6ProfSolveMs
  P6PerfMarkTangentMs = P6ProfTangentMs
  P6PivotMinTrial = 1E+308
  P6PivotNegCountTrial = 0
  P6PivotNearZeroCountTrial = 0
  P6PerfMarkLsTries = P6LsTries
  P6PerfMarkLsAccept = P6LsAccept
  P6PerfMarkLsA1 = P6LsA1
  P6PerfMarkLsA50 = P6LsA50
  P6PerfMarkLsA25 = P6LsA25
  P6PerfMarkLsA12 = P6LsA12
  P6PerfMarkLsOkA1 = P6LsOkA1
  P6PerfMarkLsOkA50 = P6LsOkA50
  P6PerfMarkLsOkA25 = P6LsOkA25
  P6PerfMarkLsOkA12 = P6LsOkA12
  P6PerfMarkLsAlphaSum = P6LsAlphaSum
  P6LsAlphaMinTrial = 1E+308
  P6LsLastAlpha = 0#
  P6PerfMarkLsQCount = P6LsQCount
  P6PerfMarkLsQSum = P6LsQSum
  P6PerfMarkLsQLt50 = P6LsQLt50
  P6PerfMarkLsQLt70 = P6LsQLt70
  P6PerfMarkLsQLt100 = P6LsQLt100
  P6PerfMarkLsQGe100 = P6LsQGe100
  P6PerfMarkLsQnCount = P6LsQnCount
  P6PerfMarkLsQnSum = P6LsQnSum
  P6PerfMarkLsQnLt70 = P6LsQnLt70
  P6PerfMarkTan0 = P6TanSolve0
  P6PerfMarkTan1 = P6TanSolve1
  P6PerfMarkTan2 = P6TanSolve2
  P6PerfMarkTan3 = P6TanSolve3
  P6PerfMarkAge1Ok = P6Age1Success
  P6PerfMarkAge1Resid = P6Age1Resid
  P6PerfMarkAge1Stag = P6Age1Stag
  P6PerfMarkAge1Active = P6Age1Active
  P6PerfMarkAge1Force = P6Age1Force
  P6PerfMarkAge1Other = P6Age1Other
  P6PerfMarkAssembleMs = P6ProfAssembleMs
  P6PerfMarkEvalMs = P6ProfEvalMs
  P6PerfMarkTimer = Timer
  P6PivotMaxTrial = 0#
  P6RollCount = 0
  P6PhysResetTrial
  P6PerfMarkLdltN = P6LdltN
  P6PerfMarkLdltOk = P6LdltOk
  P6PerfMarkLdltFail = P6LdltFail
  P6PerfMarkLdltMs = P6LdltMs
  P6LdltExMaxTrial = 0#
  P6LdltErMaxTrial = 0#
  P6LdltDRatioMinTrial = 1E+308
  P6PerfMarkLs50Tried = P6Ls50Tried
  P6PerfMarkLs50Ok = P6Ls50Ok
  P6PerfMarkLs50Resid = P6Ls50Resid
  P6PerfMarkLs50Stag = P6Ls50Stag
  P6PerfMarkLs50Worse = P6Ls50Worse
  P6PerfMarkLs50Saved = P6Ls50Saved
  P6PerfMarkLs50GatePolicy = P6Ls50GatePolicy
  P6PerfMarkLs50GateDisabled = P6Ls50GateDisabled
  P6PerfMarkLs50GateAlpha = P6Ls50GateAlpha
  P6PerfMarkLs50GateQls = P6Ls50GateQls
  P6PerfMarkLs50GateCutback = P6Ls50GateCutback
  P6PerfMarkLs50GateActive = P6Ls50GateActive
  P6PerfMarkLs50GateResidual = P6Ls50GateResidual
  P6PerfMarkLs50GateFactor = P6Ls50GateFactor
  P6LsWatch = False
  P3TangentAge = 0
  P3TangentJustRebuilt = False
End Sub

Public Sub P6LsNoteTry()
  P6LsTries = P6LsTries + 1
End Sub

Public Sub P6LsNoteAccept(ByVal alpha As Double, ByVal forceOk As Boolean)
  P6LsAccept = P6LsAccept + 1
  P6LsAlphaSum = P6LsAlphaSum + alpha
  P6LsLastAlpha = alpha
  If alpha < P6LsAlphaMin Then P6LsAlphaMin = alpha
  If alpha < P6LsAlphaMinTrial Then P6LsAlphaMinTrial = alpha
  If Abs(alpha - 1#) <= 0.01 Then
    P6LsA1 = P6LsA1 + 1
    If forceOk Then P6LsOkA1 = P6LsOkA1 + 1
  ElseIf Abs(alpha - 0.5) <= 0.01 Then
    P6LsA50 = P6LsA50 + 1
    If forceOk Then P6LsOkA50 = P6LsOkA50 + 1
  ElseIf Abs(alpha - 0.25) <= 0.01 Then
    P6LsA25 = P6LsA25 + 1
    If forceOk Then P6LsOkA25 = P6LsOkA25 + 1
  Else
    P6LsA12 = P6LsA12 + 1
    If forceOk Then P6LsOkA12 = P6LsOkA12 + 1
  End If
End Sub

Public Sub P6LsNoteQ(ByVal beforeRes As Double, ByVal afterRes As Double)
  Dim q As Double
  If P6LsWatch And P6LsWatchAfter > 0# And afterRes >= 0# Then
    q = afterRes / P6LsWatchAfter
    P6LsQnCount = P6LsQnCount + 1
    P6LsQnSum = P6LsQnSum + q
    If q < 0.7 Then P6LsQnLt70 = P6LsQnLt70 + 1
    P6LastQnext = q
    P6LastQnextOk = True
  End If
  If beforeRes > 0# Then
    q = afterRes / beforeRes
    P6LastQls = q
    P6LastQlsOk = True
    P6LsQCount = P6LsQCount + 1
    P6LsQSum = P6LsQSum + q
    If q < 0.5 Then
      P6LsQLt50 = P6LsQLt50 + 1
    ElseIf q < 0.7 Then
      P6LsQLt70 = P6LsQLt70 + 1
    ElseIf q < 1# Then
      P6LsQLt100 = P6LsQLt100 + 1
    Else
      P6LsQGe100 = P6LsQGe100 + 1
    End If
  End If
  P6LsWatch = True
  P6LsWatchAfter = afterRes
End Sub

Public Sub P6NoteAgeOutcome(ByVal rebuilt As Boolean)
  If P3TangentAge <> 1 Then Exit Sub
  If Not rebuilt Then
    P6Age1Success = P6Age1Success + 1
    Exit Sub
  End If
  Select Case P3LastRebuildWhy
    Case "RESID"
      P6Age1Resid = P6Age1Resid + 1
    Case "STAG"
      P6Age1Stag = P6Age1Stag + 1
    Case "ACTIVE"
      P6Age1Active = P6Age1Active + 1
    Case "FORCE"
      P6Age1Force = P6Age1Force + 1
    Case Else
      P6Age1Other = P6Age1Other + 1
  End Select
End Sub

Public Sub P6RollPush(ByVal lambdaNow As Double, ByVal duMax As Double)
  Dim idx As Long
  idx = P6RollCount Mod P6_ROLL_CAP
  P6RollLambda(idx) = lambdaNow
  P6RollFactor(idx) = P6FactorizationCount
  P6RollNewton(idx) = P6PerfNewtonCount
  P6RollReuse(idx) = P6FactorizationReuseCount
  P6RollAccept(idx) = P6LsAccept
  P6RollDamped(idx) = P6LsAccept - P6LsA1
  P6RollUmax(idx) = P6PerfMaxAbsDisp()
  P6RollDu(idx) = duMax
  P6RollPlastic(idx) = P3ActivePlasticPointCount
  P6RollCut(idx) = P6PerfCutbackCount
  P6RollCount = P6RollCount + 1
End Sub

Public Function P6RollWindowText(ByVal lambdaNow As Double) As String
  Dim s As String
  Dim back As Long
  Dim idxNow As Long, idxOld As Long
  Dim dN As Long, dF As Long, dR As Long, dA As Long, dD As Long
  Dim lambdaNowSnap As Double
  s = ""
  If P6RollCount < 1 Then
    P6RollWindowText = ";dlambda20=NA;dlambda80=NA;factor_ratio80=NA;reuse_ratio80=NA;damped_ls_ratio80=NA"
    Exit Function
  End If
  idxNow = (P6RollCount - 1) Mod P6_ROLL_CAP
  lambdaNowSnap = P6RollLambda(idxNow)
  For back = 20 To 80 Step 60
    If P6RollCount > back And back < P6_ROLL_CAP Then
      idxOld = (P6RollCount - 1 - back) Mod P6_ROLL_CAP
      s = s & ";dlambda" & CStr(back) & "=" & Format$(lambdaNowSnap - P6RollLambda(idxOld), "0.000E+00")
      If back = 80 Then
        dN = P6RollNewton(idxNow) - P6RollNewton(idxOld)
        dF = P6RollFactor(idxNow) - P6RollFactor(idxOld)
        dR = P6RollReuse(idxNow) - P6RollReuse(idxOld)
        dA = P6RollAccept(idxNow) - P6RollAccept(idxOld)
        dD = P6RollDamped(idxNow) - P6RollDamped(idxOld)
        If dN > 0 Then
          s = s & ";factor_ratio80=" & Format$(CDbl(dF) / CDbl(dN), "0.000")
          s = s & ";reuse_ratio80=" & Format$(CDbl(dR) / CDbl(dN), "0.000")
        Else
          s = s & ";factor_ratio80=NA;reuse_ratio80=NA"
        End If
        If dA > 0 Then
          s = s & ";damped_ls_ratio80=" & Format$(CDbl(dD) / CDbl(dA), "0.000")
        Else
          s = s & ";damped_ls_ratio80=NA"
        End If
      End If
    Else
      s = s & ";dlambda" & CStr(back) & "=NA"
      If back = 80 Then s = s & ";factor_ratio80=NA;reuse_ratio80=NA;damped_ls_ratio80=NA"
    End If
  Next back
  P6RollWindowText = s
End Function

Private Sub FEMIoAppendCsv(ByVal fileName As String, ByRef headerFlag As Boolean, ByVal headerText As String, ByVal lineText As String)
  Dim folderPath As String, pathText As String, fileNum As Integer, needHeader As Boolean
  If femIoMode = "OFF" Then Exit Sub
  On Error Resume Next
  folderPath = FEMIoOutDir()
  If Len(folderPath) = 0 Then Exit Sub
  pathText = folderPath & Application.PathSeparator & fileName
  needHeader = (Dir$(pathText) = vbNullString)
  fileNum = FreeFile
  Open pathText For Append As #fileNum
  If needHeader Then Print #fileNum, headerText
  Print #fileNum, lineText
  Close #fileNum
  headerFlag = True
  Err.Clear
End Sub

Public Sub P6SolverEvent(ByVal eventName As String, ByVal detailText As String)
  Dim headerText As String, lineText As String
  Dim stageNo As Long, kindText As String
  If P3SrmTrialRunning And P3SrmLogStageNo > 0 Then
    stageNo = P3SrmLogStageNo
    kindText = P3SrmLogKind
    If Len(kindText) = 0 Then kindText = "SRM"
  Else
    stageNo = P3RunLogStageNo
    kindText = P3RunLogKind
  End If
  headerText = P6CsvHeaderEvent()
  lineText = FEM_BUILD_STAMP & "," & CStr(stageNo) & "," & FEMIoCsv(kindText) & "," & Format$(P3CurrentStrengthFactor, "0.000") & "," & CStr(CurrentIncrement) & "," & FEMIoCsv(eventName) & "," & FEMIoCsv(detailText)
  FEMIoAppendCsv "solver_events.csv", P6CsvEventHeader, headerText, lineText
End Sub

Private Sub P6WriteTrialCsv(ByVal statusText As String)
  Dim headerText As String, lineText As String
  Dim elapsed As Double
  Dim numText As String, mechText As String
  Dim h As Double
  elapsed = Timer - P6PerfMarkTimer
  If elapsed < 0# Then elapsed = elapsed + 86400#
  P6PushLimitState
  numText = P6NumericalStatus(statusText)
  If P6Limit.candidate Then
    mechText = "LIMIT_STATE"
  ElseIf numText = "CONVERGED" Then
    mechText = "STABLE"
  Else
    mechText = "UNKNOWN"
  End If
  P3LastTrial.Fs = P3CurrentStrengthFactor
  If numText = "CONVERGED" Then
    P3LastTrial.NumericalStatus = NS_CONVERGED
  ElseIf numText = "FAILED" Then
    P3LastTrial.NumericalStatus = NS_FAILED
  Else
    P3LastTrial.NumericalStatus = NS_UNKNOWN
  End If
  If mechText = "LIMIT_STATE" Then
    P3LastTrial.MechanicalStatus = MS_LIMIT_STATE
  ElseIf mechText = "STABLE" Then
    P3LastTrial.MechanicalStatus = MS_STABLE
  Else
    P3LastTrial.MechanicalStatus = MS_UNKNOWN
  End If
  P3LastTrial.IncrementCount = P6PerfIncrementCount - P6PerfMarkIncrement
  P3LastTrial.NewtonCount = P6PerfNewtonCount - P6PerfMarkNewton
  P3LastTrial.MaxDisp = P6PerfMaxAbsDisp()
  h = P6ModelHeight()
  If h > 0# Then
    P3LastTrial.MaxDispOverH = P3LastTrial.MaxDisp / h
  Else
    P3LastTrial.MaxDispOverH = 0#
  End If
  P3LastTrial.ComplianceMaxRatio = P6Limit.RatioMax
  P3LastTrial.FactorCount = P6FactorizationCount - P6PerfMarkFactor
  P3LastTrial.FactorReuseCount = P6FactorizationReuseCount - P6PerfMarkReuse
  P3LastTrial.ElapsedSec = elapsed
  headerText = P6CsvHeaderSummary()
  lineText = FEM_BUILD_STAMP
  lineText = lineText & "," & Format$(P3CurrentStrengthFactor, "0.000")
  lineText = lineText & "," & numText
  lineText = lineText & "," & mechText
  lineText = lineText & "," & Format$(elapsed, "0.000")
  lineText = lineText & "," & CStr(P3LastTrial.IncrementCount)
  lineText = lineText & "," & CStr(P3LastTrial.NewtonCount)
  lineText = lineText & "," & CStr(P3LastTrial.FactorCount)
  lineText = lineText & "," & CStr(P3LastTrial.FactorReuseCount)
  lineText = lineText & "," & Format$(P3LastTrial.MaxDisp, "0.000E+00")
  If h > 0# Then
    lineText = lineText & "," & Format$(P3LastTrial.MaxDispOverH, "0.000E+00")
  Else
    lineText = lineText & ",NA"
  End If
  If P6Limit.candidate Then
    lineText = lineText & ",1"
    lineText = lineText & "," & CStr(P6Limit.FirstCandidateInc)
    lineText = lineText & "," & Format$(P6Limit.FirstCandidateUmax, "0.000E+00")
  Else
    lineText = lineText & ",0,NA,NA"
  End If
  If mPhysRcMaxOk Then
    lineText = lineText & "," & Format$(P6Limit.RatioMax, "0.000E+00")
  Else
    lineText = lineText & ",NA"
  End If
  lineText = lineText & "," & Format$(RelativeResidualFree, "0.000E+00")
  lineText = lineText & "," & Format$(EnergyError, "0.000E+00")
  If P3NewtonPolicy = NP_ROBUST_SRM Then
    lineText = lineText & ",ROBUST_SRM"
  Else
    lineText = lineText & ",NORMAL"
  End If
  If P3Ls50Enabled Then
    lineText = lineText & ",1"
  Else
    lineText = lineText & ",0"
  End If
  lineText = lineText & "," & CStr(P6Ls50Tried - P6PerfMarkLs50Tried)
  lineText = lineText & "," & CStr(P6Ls50Ok - P6PerfMarkLs50Ok)
  lineText = lineText & "," & CStr(P6Ls50Saved - P6PerfMarkLs50Saved)
  lineText = lineText & "," & CStr(P6Ls50GatePolicy - P6PerfMarkLs50GatePolicy)
  lineText = lineText & "," & CStr(P6Ls50GateDisabled - P6PerfMarkLs50GateDisabled)
  lineText = lineText & "," & CStr(P6Ls50GateAlpha - P6PerfMarkLs50GateAlpha)
  lineText = lineText & "," & CStr(P6Ls50GateQls - P6PerfMarkLs50GateQls)
  lineText = lineText & "," & CStr(P6Ls50GateCutback - P6PerfMarkLs50GateCutback)
  lineText = lineText & "," & CStr(P6Ls50GateActive - P6PerfMarkLs50GateActive)
  lineText = lineText & "," & CStr(P6Ls50GateResidual - P6PerfMarkLs50GateResidual)
  lineText = lineText & "," & CStr(P6Ls50GateFactor - P6PerfMarkLs50GateFactor)
  lineText = lineText & "," & P6BandKernelName()
  lineText = lineText & "," & P6TrialFactorAvgText()
  FEMIoAppendCsv "perf_summary.csv", P6CsvSummaryHeader, headerText, lineText
  FEMIoFlushPolicy statusText
End Sub


Private Function P6PerfMaxAbsDisp() As Double
  Dim i As Long, v As Double
  v = 0#
  If lastDof < 0 Then Exit Function
  For i = 0 To lastDof
    If NodeCond(i) = 0 Then
      If Abs(TDisp(i)) > v Then v = Abs(TDisp(i))
    End If
  Next i
  P6PerfMaxAbsDisp = v
End Function

Private Function P6PerfPlasticBudget() As Long
  Dim k As Long, n As Long
  n = 0
  If NumberOfElement < 1 Then
    P6PerfPlasticBudget = 0
    Exit Function
  End If
  For k = 0 To NumberOfElement - 1
    If P3IsElementActive(k) Then
      If Not P3ElementIsJoint(k) Then n = n + 4
    End If
  Next k
  P6PerfPlasticBudget = n
End Function

Private Function P6PerfDeltaLine() As String
  Dim budget As Long
  Dim s As String
  budget = P6PerfPlasticBudget()
  s = "newton=" & CStr(P3GlobalIterationCount)
  s = s & ";inc=" & CStr(P3SuccessfulIncrementCount)
  s = s & ";inc_now=" & CStr(CurrentIncrement)
  s = s & ";iter_now=" & CStr(CurrentIteration)
  s = s & ";factor=" & CStr(P6FactorizationCount - P6PerfMarkFactor)
  s = s & ";reuse=" & CStr(P6FactorizationReuseCount - P6PerfMarkReuse)
  s = s & ";band=" & CStr(P6BandSolveCallCount - P6PerfMarkBand)
  s = s & ";modnewton=" & CStr(P6PerfModNewtonSkips - P6PerfMarkModNewton)
  s = s & ";k_build=" & CStr(P6ProfTangentCount - P6PerfMarkTangent)
  s = s & ";rebuild_first=" & CStr(P6PerfRebuildFirst - P6PerfMarkRebuildFirst)
  s = s & ";rebuild_force=" & CStr(P6PerfRebuildForce - P6PerfMarkRebuildForce)
  s = s & ";force_start=" & CStr(P6PerfRebuildForceStart - P6PerfMarkRebuildForceStart)
  s = s & ";force_cut=" & CStr(P6PerfRebuildForceCutback - P6PerfMarkRebuildForceCutback)
  s = s & ";force_lscut=" & CStr(P6PerfRebuildForceLsCut - P6PerfMarkRebuildForceLsCut)
  s = s & ";force_lsok=" & CStr(P6PerfRebuildForceLsOk - P6PerfMarkRebuildForceLsOk)
  s = s & ";force_lsretry=" & CStr(P6PerfRebuildForceLsRetry - P6PerfMarkRebuildForceLsRetry)
  s = s & ";rebuild_active=" & CStr(P6PerfRebuildActiveSet - P6PerfMarkRebuildActiveSet)
  s = s & ";rebuild_resid=" & CStr(P6PerfRebuildResidual - P6PerfMarkRebuildResidual)
  s = s & ";rebuild_stag=" & CStr(P6PerfRebuildStale - P6PerfMarkRebuildStale)
  s = s & ";rebuild_facinv=" & CStr(P6PerfRebuildFactorInvalid - P6PerfMarkRebuildFactorInvalid)
  s = s & ";cutback=" & CStr(P6PerfCutbackCount - P6PerfMarkCutback)
  s = s & ";plastic=" & CStr(P3ActivePlasticPointCount) & "/" & CStr(budget)
  s = s & ";umax=" & Format$(P6PerfMaxAbsDisp(), "0.000E+00")
  s = s & ";relres=" & Format$(RelativeResidualFree, "0.000E+00")
  s = s & ";energy=" & Format$(EnergyError, "0.000E+00")
  s = s & ";sym=" & Format$(MatrixSymmetryError, "0.000E+00")
  If P6PivotMinTrial < 1E+307 Then
    s = s & ";min_pivot=" & Format$(P6PivotMinTrial, "0.000E+00")
  Else
    s = s & ";min_pivot=NA"
  End If
  s = s & ";neg_pivot=" & CStr(P6PivotNegCountTrial)
  s = s & ";near0_pivot=" & CStr(P6PivotNearZeroCountTrial)
  s = s & ";t_fac_ms=" & Format$(P6ProfFactorMs - P6PerfMarkFactorMs, "0")
  s = s & ";t_sol_ms=" & Format$(P6ProfSolveMs - P6PerfMarkSolveMs, "0")
  s = s & ";t_tan_ms=" & Format$(P6ProfTangentMs - P6PerfMarkTangentMs, "0")
  s = s & ";ls_tries=" & CStr(P6LsTries - P6PerfMarkLsTries)
  s = s & ";ls_accept=" & CStr(P6LsAccept - P6PerfMarkLsAccept)
  s = s & ";ls_a1=" & CStr(P6LsA1 - P6PerfMarkLsA1)
  s = s & ";ls_a50=" & CStr(P6LsA50 - P6PerfMarkLsA50)
  s = s & ";ls_a25=" & CStr(P6LsA25 - P6PerfMarkLsA25)
  s = s & ";ls_a12=" & CStr(P6LsA12 - P6PerfMarkLsA12)
  s = s & ";ls_ok_a1=" & CStr(P6LsOkA1 - P6PerfMarkLsOkA1)
  s = s & ";ls_ok_a50=" & CStr(P6LsOkA50 - P6PerfMarkLsOkA50)
  s = s & ";ls_ok_a25=" & CStr(P6LsOkA25 - P6PerfMarkLsOkA25)
  s = s & ";ls_ok_a12=" & CStr(P6LsOkA12 - P6PerfMarkLsOkA12)
  If (P6LsAccept - P6PerfMarkLsAccept) > 0 Then
    s = s & ";ls_alpha_mean=" & Format$((P6LsAlphaSum - P6PerfMarkLsAlphaSum) / CDbl(P6LsAccept - P6PerfMarkLsAccept), "0.000")
  Else
    s = s & ";ls_alpha_mean=NA"
  End If
  If P6LsAlphaMinTrial < 1E+307 Then
    s = s & ";ls_alpha_min=" & Format$(P6LsAlphaMinTrial, "0.000")
  Else
    s = s & ";ls_alpha_min=NA"
  End If
  If (P6LsQCount - P6PerfMarkLsQCount) > 0 Then
    s = s & ";q_ls_mean=" & Format$((P6LsQSum - P6PerfMarkLsQSum) / CDbl(P6LsQCount - P6PerfMarkLsQCount), "0.000")
  Else
    s = s & ";q_ls_mean=NA"
  End If
  s = s & ";q_ls_n=" & CStr(P6LsQCount - P6PerfMarkLsQCount)
  s = s & ";q_ls_lt50=" & CStr(P6LsQLt50 - P6PerfMarkLsQLt50)
  s = s & ";q_ls_lt70=" & CStr(P6LsQLt70 - P6PerfMarkLsQLt70)
  s = s & ";q_ls_lt100=" & CStr(P6LsQLt100 - P6PerfMarkLsQLt100)
  s = s & ";q_ls_ge100=" & CStr(P6LsQGe100 - P6PerfMarkLsQGe100)
  If (P6LsQnCount - P6PerfMarkLsQnCount) > 0 Then
    s = s & ";q_next_mean=" & Format$((P6LsQnSum - P6PerfMarkLsQnSum) / CDbl(P6LsQnCount - P6PerfMarkLsQnCount), "0.000")
  Else
    s = s & ";q_next_mean=NA"
  End If
  s = s & ";q_next_n=" & CStr(P6LsQnCount - P6PerfMarkLsQnCount)
  s = s & ";q_next_lt70=" & CStr(P6LsQnLt70 - P6PerfMarkLsQnLt70)
  s = s & ";age0=" & CStr(P6TanSolve0 - P6PerfMarkTan0)
  s = s & ";age1=" & CStr(P6TanSolve1 - P6PerfMarkTan1)
  s = s & ";age2=" & CStr(P6TanSolve2 - P6PerfMarkTan2)
  s = s & ";age3p=" & CStr(P6TanSolve3 - P6PerfMarkTan3)
  s = s & ";age1_ok=" & CStr(P6Age1Success - P6PerfMarkAge1Ok)
  s = s & ";age1_resid=" & CStr(P6Age1Resid - P6PerfMarkAge1Resid)
  s = s & ";age1_stag=" & CStr(P6Age1Stag - P6PerfMarkAge1Stag)
  s = s & ";age1_active=" & CStr(P6Age1Active - P6PerfMarkAge1Active)
  s = s & ";age1_force=" & CStr(P6Age1Force - P6PerfMarkAge1Force)
  If P6PivotMaxTrial > 0# And P6PivotMinTrial < 1E+307 Then
    s = s & ";pivot_max=" & Format$(P6PivotMaxTrial, "0.000E+00")
    s = s & ";pivot_ratio=" & Format$(P6PivotMinTrial / P6PivotMaxTrial, "0.000E+00")
  Else
    s = s & ";pivot_max=NA;pivot_ratio=NA"
  End If
  s = s & ";ldlt_n=" & CStr(P6LdltN - P6PerfMarkLdltN)
  s = s & ";ldlt_ok=" & CStr(P6LdltOk - P6PerfMarkLdltOk)
  s = s & ";ldlt_fail=" & CStr(P6LdltFail - P6PerfMarkLdltFail)
  s = s & ";ldlt_ms=" & Format$(P6LdltMs - P6PerfMarkLdltMs, "0")
  If (P6LdltOk - P6PerfMarkLdltOk) > 0 Then
    s = s & ";ldlt_ex_max=" & Format$(P6LdltExMaxTrial, "0.000E+00")
    s = s & ";ldlt_er_max=" & Format$(P6LdltErMaxTrial, "0.000E+00")
    If P6LdltDRatioMinTrial < 1E+307 Then
      s = s & ";ldlt_d_ratio_min=" & Format$(P6LdltDRatioMinTrial, "0.000E+00")
    Else
      s = s & ";ldlt_d_ratio_min=NA"
    End If
  Else
    s = s & ";ldlt_ex_max=NA;ldlt_er_max=NA;ldlt_d_ratio_min=NA"
  End If
  s = s & ";ls50_tried=" & CStr(P6Ls50Tried - P6PerfMarkLs50Tried)
  s = s & ";ls50_ok=" & CStr(P6Ls50Ok - P6PerfMarkLs50Ok)
  s = s & ";ls50_resid=" & CStr(P6Ls50Resid - P6PerfMarkLs50Resid)
  s = s & ";ls50_stag=" & CStr(P6Ls50Stag - P6PerfMarkLs50Stag)
  s = s & ";ls50_worse=" & CStr(P6Ls50Worse - P6PerfMarkLs50Worse)
  s = s & ";ls50_saved=" & CStr(P6Ls50Saved - P6PerfMarkLs50Saved)
  s = s & ";ls50_gate_policy_reject=" & CStr(P6Ls50GatePolicy - P6PerfMarkLs50GatePolicy)
  s = s & ";ls50_gate_disabled=" & CStr(P6Ls50GateDisabled - P6PerfMarkLs50GateDisabled)
  s = s & ";ls50_gate_alpha_reject=" & CStr(P6Ls50GateAlpha - P6PerfMarkLs50GateAlpha)
  s = s & ";ls50_gate_qls_reject=" & CStr(P6Ls50GateQls - P6PerfMarkLs50GateQls)
  s = s & ";ls50_gate_cutback_reject=" & CStr(P6Ls50GateCutback - P6PerfMarkLs50GateCutback)
  s = s & ";ls50_gate_active_reject=" & CStr(P6Ls50GateActive - P6PerfMarkLs50GateActive)
  s = s & ";ls50_gate_residual_reject=" & CStr(P6Ls50GateResidual - P6PerfMarkLs50GateResidual)
  s = s & ";ls50_gate_factor_invalid=" & CStr(P6Ls50GateFactor - P6PerfMarkLs50GateFactor)
  s = s & ";BandLUKernel=" & P6BandKernelName()
  s = s & ";band_kernel=" & P6BandKernelName()
  s = s & ";factor_avg_ms=" & P6TrialFactorAvgText()
  P6PerfDeltaLine = s
End Function

Private Function P6TrialFactorAvgText() As String
  Dim nFac As Long
  Dim tFac As Double
  nFac = P6FactorizationCount - P6PerfMarkFactor
  tFac = P6ProfFactorMs - P6PerfMarkFactorMs
  If nFac > 0 Then
    P6TrialFactorAvgText = Format$(tFac / CDbl(nFac), "0.000")
  Else
    P6TrialFactorAvgText = "NA"
  End If
End Function


Public Sub P6PhysResetTrial()
  P6RefN = 0
  P6PhysLastUmax = 0#
  P6PhysMaxDu = 0#
  P6PhysDuInit = False
  P6PhysCrossLevel = 0
  P6PhysLastNewP = 0#
  P6PhysHaveNewP = False
  P6PhysReasonInc = -1
  P6PhysReasons = "|"
  P6LastQlsOk = False
  P6LastQnextOk = False
  P6PhysLambda = 0#
  P6PhysStep = 0#
  ResetPhysicalFailureMonitor
End Sub

Private Sub P6PhysReadMode()
  ' PHYSICAL_FAILURE_MODE is not a switch. GRAVITY is off and SRM is SHADOW at the solver entry.
  mPhysMode = "POLICY"
End Sub

Private Sub ResetPhysicalFailureMonitor()
  Dim i As Long
  mPhysMode = "POLICY"
  mPhysRefCount = 0
  mPhysRefReady = False
  mComplianceRef = 0#
  mPhysHitCount = 0
  mPhysCandidate = False
  mPhysWarnCount = 0
  mPhysRcMax = 0#
  mPhysRcMaxOk = False
  mPhysFailInc = 0
  mPhysFailLambda = 0#
  mPhysFailUmax = 0#
  mPhysFailUmaxH = 0#
  mPhysFailHave = False
  mPhysH = 0#
  mPhysRatioNow = 0#
  For i = 1 To PHYS_REF_N
    mPhysRefVals(i) = 0#
    P6Limit.RefValue(i) = 0#
  Next i
  P6Limit.RefCount = 0
  P6Limit.ComplianceRef = 0#
  P6Limit.RefReady = False
  P6Limit.RatioCurrent = 0#
  P6Limit.RatioMax = 0#
  P6Limit.WarningCount = 0
  P6Limit.ConsecutiveTrip = 0
  P6Limit.candidate = False
  P6Limit.FirstCandidateInc = 0
  P6Limit.FirstCandidateLambda = 0#
  P6Limit.FirstCandidateUmax = 0#
End Sub

Private Sub P6PhysTakeRef(ByVal cOk As Boolean, ByVal c20 As Double, ByVal du20 As Double, ByVal cutInWindow As Boolean)
  If mPhysRefReady Then Exit Sub
  If cutInWindow Then Exit Sub
  If Not cOk Then Exit Sub
  If du20 < 0# Then Exit Sub
  If c20 <= 0# Then Exit Sub
  If Not P6PhysFinite(c20) Then Exit Sub
  If mPhysRefCount >= PHYS_REF_N Then Exit Sub
  mPhysRefCount = mPhysRefCount + 1
  mPhysRefVals(mPhysRefCount) = c20
  If mPhysRefCount >= PHYS_REF_N Then
    mComplianceRef = P6Median3(mPhysRefVals(1), mPhysRefVals(2), mPhysRefVals(3))
    mPhysRefReady = True
  End If
End Sub

Private Function P6PhysFinite(ByVal v As Double) As Boolean
  P6PhysFinite = False
  If v <> v Then Exit Function
  If Abs(v) > 1E+20 Then Exit Function
  P6PhysFinite = True
End Function

Private Function P6PhysMaxD(ByVal a As Double, ByVal b As Double) As Double
  If a > b Then
    P6PhysMaxD = a
  Else
    P6PhysMaxD = b
  End If
End Function

Private Sub P6SwapDouble(ByRef a As Double, ByRef b As Double)
  Dim t As Double
  t = a
  a = b
  b = t
End Sub

Private Function P6Median3(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
  If a > b Then P6SwapDouble a, b
  If b > c Then P6SwapDouble b, c
  If a > b Then P6SwapDouble a, b
  P6Median3 = b
End Function

Private Function P6PhysicalStatus() As String
  If mPhysMode = "OFF" Or Len(mPhysMode) = 0 Then
    P6PhysicalStatus = "NOT_ASSESSED"
  ElseIf mPhysCandidate Then
    P6PhysicalStatus = "FAILURE_CANDIDATE"
  ElseIf mPhysRefReady Then
    P6PhysicalStatus = "OK"
  Else
    P6PhysicalStatus = "INSUFFICIENT"
  End If
End Function

Private Function P6WouldPhysText() As String
  If mPhysMode = "OFF" Or Len(mPhysMode) = 0 Then
    P6WouldPhysText = "NA"
  ElseIf mPhysCandidate Then
    P6WouldPhysText = "1"
  Else
    P6WouldPhysText = "0"
  End If
End Function

Private Function P6PhysSummaryTail() As String
  Dim s As String
  If Len(mPhysMode) = 0 Then
    s = ",NA,NA,0,NA,NA,NA,NA,NA,NA"
    P6PhysSummaryTail = s
    Exit Function
  End If
  s = "," & FEMIoCsv(mPhysMode)
  If mPhysRefReady Then
    s = s & "," & Format$(mComplianceRef, "0.000E+00")
  Else
    s = s & ",NA"
  End If
  s = s & "," & CStr(mPhysWarnCount)
  If mPhysMode = "OFF" Then
    s = s & ",NA"
  ElseIf mPhysCandidate Then
    s = s & ",1"
  Else
    s = s & ",0"
  End If
  If mPhysFailHave Then
    s = s & "," & CStr(mPhysFailInc)
    s = s & "," & Format$(mPhysFailLambda, "0.000E+00")
    s = s & "," & Format$(mPhysFailUmax, "0.000E+00")
    s = s & "," & Format$(mPhysFailUmaxH, "0.000E+00")
  Else
    s = s & ",NA,NA,NA,NA"
  End If
  If mPhysRcMaxOk Then
    s = s & "," & Format$(mPhysRcMax, "0.000E+00")
  Else
    s = s & ",NA"
  End If
  P6PhysSummaryTail = s
End Function

Private Function P6PhysAlarmDetail(ByVal lambdaNow As Double, ByVal umax As Double, ByVal h As Double, ByVal du20 As Double, ByVal duOk As Boolean, ByVal du20h As Double, ByVal du20hOk As Boolean, ByVal dl20 As Double, ByVal c20 As Double, ByVal cOk As Boolean, ByVal cRatio As Double, ByVal newP20 As Long, ByVal newPOk As Boolean, ByVal dF20 As Long, ByVal dR20 As Long, ByVal dN20 As Long, ByVal okN20 As Boolean) As String
  Dim s As String
  Dim budget As Long
  s = "mode=SHADOW"
  s = s & ";fs=" & Format$(P3CurrentStrengthFactor, "0.000")
  s = s & ";inc=" & CStr(CurrentIncrement)
  s = s & ";lambda=" & Format$(lambdaNow, "0.000E+00")
  s = s & ";umax=" & Format$(umax, "0.000E+00")
  If h > 0# Then
    s = s & ";umax_over_H=" & Format$(umax / h, "0.000E+00")
  Else
    s = s & ";umax_over_H=NA"
  End If
  s = s & ";du20=" & P6NumText(du20, duOk)
  s = s & ";du20_over_H=" & P6NumText(du20h, du20hOk)
  s = s & ";dlambda20=" & P6NumText(dl20, duOk)
  s = s & ";compliance20=" & P6NumText(c20, cOk)
  If mPhysRefReady Then
    s = s & ";compliance_ref=" & Format$(mComplianceRef, "0.000E+00")
  Else
    s = s & ";compliance_ref=NA"
  End If
  s = s & ";compliance_ratio=" & Format$(cRatio, "0.000E+00")
  budget = P6PerfPlasticBudget()
  If budget > 0 Then
    s = s & ";plastic_gp_ratio=" & Format$(CDbl(P3ActivePlasticPointCount) / CDbl(budget), "0.000E+00")
  Else
    s = s & ";plastic_gp_ratio=NA"
  End If
  s = s & ";new_plastic_gp20=" & P6NumText(CDbl(newP20), newPOk)
  s = s & ";relres=" & Format$(RelativeResidualFree, "0.000E+00")
  If okN20 And dN20 > 0 Then
    s = s & ";factor_ratio20=" & Format$(CDbl(dF20) / CDbl(dN20), "0.000")
    s = s & ";reuse_ratio20=" & Format$(CDbl(dR20) / CDbl(dN20), "0.000")
  Else
    s = s & ";factor_ratio20=NA;reuse_ratio20=NA"
  End If
  s = s & ";ls_alpha=" & Format$(P6LsLastAlpha, "0.000")
  If P6LastQlsOk Then
    s = s & ";q_ls=" & Format$(P6LastQls, "0.000E+00")
  Else
    s = s & ";q_ls=NA"
  End If
  If P6LastQnextOk Then
    s = s & ";q_next=" & Format$(P6LastQnext, "0.000E+00")
  Else
    s = s & ";q_next=NA"
  End If
  P6PhysAlarmDetail = s
End Function

Private Function P6NumericalStatus(ByVal statusText As String) As String
  Dim s As String
  s = UCase$(statusText)
  If InStr(s, "FAIL") > 0 Then
    P6NumericalStatus = "FAILED"
  ElseIf InStr(s, "PASS") > 0 Then
    P6NumericalStatus = "CONVERGED"
  Else
    P6NumericalStatus = "UNKNOWN"
  End If
End Function

Private Function P6LambdaFinalText() As String
  Dim idx As Long
  If P6RollCount < 1 Then
    P6LambdaFinalText = "NA"
  Else
    idx = (P6RollCount - 1) Mod P6_ROLL_CAP
    P6LambdaFinalText = Format$(P6RollLambda(idx), "0.000E+00")
  End If
End Function

Public Function P6ModelHeight() As Double
  Dim i As Long, y0 As Double, y1 As Double, y As Double
  P6ModelHeight = 0#
  On Error GoTo UseScale
  If P6InputNodeCount < 1 Then GoTo UseScale
  y0 = P6InputNodeY(1)
  y1 = y0
  For i = 1 To P6InputNodeCount
    y = P6InputNodeY(i)
    If y < y0 Then y0 = y
    If y > y1 Then y1 = y
  Next i
  P6ModelHeight = y1 - y0
  If P6ModelHeight > 0# Then Exit Function
UseScale:
  Err.Clear
  If P5ModelScale > 0# Then P6ModelHeight = P5ModelScale
End Function

Private Function P6UmaxOverHText() As String
  Dim h As Double
  h = P6ModelHeight()
  If h <= 0# Then
    P6UmaxOverHText = "NA"
  Else
    P6UmaxOverHText = Format$(P6PerfMaxAbsDisp() / h, "0.000E+00")
  End If
End Function

Private Function P6PlasticRatioText() As String
  Dim budget As Long
  budget = P6PerfPlasticBudget()
  If budget <= 0 Then
    P6PlasticRatioText = "NA"
  Else
    P6PlasticRatioText = Format$(CDbl(P3ActivePlasticPointCount) / CDbl(budget), "0.000E+00")
  End If
End Function

Private Function P6NumText(ByVal v As Double, ByVal ok As Boolean) As String
  If Not ok Then
    P6NumText = "NA"
  Else
    P6NumText = Format$(v, "0.000E+00")
  End If
End Function

Private Function P6MedianOf(ByRef src() As Double, ByVal n As Long) As Double
  Dim a() As Double, i As Long, j As Long
  Dim key As Double
  P6MedianOf = 0#
  If n < 1 Then Exit Function
  ReDim a(0 To n - 1)
  For i = 0 To n - 1
    a(i) = src(i)
  Next i
  For i = 1 To n - 1
    key = a(i)
    j = i - 1
    Do While j >= 0
      If a(j) <= key Then Exit Do
      a(j + 1) = a(j)
      j = j - 1
    Loop
    a(j + 1) = key
  Next i
  If (n Mod 2) = 1 Then
    P6MedianOf = a(n \ 2)
  Else
    P6MedianOf = 0.5 * (a(n \ 2 - 1) + a(n \ 2))
  End If
End Function

Private Function P6RollD(ByRef arr() As Double, ByVal back As Long, ByRef ok As Boolean) As Double
  Dim idxNow As Long, idxOld As Long
  ok = False
  P6RollD = 0#
  If P6RollCount <= back Or back >= P6_ROLL_CAP Then Exit Function
  idxNow = (P6RollCount - 1) Mod P6_ROLL_CAP
  idxOld = (P6RollCount - 1 - back) Mod P6_ROLL_CAP
  P6RollD = arr(idxNow) - arr(idxOld)
  ok = True
End Function

Private Function P6RollL(ByRef arr() As Long, ByVal back As Long, ByRef ok As Boolean) As Long
  Dim idxNow As Long, idxOld As Long
  ok = False
  P6RollL = 0
  If P6RollCount <= back Or back >= P6_ROLL_CAP Then Exit Function
  idxNow = (P6RollCount - 1) Mod P6_ROLL_CAP
  idxOld = (P6RollCount - 1 - back) Mod P6_ROLL_CAP
  P6RollL = arr(idxNow) - arr(idxOld)
  ok = True
End Function

Private Function P6FindParent(ByRef parent() As Long, ByVal x As Long) As Long
  Dim r As Long, p As Long
  r = x
  Do While parent(r) <> r
    r = parent(r)
  Loop
  Do While parent(x) <> r
    p = parent(x)
    parent(x) = r
    x = p
  Loop
  P6FindParent = r
End Function

Private Sub P6UnionParent(ByRef parent() As Long, ByVal a As Long, ByVal b As Long)
  Dim ra As Long, rb As Long
  ra = P6FindParent(parent, a)
  rb = P6FindParent(parent, b)
  If ra <> rb Then parent(rb) = ra
End Sub

Private Sub P6PlasticStats(ByRef nElem As Long, ByRef largest As Long, ByRef maxEq As Double, ByRef meanEq As Double)
  Dim k As Long, im As Long, i As Long, yielded As Long, nEq As Long
  Dim ex As Double, ey As Double, ez As Double, gxy As Double, eq As Double, sumEq As Double
  Dim parent() As Long, head() As Long, sz() As Long
  Dim nodeMax As Long, dof As Long, nd As Long, r As Long
  Dim isP() As Boolean
  nElem = 0
  largest = 0
  maxEq = 0#
  meanEq = 0#
  nEq = 0
  sumEq = 0#
  If NumberOfElement < 1 Then Exit Sub
  ReDim isP(0 To NumberOfElement - 1)
  ReDim parent(0 To NumberOfElement - 1)
  For k = 0 To NumberOfElement - 1
    parent(k) = k
    isP(k) = False
    If P3IsElementActive(k) Then
      If Not Elem(k).IsJoint Then
        yielded = 0
        For im = 0 To 3
          If Elem(k).P2Yielded(im) Then
            yielded = yielded + 1
            ex = Elem(k).P2PlasticStrain(0, im)
            ey = Elem(k).P2PlasticStrain(1, im)
            gxy = Elem(k).P2PlasticStrain(2, im)
            ez = Elem(k).P2PlasticStrain(3, im)
            eq = Sqr((2# / 3#) * (ex * ex + ey * ey + ez * ez + 0.5 * gxy * gxy))
            If eq > maxEq Then maxEq = eq
            sumEq = sumEq + eq
            nEq = nEq + 1
          End If
        Next im
        If yielded >= 2 Then
          isP(k) = True
          nElem = nElem + 1
        End If
      End If
    End If
  Next k
  If nEq > 0 Then meanEq = sumEq / CDbl(nEq)
  nodeMax = NumberOfNode
  If lastDof \ 2 > nodeMax Then nodeMax = lastDof \ 2
  If nodeMax < 0 Then Exit Sub
  ReDim head(0 To nodeMax)
  For i = 0 To nodeMax
    head(i) = -1
  Next i
  For k = 0 To NumberOfElement - 1
    If isP(k) Then
      For i = 0 To 7
        dof = Elem(k).ElNode(i)
        If dof >= 0 Then
          nd = dof \ 2
          If nd >= 0 And nd <= nodeMax Then
            If head(nd) >= 0 Then
              P6UnionParent parent, k, head(nd)
            Else
              head(nd) = k
            End If
          End If
        End If
      Next i
    End If
  Next k
  ReDim sz(0 To NumberOfElement - 1)
  For k = 0 To NumberOfElement - 1
    If isP(k) Then
      r = P6FindParent(parent, k)
      sz(r) = sz(r) + 1
      If sz(r) > largest Then largest = sz(r)
    End If
  Next k
End Sub

Public Sub P6Ls50Event(ByVal eventName As String, ByVal phaseText As String)
  Dim s As String
  s = "phase=" & phaseText
  s = s & ";newton=" & CStr(P6PerfNewtonCount)
  s = s & ";lambda=" & Format$(P6PhysLambda, "0.000E+00")
  s = s & ";umax=" & Format$(P6PerfMaxAbsDisp(), "0.000E+00")
  If P6LastQlsOk Then
    s = s & ";q_ls=" & Format$(P6LastQls, "0.000E+00")
  Else
    s = s & ";q_ls=NA"
  End If
  If P6LastQnextOk Then
    s = s & ";q_next=" & Format$(P6LastQnext, "0.000E+00")
  Else
    s = s & ";q_next=NA"
  End If
  s = s & ";tangent_age=" & CStr(P3TangentAge)
  P6SolverEvent eventName, s
End Sub

Public Sub P6SolverEventOnce(ByVal eventName As String, ByVal detailText As String)
  Dim token As String
  If femIoMode = "OFF" Then Exit Sub
  token = "EV:" & eventName
  If P6PhysReasonInc = CurrentIncrement Then
    If InStr(1, P6PhysReasons, "|" & token & "|") > 0 Then Exit Sub
  Else
    P6PhysReasonInc = CurrentIncrement
    P6PhysReasons = "|"
  End If
  P6PhysReasons = P6PhysReasons & token & "|"
  P6SolverEvent eventName, detailText
End Sub

Public Sub P6PhysEmitOnce(ByVal reason As String)
  ' Kept so older call sites compile. Non-periodic events are not logged.
End Sub

Public Sub P6PhysOnRebuild()
  If P3LastRebuildWhy = "RESID" Or P3LastRebuildWhy = "STAG" Then
    P6SolverEventOnce "RESID_STAGNATION", "why=" & P3LastRebuildWhy
  End If
End Sub

Public Sub P6PhysWatchIncrement(ByVal stepSize As Double, ByVal duMax As Double, ByVal lambdaNow As Double)
  ' Per-increment watch is off. The limit monitor runs on PERIODIC rows only.
End Sub

Private Sub P6PushLimitState()
  Dim i As Long
  P6Limit.RefCount = mPhysRefCount
  P6Limit.ComplianceRef = mComplianceRef
  P6Limit.RefReady = mPhysRefReady
  P6Limit.RatioCurrent = mPhysRatioNow
  If mPhysRcMaxOk Then
    P6Limit.RatioMax = mPhysRcMax
  Else
    P6Limit.RatioMax = 0#
  End If
  P6Limit.WarningCount = mPhysWarnCount
  P6Limit.ConsecutiveTrip = mPhysHitCount
  P6Limit.candidate = mPhysCandidate
  P6Limit.FirstCandidateInc = mPhysFailInc
  P6Limit.FirstCandidateLambda = mPhysFailLambda
  P6Limit.FirstCandidateUmax = mPhysFailUmax
  For i = 1 To PHYS_REF_N
    P6Limit.RefValue(i) = mPhysRefVals(i)
  Next i
End Sub

Private Sub P6WriteFixedTrace(ByVal lambdaNow As Double, ByVal umax As Double, ByVal du20 As Double, ByVal ok20 As Boolean, ByVal du80 As Double, ByVal ok80 As Boolean, ByVal dl20 As Double, ByVal dl80 As Double, ByVal c20 As Double, ByVal cOk As Boolean, ByVal cRatio As Double, ByVal ratioOk As Boolean, ByVal dP20 As Long, ByVal okP20 As Boolean, ByVal dF20 As Long, ByVal dR20 As Long, ByVal dD20 As Long, ByVal dN20 As Long, ByVal dA20 As Long, ByVal okN20 As Boolean)
  Dim headerText As String, lineText As String
  Dim budget As Long
  headerText = P6CsvHeaderTrace()
  lineText = FEM_BUILD_STAMP
  lineText = lineText & "," & Format$(P3CurrentStrengthFactor, "0.000")
  lineText = lineText & "," & CStr(CurrentIncrement)
  lineText = lineText & "," & Format$(lambdaNow, "0.000E+00")
  lineText = lineText & "," & P6NumText(dl20, ok20)
  lineText = lineText & "," & P6NumText(dl80, ok80)
  lineText = lineText & "," & Format$(umax, "0.000E+00")
  lineText = lineText & "," & P6NumText(du20, ok20)
  lineText = lineText & "," & P6NumText(du80, ok80)
  lineText = lineText & "," & P6NumText(c20, cOk)
  If mPhysRefReady Then
    lineText = lineText & "," & Format$(mComplianceRef, "0.000E+00")
  Else
    lineText = lineText & ",NA"
  End If
  If ratioOk And mPhysRefReady Then
    lineText = lineText & "," & Format$(cRatio, "0.000E+00")
  Else
    lineText = lineText & ",NA"
  End If
  budget = P6PerfPlasticBudget()
  If budget > 0 Then
    lineText = lineText & "," & Format$(CDbl(P3ActivePlasticPointCount) / CDbl(budget), "0.000E+00")
  Else
    lineText = lineText & ",NA"
  End If
  lineText = lineText & "," & P6NumText(CDbl(dP20), okP20)
  If okN20 And dN20 > 0 Then
    lineText = lineText & "," & Format$(CDbl(dF20) / CDbl(dN20), "0.000")
    lineText = lineText & "," & Format$(CDbl(dR20) / CDbl(dN20), "0.000")
  Else
    lineText = lineText & ",NA,NA"
  End If
  If okN20 And dA20 > 0 Then
    lineText = lineText & "," & Format$(CDbl(dD20) / CDbl(dA20), "0.000")
  Else
    lineText = lineText & ",NA"
  End If
  FEMIoAppendCsv "perf_trace.csv", P6CsvTraceHeader, headerText, lineText
End Sub

Public Sub P6PhysConsider(ByVal reason As String, ByVal forceWrite As Boolean, ByVal stepSize As Double, ByVal duMax As Double, ByVal lambdaNow As Double)
  Dim umax As Double, h As Double
  Dim du20 As Double, du80 As Double, dl20 As Double, dl80 As Double
  Dim ok20 As Boolean, ok80 As Boolean, okN20 As Boolean
  Dim okP20 As Boolean, okCut As Boolean, cutInWindow As Boolean
  Dim dN20 As Long, dF20 As Long, dR20 As Long, dA20 As Long, dD20 As Long, dP20 As Long
  Dim dCut20 As Long
  Dim c20 As Double, cRatio As Double
  Dim cOk As Boolean, ratioOk As Boolean
  Dim du20h As Double, du20hOk As Boolean
  Dim warnNow As Boolean, failNow As Boolean
  Dim kindText As String
  If femIoMode = "OFF" Then Exit Sub
  If reason <> "PERIODIC" Then Exit Sub
  umax = P6PerfMaxAbsDisp()
  If mPhysH <= 0# Then mPhysH = P6ModelHeight()
  h = mPhysH
  du20 = P6RollD(P6RollUmax, 20, ok20)
  du80 = P6RollD(P6RollUmax, 80, ok80)
  dl20 = P6RollD(P6RollLambda, 20, ok20)
  dl80 = P6RollD(P6RollLambda, 80, ok80)
  dN20 = P6RollL(P6RollNewton, 20, okN20)
  dF20 = P6RollL(P6RollFactor, 20, okN20)
  dR20 = P6RollL(P6RollReuse, 20, okN20)
  dA20 = P6RollL(P6RollAccept, 20, okN20)
  dD20 = P6RollL(P6RollDamped, 20, okN20)
  dP20 = P6RollL(P6RollPlastic, 20, okP20)
  dCut20 = P6RollL(P6RollCut, 20, okCut)
  cutInWindow = True
  If okCut Then
    If dCut20 <= 0 Then cutInWindow = False
  End If
  cOk = False
  c20 = 0#
  If ok20 And h > 0# Then
    If dl20 > PHYS_DLAM_EPS Then
      If P6PhysFinite(du20) And P6PhysFinite(dl20) Then
        c20 = (du20 / h) / dl20
        If P6PhysFinite(c20) Then cOk = True
      End If
    End If
  End If
  ratioOk = False
  cRatio = 0#
  warnNow = False
  failNow = False
  mPhysRatioNow = 0#
  If P6LimitMonitorOn Then
    If Not mPhysRefReady Then P6PhysTakeRef cOk, c20, du20, cutInWindow
    If mPhysRefReady And cOk And c20 > 0# Then
      cRatio = c20 / P6PhysMaxD(mComplianceRef, 1E-30)
      ratioOk = P6PhysFinite(cRatio)
      If ratioOk Then
        mPhysRatioNow = cRatio
        If (Not mPhysRcMaxOk) Or cRatio > mPhysRcMax Then
          mPhysRcMax = cRatio
          mPhysRcMaxOk = True
        End If
        If cRatio >= LIMIT_TRIP_RATIO Then
          mPhysHitCount = mPhysHitCount + 1
        Else
          mPhysHitCount = 0
        End If
        If cRatio >= LIMIT_WARN_RATIO Then
          mPhysWarnCount = mPhysWarnCount + 1
          warnNow = True
        End If
        If mPhysHitCount >= LIMIT_TRIP_COUNT And Not mPhysCandidate Then
          mPhysCandidate = True
          mPhysFailHave = True
          mPhysFailInc = CurrentIncrement
          mPhysFailLambda = lambdaNow
          mPhysFailUmax = umax
          If h > 0# Then
            mPhysFailUmaxH = umax / h
          Else
            mPhysFailUmaxH = 0#
          End If
          failNow = True
        End If
      End If
    End If
    If warnNow Or failNow Then
      du20hOk = False
      du20h = 0#
      If ok20 And h > 0# And P6PhysFinite(du20) Then
        du20h = du20 / h
        If P6PhysFinite(du20h) Then du20hOk = True
      End If
      kindText = P6PhysAlarmDetail(lambdaNow, umax, h, du20, ok20, du20h, du20hOk, dl20, c20, cOk, cRatio, dP20, okP20, dF20, dR20, dN20, okN20)
      If warnNow Then P6SolverEvent "LIMIT_WARN", kindText
      If failNow Then P6SolverEvent "LIMIT_STATE_CANDIDATE", kindText
    End If
  End If
  P6PushLimitState
  P6WriteFixedTrace lambdaNow, umax, du20, ok20, du80, ok80, dl20, dl80, c20, cOk, cRatio, ratioOk, dP20, okP20, dF20, dR20, dD20, dN20, dA20, okN20
End Sub

Public Sub P6PerfEmit(ByVal placeText As String, ByVal statusText As String)
  FEMAppendRunLog "PERF", statusText, P6PerfDeltaLine()
  If femIoMode = "OFF" Then
    P6FlushBandBench
    Exit Sub
  End If
  FEMIoLog placeText, "PERF", P6PerfDeltaLine()
  P6WriteTrialCsv statusText
  P6FlushBandBench
End Sub

Public Sub P6PerfEmitReuse(ByVal statusText As String)
  Dim s As String
  Dim budget As Long
  budget = P6PerfPlasticBudget()
  s = "newton=0;inc=0;inc_now=" & CStr(CurrentIncrement)
  s = s & ";iter_now=" & CStr(CurrentIteration)
  s = s & ";factor=0;reuse=0;band=0;modnewton=0;k_build=0"
  s = s & ";rebuild_first=0;rebuild_force=0;force_start=0;force_cut=0;force_lscut=0;force_lsok=0;force_lsretry=0"
  s = s & ";rebuild_active=0;rebuild_resid=0;rebuild_stag=0;rebuild_facinv=0;cutback=0"
  s = s & ";plastic=" & CStr(P3ActivePlasticPointCount) & "/" & CStr(budget)
  s = s & ";umax=" & Format$(P6PerfMaxAbsDisp(), "0.000E+00")
  s = s & ";relres=" & Format$(RelativeResidualFree, "0.000E+00")
  s = s & ";energy=" & Format$(EnergyError, "0.000E+00")
  s = s & ";sym=" & Format$(MatrixSymmetryError, "0.000E+00")
  s = s & ";min_pivot=NA;neg_pivot=0;near0_pivot=0;t_fac_ms=0;t_sol_ms=0;t_tan_ms=0;reused=1"
  s = s & ";ls_tries=0;ls_accept=0;ls_a1=0;ls_a50=0;ls_a25=0;ls_a12=0"
  s = s & ";ls_ok_a1=0;ls_ok_a50=0;ls_ok_a25=0;ls_ok_a12=0;ls_alpha_mean=NA;ls_alpha_min=NA"
  s = s & ";BandLUKernel=" & P6BandKernelName() & ";band_kernel=" & P6BandKernelName() & ";factor_avg_ms=NA"
  FEMAppendRunLog "PERF", statusText, s
  If femIoMode = "OFF" Then
    P6FlushBandBench
    Exit Sub
  End If
  FEMIoLog "SRM試行", "PERF", s
  P6WriteTrialCsv statusText
  P6FlushBandBench
End Sub

Public Sub P6PerfTrace(ByVal eventName As String, ByVal stepSize As Double, ByVal consecutiveCutback As Long, ByVal duMax As Double, ByVal lambdaNow As Double, ByVal lambdaTarget As Double, ByVal traceStage As Long, ByVal traceKind As String)
  Dim s As String
  Dim minStepRatio As Double
  Dim remaining As Double
  If femIoMode = "OFF" Then Exit Sub
  If P3_MIN_STEP_FACTOR > 0# Then
    minStepRatio = stepSize / P3_MIN_STEP_FACTOR
  Else
    minStepRatio = 0#
  End If
  remaining = 1# - lambdaNow
  If remaining < 0# Then remaining = 0#
  s = P6PerfDeltaLine()
  s = s & ";event=" & eventName
  s = s & ";step=" & Format$(stepSize, "0.000E+00")
  s = s & ";min_step_ratio=" & Format$(minStepRatio, "0.000E+00")
  s = s & ";lambda_now=" & Format$(lambdaNow, "0.000E+00")
  s = s & ";lambda_tgt=" & Format$(lambdaTarget, "0.000E+00")
  s = s & ";remaining=" & Format$(remaining, "0.000E+00")
  s = s & ";consec_cut=" & CStr(consecutiveCutback)
  s = s & ";du_max=" & Format$(duMax, "0.000E+00")
  s = s & ";ls_alpha=" & Format$(P6LsLastAlpha, "0.000")
  If eventName = "INC" Then
    s = s & P6RollWindowText(lambdaNow)
  Else
    s = s & ";dlambda20=NA;dlambda80=NA;factor_ratio80=NA;reuse_ratio80=NA;damped_ls_ratio80=NA"
  End If
  FEMIoWriteLine "PERF_TRACE", eventName, s, traceStage, traceKind
  P6WriteTraceCsv eventName, traceStage, traceKind, stepSize, lambdaNow, lambdaTarget, remaining, duMax, consecutiveCutback
  If eventName = "INC" Then
    P6PhysConsider "PERIODIC", True, stepSize, duMax, lambdaNow
  ElseIf eventName = "CUTBACK" Then
    P6PhysConsider "CUTBACK", True, stepSize, duMax, lambdaNow
  End If
End Sub

Private Sub P6WriteTraceCsv(ByVal eventName As String, ByVal traceStage As Long, ByVal traceKind As String, ByVal stepSize As Double, ByVal lambdaNow As Double, ByVal lambdaTarget As Double, ByVal remaining As Double, ByVal duMax As Double, ByVal consecutiveCutback As Long)
  ' perf_trace.csv is the 20-increment limit trace written by P6WriteFixedTrace.
End Sub

Private Sub P6PerfWriteSummary()
  Dim folderPath As String, summaryPath As String, fileNum As Integer
  Dim elapsed As Double, budget As Long, avgReuse As Double
  Dim lineText As String
  If P6PerfSummaryWritten Then Exit Sub
  AdaptiveV2aEndAttempt False, "RUN_CLOSE"
  IncrementLogClosePending "RUN_CLOSE"
  FEMIoFlushPolicy "RUN_CLOSE"
  P6PerfSummaryWritten = True
  If femIoMode = "OFF" Then Exit Sub
  elapsed = Timer - femIoStartedAt
  If elapsed < 0# Then elapsed = elapsed + 86400#
  budget = P6PerfPlasticBudget()
  If P6FactorizationCount > 0 Then
    avgReuse = CDbl(P6FactorizationReuseCount) / CDbl(P6FactorizationCount)
  Else
    avgReuse = 0#
  End If
  lineText = "ElapsedTotal=" & Format$(elapsed, "0.000") & " sec" & vbCrLf
  lineText = lineText & "SRMTrialCount=" & CStr(P3SrmTrialCount) & vbCrLf
  lineText = lineText & "Fs1ReuseCount=" & CStr(P6SrmFs1ReuseCount) & vbCrLf
  lineText = lineText & "FOS_PASS=" & Format$(P3FosPass, "0.000") & vbCrLf
  If P3FosBracket Then
    lineText = lineText & "FOS_FAIL=" & Format$(P3FosFail, "0.000") & vbCrLf
    lineText = lineText & "FOS_MID=" & Format$(P3FosMid, "0.000") & vbCrLf
    lineText = lineText & "FOS_WIDTH=" & Format$(P3FosWidth, "0.000") & vbCrLf
    lineText = lineText & "FOS=" & Format$(P3FosMid, "0.000") & " ± " & Format$(P3FosWidth * 0.5, "0.000") & vbCrLf
  Else
    lineText = lineText & "FOS_FAIL=NA" & vbCrLf
    lineText = lineText & "FOS_MID=" & Format$(P3FosPass, "0.000") & vbCrLf
    lineText = lineText & "FOS_WIDTH=NA" & vbCrLf
  End If
  lineText = lineText & AccelSummary()
  lineText = lineText & "NewtonIterationCount=" & CStr(P6PerfNewtonCount) & vbCrLf
  lineText = lineText & "LoadIncrementCount=" & CStr(P6PerfIncrementCount) & vbCrLf
  lineText = lineText & "CutbackCount=" & CStr(P6PerfCutbackCount) & vbCrLf
  lineText = lineText & "FactorizationCount=" & CStr(P6FactorizationCount) & vbCrLf
  lineText = lineText & "FactorizationReuseCount=" & CStr(P6FactorizationReuseCount) & vbCrLf
  lineText = lineText & "AVG_FACTOR_REUSE=" & Format$(avgReuse, "0.000") & vbCrLf
  lineText = lineText & "BandSolveCallCount=" & CStr(P6BandSolveCallCount) & vbCrLf
  lineText = lineText & "TangentBuildCount=" & CStr(P6ProfTangentCount) & vbCrLf
  lineText = lineText & "TangentReuseCount=" & CStr(P6PerfModNewtonSkips) & vbCrLf
  lineText = lineText & "REBUILD_FIRST_ITER=" & CStr(P6PerfRebuildFirst) & vbCrLf
  lineText = lineText & "REBUILD_FORCE=" & CStr(P6PerfRebuildForce) & vbCrLf
  lineText = lineText & "REBUILD_FORCE_START=" & CStr(P6PerfRebuildForceStart) & vbCrLf
  lineText = lineText & "REBUILD_FORCE_CUTBACK=" & CStr(P6PerfRebuildForceCutback) & vbCrLf
  lineText = lineText & "REBUILD_FORCE_LSCUT=" & CStr(P6PerfRebuildForceLsCut) & vbCrLf
  lineText = lineText & "REBUILD_FORCE_LSOK=" & CStr(P6PerfRebuildForceLsOk) & vbCrLf
  lineText = lineText & "REBUILD_FORCE_LSRETRY=" & CStr(P6PerfRebuildForceLsRetry) & vbCrLf
  lineText = lineText & "REBUILD_ACTIVESET=" & CStr(P6PerfRebuildActiveSet) & vbCrLf
  lineText = lineText & "REBUILD_RESIDUAL=" & CStr(P6PerfRebuildResidual) & vbCrLf
  lineText = lineText & "REBUILD_STAGNATION=" & CStr(P6PerfRebuildStale) & vbCrLf
  lineText = lineText & "REBUILD_FACTOR_INVALID=" & CStr(P6PerfRebuildFactorInvalid) & vbCrLf
  lineText = lineText & "MaxDisp=" & Format$(P6PerfMaxAbsDisp(), "0.000E+00") & vbCrLf
  lineText = lineText & "RelativeResidualFree=" & Format$(RelativeResidualFree, "0.000E+00") & vbCrLf
  lineText = lineText & "EnergyError=" & Format$(EnergyError, "0.000E+00") & vbCrLf
  lineText = lineText & "MatrixSymmetryError=" & Format$(MatrixSymmetryError, "0.000E+00") & vbCrLf
  If P6PivotMin < 1E+307 Then
    lineText = lineText & "MinPivot=" & Format$(P6PivotMin, "0.000E+00") & vbCrLf
  Else
    lineText = lineText & "MinPivot=NA" & vbCrLf
  End If
  lineText = lineText & "NegativePivotCount=" & CStr(P6PivotNegCount) & vbCrLf
  lineText = lineText & "NearZeroPivotCount=" & CStr(P6PivotNearZeroCount) & vbCrLf
  If P6PivotMax > 0# And P6PivotMin < 1E+307 Then
    lineText = lineText & "PivotMax=" & Format$(P6PivotMax, "0.000E+00") & vbCrLf
    lineText = lineText & "PivotRatio=" & Format$(P6PivotMin / P6PivotMax, "0.000E+00") & vbCrLf
  Else
    lineText = lineText & "PivotMax=NA" & vbCrLf
    lineText = lineText & "PivotRatio=NA" & vbCrLf
  End If
  lineText = lineText & "LsTries=" & CStr(P6LsTries) & vbCrLf
  lineText = lineText & "LsAccept=" & CStr(P6LsAccept) & vbCrLf
  lineText = lineText & "LsA1=" & CStr(P6LsA1) & vbCrLf
  lineText = lineText & "LsA50=" & CStr(P6LsA50) & vbCrLf
  lineText = lineText & "LsA25=" & CStr(P6LsA25) & vbCrLf
  lineText = lineText & "LsA12=" & CStr(P6LsA12) & vbCrLf
  lineText = lineText & "LsOkA1=" & CStr(P6LsOkA1) & vbCrLf
  lineText = lineText & "LsOkA50=" & CStr(P6LsOkA50) & vbCrLf
  lineText = lineText & "LsOkA25=" & CStr(P6LsOkA25) & vbCrLf
  lineText = lineText & "LsOkA12=" & CStr(P6LsOkA12) & vbCrLf
  If P6LsAccept > 0 Then
    lineText = lineText & "LsAlphaMean=" & Format$(P6LsAlphaSum / CDbl(P6LsAccept), "0.000") & vbCrLf
  Else
    lineText = lineText & "LsAlphaMean=NA" & vbCrLf
  End If
  If P6LsAlphaMin < 1E+307 Then
    lineText = lineText & "LsAlphaMin=" & Format$(P6LsAlphaMin, "0.000") & vbCrLf
  Else
    lineText = lineText & "LsAlphaMin=NA" & vbCrLf
  End If
  lineText = lineText & "PlasticPoints=" & CStr(P3ActivePlasticPointCount) & "/" & CStr(budget) & vbCrLf
  lineText = lineText & "TimeFactorization_ms=" & Format$(P6ProfFactorMs, "0") & vbCrLf
  lineText = lineText & "TimeLinearSolve_ms=" & Format$(P6ProfSolveMs, "0") & vbCrLf
  lineText = lineText & "TimeTangent_ms=" & Format$(P6ProfTangentMs, "0") & vbCrLf
  lineText = lineText & "TimeAssembly_ms=" & Format$(P6ProfAssembleMs, "0") & vbCrLf
  lineText = lineText & "TimeEval_ms=" & Format$(P6ProfEvalMs, "0") & vbCrLf
  lineText = lineText & "SolverMode=" & P6SolverMode & vbCrLf
  lineText = lineText & "BandLUKernel=" & P6BandKernelName() & vbCrLf
  If P6FactorizationCount > 0 Then
    lineText = lineText & "FactorAvgMs=" & Format$(P6ProfFactorMs / CDbl(P6FactorizationCount), "0.000") & vbCrLf
  Else
    lineText = lineText & "FactorAvgMs=NA" & vbCrLf
  End If
  lineText = lineText & "LDLTShadowCount=" & CStr(P6LdltN) & vbCrLf
  lineText = lineText & "LDLTOkCount=" & CStr(P6LdltOk) & vbCrLf
  lineText = lineText & "LDLTFailCount=" & CStr(P6LdltFail) & vbCrLf
  lineText = lineText & "LDLTFactorMs=" & Format$(P6LdltMs, "0.000") & vbCrLf
  lineText = lineText & "LimitStateMonitor=SRM_SHADOW" & vbCrLf
  lineText = lineText & "LimitStateNote=per-Fs mechanical_status is in perf_summary.csv" & vbCrLf
  On Error Resume Next
  If femIoFileNum <> 0 Then
    Print #femIoFileNum, "PERF,SUMMARY," & FEMIoCsv(Replace$(lineText, vbCrLf, " | "))
  End If
  folderPath = FEMIoOutDir()
  If Len(folderPath) > 0 Then
    summaryPath = folderPath & Application.PathSeparator & "perf_summary.txt"
    fileNum = FreeFile
    Open summaryPath For Output As #fileNum
    Print #fileNum, "========== PERFORMANCE SUMMARY =========="
    Print #fileNum, "Ver=" & FEM_BUILD_STAMP
    Print #fileNum, lineText
    Print #fileNum, "========================================="
    Close #fileNum
  End If
  Err.Clear
End Sub

Private Function FEMIoOutDir() As String
  Dim folderPath As String, baseName As String, fileName As String
  FEMIoOutDir = vbNullString
  fileName = ThisWorkbook.name
  If InStrRev(fileName, ".") > 1 Then
    baseName = left$(fileName, InStrRev(fileName, ".") - 1)
  Else
    baseName = fileName
  End If
  If Len(ThisWorkbook.path) = 0 Then
    folderPath = CurDir$() & Application.PathSeparator & baseName & "_out"
  Else
    folderPath = ThisWorkbook.path & Application.PathSeparator & baseName & "_out"
  End If
  On Error Resume Next
  If Dir$(folderPath, vbDirectory) = vbNullString Then MkDir folderPath
  If Dir$(folderPath, vbDirectory) = vbNullString Then Exit Function
  On Error GoTo 0
  FEMIoOutDir = folderPath
End Function

Private Function FEMIoModelSignature() As String
  FEMIoModelSignature = FEMIoContentHash()
End Function

Private Function FEMIoMix(ByVal currentValue As Double, ByVal nextValue As Double) As Double
  FEMIoMix = P6FingerprintStep(currentValue, nextValue)
End Function

Private Function FEMIoMixText(ByVal currentValue As Double, ByVal textValue As String) As Double
  Dim i As Long, h As Double
  h = P6FingerprintStep(currentValue, CDbl(Len(textValue)))
  For i = 1 To Len(textValue)
    h = P6FingerprintStep(h, CDbl(AscW(mid$(textValue, i, 1))))
  Next i
  FEMIoMixText = h
End Function

Private Function FEMIoMixFlag(ByVal currentValue As Double, ByVal flagValue As Boolean) As Double
  If flagValue Then
    FEMIoMixFlag = P6FingerprintStep(currentValue, 2#)
  Else
    FEMIoMixFlag = P6FingerprintStep(currentValue, 1#)
  End If
End Function

Private Function FEMIoContentHash() As String
  Dim h As Double, i As Long, j As Long
  h = 17#
  h = FEMIoMixText(h, FEM_BUILD_STAMP)
  h = FEMIoMixText(h, UCase$(Trim$(P6ReadTextSetting("RCM_POLICY", "AUTO"))))
  h = FEMIoMixText(h, UCase$(Trim$(P6ReadTextSetting("FLOW_POLICY", "INCONSISTENT"))))
  h = FEMIoMixText(h, UCase$(Trim$(P6ReadTextSetting("WEIGHT_MODE", "TOTAL"))))
  h = FEMIoMix(h, P6QuantizeScaled(P6ReadSetting("Q8_HOURGLASS_FACTOR", 0.05), 1000000#))
  If P6InputCacheReady Then
    h = FEMIoMix(h, CDbl(P6InputNodeCount))
    h = FEMIoMix(h, CDbl(P6InputElementCount))
    h = FEMIoMix(h, CDbl(P6InputMaterialCount))
    For i = 1 To P6InputNodeCount
      h = FEMIoMix(h, CDbl(P6InputNodeId(i)))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputNodeX(i), 1000000#))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputNodeY(i), 1000000#))
      h = FEMIoMix(h, CDbl(P6InputNodeCondX(i)))
      h = FEMIoMix(h, CDbl(P6InputNodeCondY(i)))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputNodeDispX(i), 1000000#))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputNodeDispY(i), 1000000#))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputNodeForceX(i), 1000000#))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputNodeForceY(i), 1000000#))
    Next i
    For i = 1 To P6InputMaterialCount
      h = FEMIoMix(h, CDbl(P6InputMaterialId(i)))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputMaterialYoung(i), 1000000#))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputMaterialPoisson(i), 1000000000#))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputMaterialThickness(i), 1000000#))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputMaterialWeight(i), 1000000#))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputMaterialFriction(i), 1000000#))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputMaterialCohesion(i), 1000000#))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputMaterialDilation(i), 1000000#))
      h = FEMIoMix(h, P6QuantizeScaled(P6InputMaterialKs(i), 1000000#))
      h = FEMIoMixFlag(h, P6InputMaterialReduce(i))
      h = FEMIoMixFlag(h, P6InputMaterialInitialActive(i))
      h = FEMIoMixFlag(h, P6InputMaterialAllowTension(i))
      h = FEMIoMixText(h, UCase$(Trim$(P6InputMaterialKind(i))))
    Next i
    For i = 1 To P6InputElementCount
      h = FEMIoMix(h, CDbl(P6InputElementId(i)))
      h = FEMIoMix(h, CDbl(P6InputElementMaterial(i)))
      For j = 1 To 8
        h = FEMIoMix(h, CDbl(P6InputElementNode(i, j)))
      Next j
    Next i
  Else
    h = FEMIoMix(h, CDbl(NumberOfNode))
    h = FEMIoMix(h, CDbl(NumberOfElement))
    h = FEMIoMix(h, CDbl(NumberOfMaterial))
  End If
  h = FEMIoMixText(h, FEMIoStagePlanHash())
  h = FEMIoMixText(h, FEMIoLoadingHash())
  FEMIoContentHash = Format$(h, "0")
End Function

Private Function FEMIoStagePlanHash() As String
  Dim h As Double, i As Long
  h = 19#
  h = FEMIoMix(h, CDbl(P3StageN))
  If P3StageN >= 1 Then
    For i = 1 To P3StageN
      h = FEMIoMix(h, CDbl(P3StageId(i)))
      h = FEMIoMixText(h, P3StageKind(i))
      h = FEMIoMixText(h, P3StageGroup(i))
      h = FEMIoMixText(h, P3StageParam(i))
      h = FEMIoMixFlag(h, P3StageOn(i))
    Next i
  End If
  FEMIoStagePlanHash = Format$(h, "0")
End Function

Private Function FEMIoLoadingHash() As String
  Dim ws As Worksheet, lastRow As Long, rowNo As Long, colNo As Long, h As Double
  h = 23#
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("載荷")
  On Error GoTo 0
  If ws Is Nothing Then
    FEMIoLoadingHash = Format$(h, "0")
    Exit Function
  End If
  lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
  h = FEMIoMix(h, CDbl(lastRow))
  For rowNo = 2 To lastRow
    For colNo = 1 To 9
      h = FEMIoMixText(h, Trim$(CStr(ws.Cells(rowNo, colNo).value2)))
    Next colNo
  Next rowNo
  FEMIoLoadingHash = Format$(h, "0")
End Function

Private Function FEMIoParseLong(ByVal textValue As String, ByRef value As Long) As Boolean
  Dim t As String, d As Double
  FEMIoParseLong = False
  value = 0
  t = Trim$(textValue)
  If Len(t) = 0 Then Exit Function
  If Not IsNumeric(t) Then Exit Function
  On Error GoTo ParseFail
  d = CDbl(t)
  If d <> Fix(d) Then Exit Function
  If d < -2147483648# Or d > 2147483647# Then Exit Function
  value = CLng(d)
  FEMIoParseLong = True
  Exit Function
ParseFail:
End Function

Private Function FEMIoParseDouble(ByVal textValue As String, ByRef value As Double) As Boolean
  Dim t As String
  FEMIoParseDouble = False
  value = 0#
  t = Trim$(textValue)
  If Len(t) = 0 Then Exit Function
  If Not IsNumeric(t) Then Exit Function
  On Error GoTo ParseFail
  value = CDbl(t)
  If Not P2IsFinite(value) Then
    value = 0#
    Exit Function
  End If
  FEMIoParseDouble = True
  Exit Function
ParseFail:
End Function

Private Function FEMIoFolderExists(ByVal folderPath As String) As Boolean
  On Error Resume Next
  FEMIoFolderExists = False
  If Len(folderPath) = 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  FEMIoFolderExists = (Dir$(folderPath, vbDirectory) <> vbNullString)
  Err.Clear
  On Error GoTo 0
End Function

Private Sub FEMIoDeleteFolder(ByVal folderPath As String)
  Dim fn As String, sep As String
  If Len(folderPath) = 0 Then Exit Sub
  sep = Application.PathSeparator
  On Error Resume Next
  If Dir$(folderPath, vbDirectory) = vbNullString Then
    Err.Clear
    On Error GoTo 0
    Exit Sub
  End If
  fn = Dir$(folderPath & sep & "*.*")
  Do While Len(fn) > 0
    Kill folderPath & sep & fn
    fn = Dir$
  Loop
  RmDir folderPath
  Err.Clear
  On Error GoTo 0
End Sub

Private Function FEMIoOriginalFreeNode(ByVal internalNode As Long) As Long
  FEMIoOriginalFreeNode = P6GetOriginalFreeNode(internalNode)
End Function

Private Function FEMIoInternalFreeNode(ByVal originalFreeNode As Long) As Long
  FEMIoInternalFreeNode = P6GetInternalFreeNode(originalFreeNode)
End Function

' Write a stage folder under *_out. SUCCESS also stores node/gauss/dof and origin.csv when a replay origin exists.
Public Sub FEMIoExportStage(ByVal stageNo As Long, ByVal kindText As String, ByVal okFlag As Boolean)
  Dim folderPath As String, destFolder As String, tmpFolder As String, modeText As String
  Dim fullState As Boolean
  If P3ReplayQuiet Then Exit Sub
  modeText = femIoExportMode
  If Len(modeText) = 0 Then modeText = UCase$(Trim$(FEMReadTextSetting("EXPORT_MODE", "STAGE")))
  If modeText = "OFF" Then Exit Sub
  fullState = False
  If okFlag Then
    If modeText <> "STAGE" And modeText <> "ON_FAIL" Then Exit Sub
    fullState = True
  Else
    If modeText <> "STAGE" And modeText <> "ON_FAIL" Then Exit Sub
    fullState = (modeText = "ON_FAIL")
  End If
  folderPath = FEMIoOutDir()
  If Len(folderPath) = 0 Then Exit Sub
  destFolder = folderPath & Application.PathSeparator & "stg" & Format$(stageNo, "00") & "_" & FEMIoSafeName(kindText)
  tmpFolder = destFolder & ".tmp"
  FEMIoDeleteFolder tmpFolder
  On Error GoTo ExportFail
  MkDir tmpFolder
  FEMIoWriteMeta tmpFolder, stageNo, kindText, okFlag
  If fullState Then
    If Not FEMIoWriteNode(tmpFolder) Then GoTo ExportFail
    If Not FEMIoWriteGauss(tmpFolder) Then GoTo ExportFail
    If Not PlasticLogWrite(tmpFolder) Then FEMIoLog "累積塑性診断", "WARN", "plastic_cumulative.csvを保存できませんでした。"
    If Not FEMIoWriteDofLock(tmpFolder) Then GoTo ExportFail
    If Not FEMIoWriteOrigin(tmpFolder) Then GoTo ExportFail
  End If
  FEMIoDeleteFolder destFolder
  Name tmpFolder As destFolder
  FEMIoAppendManifest destFolder, stageNo, kindText, okFlag
  On Error GoTo 0
  Exit Sub
ExportFail:
  On Error Resume Next
  FEMIoDeleteFolder tmpFolder
  Err.Clear
  On Error GoTo 0
End Sub

Private Sub FEMIoWriteMeta(ByVal stageFolder As String, ByVal stageNo As Long, ByVal kindText As String, ByVal okFlag As Boolean)
  Dim fileNumber As Integer
  fileNumber = FreeFile
  Open stageFolder & Application.PathSeparator & "meta.txt" For Output As #fileNumber
  Print #fileNumber, "STAMP=" & FEM_BUILD_STAMP
  Print #fileNumber, "STAGE=" & CStr(stageNo)
  Print #fileNumber, "KIND=" & kindText
  Print #fileNumber, "STATUS=" & IIf(okFlag, "SUCCESS", "FAIL")
  Print #fileNumber, "SIG=" & FEMIoModelSignature()
  Print #fileNumber, "PLAN=" & FEMIoStagePlanHash()
  Print #fileNumber, "RCM=" & UCase$(Trim$(P6ReadTextSetting("RCM_POLICY", "AUTO")))
  Print #fileNumber, "FS=" & Format$(P3CurrentStrengthFactor, "0.000000000000")
  Print #fileNumber, "FSS=" & Format$(FSS, "0.000000000000")
  Print #fileNumber, "GRAVITY=" & IIf(P3GravityCommitted, "1", "0")
  Print #fileNumber, "PREFIX=" & CStr(P3ActivePrefix)
  Print #fileNumber, "LASTINC=" & CStr(P3SuccessfulIncrementCount)
  Print #fileNumber, "NODES=" & CStr(NumberOfNode)
  Print #fileNumber, "ELEMS=" & CStr(NumberOfElement)
  Print #fileNumber, "FREE=" & CStr(NumberOfFreeNode)
  Print #fileNumber, "DOF=" & CStr(lastDof)
  Print #fileNumber, "ORIGIN=" & IIf(P3OriginSnapReady, "1", "0")
  Print #fileNumber, "REPLAY_START=" & CStr(P3SrmReplayStart)
  Close #fileNumber
End Sub

Private Function FEMIoWriteNode(ByVal stageFolder As String) As Boolean
  Dim fileNumber As Integer, origFree As Long, internalNode As Long, dof0 As Long
  FEMIoWriteNode = False
  If NumberOfFreeNode <= 0 Then Exit Function
  fileNumber = FreeFile
  Open stageFolder & Application.PathSeparator & "node.csv" For Output As #fileNumber
  Print #fileNumber, "node,x,y,ux,uy,condx,condy"
  For origFree = 0 To NumberOfFreeNode - 1
    internalNode = FEMIoInternalFreeNode(origFree)
    dof0 = 2 * internalNode
    Print #fileNumber, CStr(origFree + 1) & "," & Format$(XXX(dof0), "0.0000000000") & "," & Format$(XXX(dof0 + 1), "0.0000000000") & "," & Format$(TDisp(dof0), "0.000000000000") & "," & Format$(TDisp(dof0 + 1), "0.000000000000") & "," & CStr(NodeCond(dof0)) & "," & CStr(NodeCond(dof0 + 1))
  Next origFree
  Close #fileNumber
  FEMIoWriteNode = True
End Function

Private Function FEMIoWriteGauss(ByVal stageFolder As String) As Boolean
  Dim fileNumber As Integer, k As Long, im As Long
  FEMIoWriteGauss = False
  fileNumber = FreeFile
  Open stageFolder & Application.PathSeparator & "gauss.csv" For Output As #fileNumber
  Print #fileNumber, "elem,gp,sx,sy,txy,sz,ex_pl,ey_pl,gxy_pl,ez_pl,mult,f,yielded"
  For k = 0 To NumberOfElement - 1
    For im = 0 To 3
      Print #fileNumber, CStr(k + 1) & "," & CStr(im) & "," & Format$(Elem(k).Stmat(0, im), "0.000000000000") & "," & Format$(Elem(k).Stmat(1, im), "0.000000000000") & "," & Format$(Elem(k).Stmat(2, im), "0.000000000000") & "," & Format$(Elem(k).Stmat(16, im), "0.000000000000") & "," & Format$(Elem(k).P2PlasticStrain(0, im), "0.000000000000") & "," & Format$(Elem(k).P2PlasticStrain(1, im), "0.000000000000") & "," & Format$(Elem(k).P2PlasticStrain(2, im), "0.000000000000") & "," & Format$(Elem(k).P2PlasticStrain(3, im), "0.000000000000") & "," & Format$(Elem(k).P2PlasticMultiplier(im), "0.000000000000") & "," & Format$(Elem(k).P2YieldFunction(im), "0.000000000000") & "," & IIf(Elem(k).P2Yielded(im), "1", "0")
    Next im
  Next k
  Close #fileNumber
  FEMIoWriteGauss = True
End Function

Private Function FEMIoWriteDofLock(ByVal stageFolder As String) As Boolean
  Dim fileNumber As Integer, origFree As Long, localDof As Long, origDof As Long, intDof As Long
  Dim swValue As Double, apValue As Double, bdValue As Double, condValue As Long
  Dim swBound As Long, apBound As Long, bdBound As Long
  FEMIoWriteDofLock = False
  If lastDof < 0 Then Exit Function
  swBound = -1
  apBound = -1
  bdBound = -1
  On Error Resume Next
  swBound = UBound(P3LockedSelfWeight)
  apBound = UBound(P3LockedApplied)
  bdBound = UBound(P3BoundaryDisp)
  Err.Clear
  On Error GoTo 0
  If swBound <> lastDof Or apBound <> lastDof Or bdBound <> lastDof Then Exit Function
  fileNumber = FreeFile
  Open stageFolder & Application.PathSeparator & "dof.csv" For Output As #fileNumber
  Print #fileNumber, "dof,locked_sw,locked_ap,bound_disp,cond"
  For origFree = 0 To NumberOfFreeNode - 1
    For localDof = 0 To 1
      origDof = 2 * origFree + localDof
      intDof = 2 * FEMIoInternalFreeNode(origFree) + localDof
      If intDof < 0 Or intDof > lastDof Then
        Close #fileNumber
        Exit Function
      End If
      swValue = P3LockedSelfWeight(intDof)
      apValue = P3LockedApplied(intDof)
      bdValue = P3BoundaryDisp(intDof)
      condValue = NodeCond(intDof)
      Print #fileNumber, CStr(origDof) & "," & Format$(swValue, "0.000000000000") & "," & Format$(apValue, "0.000000000000") & "," & Format$(bdValue, "0.000000000000") & "," & CStr(condValue)
    Next localDof
  Next origFree
  Close #fileNumber
  FEMIoWriteDofLock = True
End Function

' Persist the RESET/MATSET origin so a later SRM can replay after Excel restart.
Private Function FEMIoWriteOrigin(ByVal stageFolder As String) As Boolean
  Dim fileNumber As Integer, i As Long, origFree As Long, localDof As Long, origDof As Long, intDof As Long
  FEMIoWriteOrigin = True
  If Not P3OriginSnapReady Then Exit Function
  If NumberOfElement <= 0 Or lastDof < 0 Then Exit Function
  If UBound(P3OriginElementActive) <> NumberOfElement - 1 Then
    FEMIoWriteOrigin = False
    Exit Function
  End If
  If UBound(P3OriginNodeCond) <> lastDof Then
    FEMIoWriteOrigin = False
    Exit Function
  End If
  fileNumber = FreeFile
  Open stageFolder & Application.PathSeparator & "origin.csv" For Output As #fileNumber
  Print #fileNumber, "kind,id,v1,v2,v3,v4,v5"
  Print #fileNumber, "KEY,replay_start," & CStr(P3SrmReplayStart)
  Print #fileNumber, "KEY,has_gravity," & IIf(P3OriginHasGravity, "1", "0")
  Print #fileNumber, "KEY,has_apply," & IIf(P3OriginHasApply, "1", "0")
  Print #fileNumber, "KEY,elems," & CStr(NumberOfElement)
  Print #fileNumber, "KEY,dof," & CStr(lastDof)
  For i = 0 To NumberOfElement - 1
    Print #fileNumber, "ELEM," & CStr(i + 1) & "," & IIf(P3OriginElementActive(i), "1", "0")
  Next i
  For origFree = 0 To NumberOfFreeNode - 1
    For localDof = 0 To 1
      origDof = 2 * origFree + localDof
      intDof = 2 * FEMIoInternalFreeNode(origFree) + localDof
      If intDof < 0 Or intDof > lastDof Then
        Close #fileNumber
        FEMIoWriteOrigin = False
        Exit Function
      End If
      Print #fileNumber, "DOF," & CStr(origDof) & "," & CStr(P3OriginNodeCond(intDof)) & "," & Format$(P3OriginBoundaryDisp(intDof), "0.000000000000") & "," & Format$(P3OriginAppliedForce(intDof), "0.000000000000") & "," & Format$(P3OriginLockedSW(intDof), "0.000000000000") & "," & Format$(P3OriginLockedAP(intDof), "0.000000000000")
    Next localDof
  Next origFree
  Close #fileNumber
End Function

' Missing origin.csv is allowed (returns True, OriginSnapReady=False). Corrupt file fails resume.
Private Function FEMIoLoadOrigin(ByVal stageFolder As String) As Boolean
  Dim fileNumber As Integer, lineText As String, parts() As String
  Dim kindText As String, keyName As String, rhs As String
  Dim elemId As Long, origDof As Long, origFree As Long, localDof As Long, intDof As Long
  Dim flagValue As Long, condValue As Long
  Dim boundValue As Double, appliedValue As Double, swValue As Double, apValue As Double
  Dim nElem As Long, nDof As Long, replayStart As Long
  Dim hasGravity As Boolean, hasApply As Boolean
  Dim seenElem() As Boolean, seenDof() As Boolean
  Dim loadActive() As Boolean
  Dim loadCond() As Long
  Dim loadBound() As Double, loadApplied() As Double, loadSw() As Double, loadAp() As Double
  Dim i As Long
  FEMIoLoadOrigin = False
  P3OriginSnapReady = False
  If Dir$(stageFolder & Application.PathSeparator & "origin.csv") = vbNullString Then
    FEMIoLoadOrigin = True
    Exit Function
  End If
  If NumberOfElement <= 0 Or lastDof < 0 Then Exit Function
  ReDim seenElem(0 To NumberOfElement - 1)
  ReDim loadActive(0 To NumberOfElement - 1)
  ReDim seenDof(0 To lastDof)
  ReDim loadCond(0 To lastDof)
  ReDim loadBound(0 To lastDof)
  ReDim loadApplied(0 To lastDof)
  ReDim loadSw(0 To lastDof)
  ReDim loadAp(0 To lastDof)
  nElem = -1
  nDof = -1
  replayStart = 0
  fileNumber = FreeFile
  Open stageFolder & Application.PathSeparator & "origin.csv" For Input As #fileNumber
  If Not EOF(fileNumber) Then Line Input #fileNumber, lineText
  Do While Not EOF(fileNumber)
    Line Input #fileNumber, lineText
    If Len(Trim$(lineText)) = 0 Then GoTo NextOriginLine
    parts = Split(lineText, ",")
    If UBound(parts) < 2 Then
      Close #fileNumber
      Exit Function
    End If
    kindText = UCase$(Trim$(parts(0)))
    If kindText = "KEY" Then
      keyName = LCase$(Trim$(parts(1)))
      rhs = parts(2)
      If keyName = "replay_start" Then
        If Not FEMIoParseLong(rhs, replayStart) Then
          Close #fileNumber
          Exit Function
        End If
      ElseIf keyName = "has_gravity" Then
        hasGravity = (Trim$(rhs) = "1")
      ElseIf keyName = "has_apply" Then
        hasApply = (Trim$(rhs) = "1")
      ElseIf keyName = "elems" Then
        If Not FEMIoParseLong(rhs, nElem) Then
          Close #fileNumber
          Exit Function
        End If
      ElseIf keyName = "dof" Then
        If Not FEMIoParseLong(rhs, nDof) Then
          Close #fileNumber
          Exit Function
        End If
      End If
    ElseIf kindText = "ELEM" Then
      If Not FEMIoParseLong(parts(1), elemId) Then
        Close #fileNumber
        Exit Function
      End If
      If elemId < 1 Or elemId > NumberOfElement Then
        Close #fileNumber
        Exit Function
      End If
      If seenElem(elemId - 1) Then
        Close #fileNumber
        Exit Function
      End If
      If Not FEMIoParseLong(parts(2), flagValue) Then
        Close #fileNumber
        Exit Function
      End If
      seenElem(elemId - 1) = True
      loadActive(elemId - 1) = (flagValue <> 0)
    ElseIf kindText = "DOF" Then
      If UBound(parts) < 6 Then
        Close #fileNumber
        Exit Function
      End If
      If Not FEMIoParseLong(parts(1), origDof) Then
        Close #fileNumber
        Exit Function
      End If
      If origDof < 0 Or origDof > lastDof Then
        Close #fileNumber
        Exit Function
      End If
      If seenDof(origDof) Then
        Close #fileNumber
        Exit Function
      End If
      If Not FEMIoParseLong(parts(2), condValue) Then
        Close #fileNumber
        Exit Function
      End If
      If Not FEMIoParseDouble(parts(3), boundValue) Then
        Close #fileNumber
        Exit Function
      End If
      If Not FEMIoParseDouble(parts(4), appliedValue) Then
        Close #fileNumber
        Exit Function
      End If
      If Not FEMIoParseDouble(parts(5), swValue) Then
        Close #fileNumber
        Exit Function
      End If
      If Not FEMIoParseDouble(parts(6), apValue) Then
        Close #fileNumber
        Exit Function
      End If
      seenDof(origDof) = True
      origFree = origDof \ 2
      localDof = origDof Mod 2
      intDof = 2 * FEMIoInternalFreeNode(origFree) + localDof
      If intDof < 0 Or intDof > lastDof Then
        Close #fileNumber
        Exit Function
      End If
      loadCond(intDof) = condValue
      loadBound(intDof) = boundValue
      loadApplied(intDof) = appliedValue
      loadSw(intDof) = swValue
      loadAp(intDof) = apValue
    End If
NextOriginLine:
  Loop
  Close #fileNumber
  If nElem >= 0 And nElem <> NumberOfElement Then Exit Function
  If nDof >= 0 And nDof <> lastDof Then Exit Function
  For i = 0 To NumberOfElement - 1
    If Not seenElem(i) Then Exit Function
  Next i
  For i = 0 To lastDof
    If Not seenDof(i) Then Exit Function
  Next i
  ReDim P3OriginElementActive(0 To NumberOfElement - 1)
  ReDim P3OriginNodeCond(lastDof)
  ReDim P3OriginBoundaryDisp(lastDof)
  ReDim P3OriginAppliedForce(lastDof)
  ReDim P3OriginLockedSW(lastDof)
  ReDim P3OriginLockedAP(lastDof)
  For i = 0 To NumberOfElement - 1
    P3OriginElementActive(i) = loadActive(i)
  Next i
  For i = 0 To lastDof
    P3OriginNodeCond(i) = loadCond(i)
    P3OriginBoundaryDisp(i) = loadBound(i)
    P3OriginAppliedForce(i) = loadApplied(i)
    P3OriginLockedSW(i) = loadSw(i)
    P3OriginLockedAP(i) = loadAp(i)
  Next i
  P3OriginHasGravity = hasGravity
  P3OriginHasApply = hasApply
  If replayStart > 0 Then P3SrmReplayStart = replayStart
  P3OriginSnapReady = True
  FEMIoLoadOrigin = True
End Function

' After XXX is restored, rebuild element coords, B/S/k, then SetTotalMat. Always rebuild, never reuse cache.
Private Function FEMIoRebuildResumeGeometry() As Boolean
  Dim k As Long, j As Long, nd As Long
  FEMIoRebuildResumeGeometry = False
  For k = 0 To NumberOfElement - 1
    For j = 0 To 7
      nd = Elem(k).node(j)
      If nd >= 0 Then
        Elem(k).x(j) = XXX(2 * nd)
        Elem(k).y(j) = XXX(2 * nd + 1)
      End If
    Next j
  Next k
  P6InvalidateTangentGeneration
  If Not SetElmMat() Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "再開形状から要素行列を再構築できませんでした。", vbObjectError + 3099, -1, -1, 0, 0
    Exit Function
  End If
  SetTotalMat
  FEMIoRebuildResumeGeometry = True
End Function

Private Sub FEMIoAppendManifest(ByVal stageFolder As String, ByVal stageNo As Long, ByVal kindText As String, ByVal okFlag As Boolean)
  Dim fileNumber As Integer, folderPath As String, leaf As String, p As Long
  folderPath = FEMIoOutDir()
  If Len(folderPath) = 0 Then Exit Sub
  p = InStrRev(stageFolder, Application.PathSeparator)
  If p > 0 Then leaf = mid$(stageFolder, p + 1) Else leaf = stageFolder
  fileNumber = FreeFile
  Open folderPath & Application.PathSeparator & "manifest.csv" For Append As #fileNumber
  Print #fileNumber, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "," & leaf & "," & CStr(stageNo) & "," & FEMIoCsv(kindText) & "," & IIf(okFlag, "SUCCESS", "FAIL") & "," & FEMIoModelSignature()
  Close #fileNumber
End Sub

Private Function FEMIoSafeName(ByVal textValue As String) As String
  Dim s As String
  s = UCase$(Trim$(textValue))
  If Len(s) = 0 Then s = "STAGE"
  s = Replace$(s, " ", "_")
  s = Replace$(s, "/", "_")
  s = Replace$(s, "\", "_")
  FEMIoSafeName = left$(s, 24)
End Function

' Load EXPORT_LOAD checkpoint. Returns resumed stage number, or 0 to start from the beginning.
Public Function FEMIoTryResume() As Long
  Dim want As String, folderPath As String, stageFolder As String, fileNumber As Integer
  Dim lineText As String, parts() As String, lastOk As String
  Dim metaStamp As String, metaSig As String, metaPlan As String, metaStatus As String
  Dim metaStage As Long, metaKind As String, metaNodes As Long, metaElems As Long, metaDof As Long
  Dim rhs As String, matchCount As Long, stageIndex As Long, i As Long
  FEMIoTryResume = 0
  femIoResumeIndex = 0
  femIoResumeFs = 1#
  femIoResumeFss = 1#
  femIoResumeGravity = False
  femIoResumeHasLock = False
  femIoResumePrefix = 0
  femIoResumeKind = vbNullString
  want = UCase$(Trim$(FEMReadTextSetting("EXPORT_LOAD", "")))
  If Len(want) = 0 Then Exit Function
  folderPath = FEMIoOutDir()
  If Len(folderPath) = 0 Then Exit Function
  If want = "AUTO" Then
    lastOk = vbNullString
    On Error Resume Next
    fileNumber = FreeFile
    Open folderPath & Application.PathSeparator & "manifest.csv" For Input As #fileNumber
    If Err.Number = 0 Then
      Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineText
        parts = Split(lineText, ",")
        If UBound(parts) >= 4 Then
          If UCase$(Trim$(parts(4))) = "SUCCESS" Then lastOk = Trim$(parts(1))
        End If
      Loop
    End If
    Close #fileNumber
    Err.Clear
    On Error GoTo 0
    If Len(lastOk) = 0 Then Exit Function
    stageFolder = folderPath & Application.PathSeparator & lastOk
  Else
    stageFolder = folderPath & Application.PathSeparator & want
  End If
  If Dir$(stageFolder & Application.PathSeparator & "meta.txt") = vbNullString Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "再開フォルダが見つかりません: " & stageFolder, vbObjectError + 3091, -1, -1, 0, 0
    Exit Function
  End If
  metaStage = -1
  metaNodes = -1
  metaElems = -1
  metaDof = -1
  fileNumber = FreeFile
  Open stageFolder & Application.PathSeparator & "meta.txt" For Input As #fileNumber
  Do While Not EOF(fileNumber)
    Line Input #fileNumber, lineText
    If left$(lineText, 6) = "STAMP=" Then metaStamp = mid$(lineText, 7)
    If left$(lineText, 4) = "SIG=" Then metaSig = mid$(lineText, 5)
    If left$(lineText, 5) = "PLAN=" Then metaPlan = mid$(lineText, 6)
    If left$(lineText, 7) = "STATUS=" Then metaStatus = UCase$(Trim$(mid$(lineText, 8)))
    If left$(lineText, 6) = "STAGE=" Then
      rhs = mid$(lineText, 7)
      If Not FEMIoParseLong(rhs, metaStage) Then metaStage = -1
    End If
    If left$(lineText, 5) = "KIND=" Then metaKind = mid$(lineText, 6)
    If left$(lineText, 3) = "FS=" Then
      rhs = mid$(lineText, 4)
      If IsNumeric(rhs) Then femIoResumeFs = CDbl(rhs)
    End If
    If left$(lineText, 4) = "FSS=" Then
      rhs = mid$(lineText, 5)
      If IsNumeric(rhs) Then femIoResumeFss = CDbl(rhs)
    End If
    If left$(lineText, 8) = "GRAVITY=" Then
      rhs = Trim$(mid$(lineText, 9))
      femIoResumeGravity = (rhs = "1" Or UCase$(rhs) = "TRUE")
    End If
    If left$(lineText, 7) = "PREFIX=" Then
      rhs = mid$(lineText, 8)
      If IsNumeric(rhs) Then femIoResumePrefix = CLng(CDbl(rhs))
    End If
    If left$(lineText, 6) = "NODES=" Then
      rhs = mid$(lineText, 7)
      If Not FEMIoParseLong(rhs, metaNodes) Then metaNodes = -1
    End If
    If left$(lineText, 6) = "ELEMS=" Then
      rhs = mid$(lineText, 7)
      If Not FEMIoParseLong(rhs, metaElems) Then metaElems = -1
    End If
    If left$(lineText, 4) = "DOF=" Then
      rhs = mid$(lineText, 5)
      If Not FEMIoParseLong(rhs, metaDof) Then metaDof = -1
    End If
  Loop
  Close #fileNumber
  If metaStatus <> "SUCCESS" Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "FAIL状態のチェックポイントからは再開できません。", vbObjectError + 3093, -1, -1, 0, 0
    Exit Function
  End If
  If metaStamp <> FEM_BUILD_STAMP Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "外部出力のビルドが現在のブックと一致しません。再計算してください。", vbObjectError + 3092, -1, -1, 0, 0
    Exit Function
  End If
  If metaSig <> FEMIoModelSignature() Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "外部出力のモデル署名が現在のメッシュ・拘束・材料・ステージ・載荷と一致しません。再計算してください。", vbObjectError + 3092, -1, -1, 0, 0
    Exit Function
  End If
  If Len(metaPlan) > 0 Then
    If metaPlan <> FEMIoStagePlanHash() Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "外部出力のステージ計画が現在のステージシートと一致しません。再計算してください。", vbObjectError + 3094, -1, -1, 0, 0
      Exit Function
    End If
  End If
  If metaNodes >= 0 And metaNodes <> NumberOfNode Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "チェックポイントの節点数が一致しません。", vbObjectError + 3092, -1, -1, 0, 0
    Exit Function
  End If
  If metaElems >= 0 And metaElems <> NumberOfElement Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "チェックポイントの要素数が一致しません。", vbObjectError + 3092, -1, -1, 0, 0
    Exit Function
  End If
  If metaDof >= 0 And metaDof <> lastDof Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "チェックポイントの自由度数が一致しません。", vbObjectError + 3092, -1, -1, 0, 0
    Exit Function
  End If
  If metaStage <= 0 Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "チェックポイントのステージ番号が不正です。", vbObjectError + 3095, -1, -1, 0, 0
    Exit Function
  End If
  matchCount = 0
  For stageIndex = 1 To P3StageN
    If P3StageId(stageIndex) = metaStage Then
      If Len(metaKind) > 0 Then
        If P3StageKind(stageIndex) <> metaKind Then
          SetAnalysisFailure RESULT_INPUT_ERROR, "チェックポイントのステージ種別が一致しません。保存=" & metaKind & " 現在=" & P3StageKind(stageIndex), vbObjectError + 3096, -1, -1, 0, 0
          Exit Function
        End If
      End If
      matchCount = matchCount + 1
      femIoResumeIndex = stageIndex
    End If
  Next stageIndex
  If matchCount = 0 Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "チェックポイントのステージ番号が現在のステージ計画にありません。STAGE=" & CStr(metaStage), vbObjectError + 3095, -1, -1, 0, 0
    Exit Function
  End If
  If matchCount > 1 Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "ステージ番号が重複しているため再開できません。STAGE=" & CStr(metaStage), vbObjectError + 3095, -1, -1, 0, 0
    Exit Function
  End If
  If Not FEMIoLoadNode(stageFolder) Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "node.csv が不完全です。欠損・重複・不正数値のため再開できません。", vbObjectError + 3097, -1, -1, 0, 0
    Exit Function
  End If
  If Not FEMIoLoadGauss(stageFolder) Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "gauss.csv が不完全です。欠損・重複・不正数値のため再開できません。", vbObjectError + 3098, -1, -1, 0, 0
    Exit Function
  End If
  If Dir$(stageFolder & Application.PathSeparator & "dof.csv") = vbNullString Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "dof.csv がありません。不完全なチェックポイントからは再開できません。", vbObjectError + 3099, -1, -1, 0, 0
    Exit Function
  End If
  If Not FEMIoLoadOrigin(stageFolder) Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "origin.csv が不完全です。RESET/MATSET originを復元できないため再開できません。", vbObjectError + 3100, -1, -1, 0, 0
    Exit Function
  End If
  P3CommitMaterialState
  PlasticLogResumeUnknown
  For i = 0 To lastDof
    P3CommittedDisp(i) = TDisp(i)
    UDisp(i) = TDisp(i)
  Next i
  femIoResumeFolder = stageFolder
  femIoResumeStage = metaStage
  femIoResumeKind = metaKind
  FEMIoTryResume = metaStage
  FEMIoLog "再開", "LOAD", stageFolder & " stage=" & CStr(metaStage) & " " & metaKind & " fs=" & Format$(femIoResumeFs, "0.000") & " gravity=" & IIf(femIoResumeGravity, "1", "0")
End Function

' Restore analysis coordinates XXX and displacements TDisp. x,y are required after RESET_STRESS.
Private Function FEMIoLoadNode(ByVal stageFolder As String) As Boolean
  Dim fileNumber As Integer, lineText As String, parts() As String
  Dim nodeId As Long, origFree As Long, internalNode As Long, nRead As Long
  Dim uxValue As Double, uyValue As Double, xValue As Double, yValue As Double
  Dim seen() As Boolean, loadUx() As Double, loadUy() As Double, loadX() As Double, loadY() As Double
  FEMIoLoadNode = False
  If NumberOfFreeNode <= 0 Then Exit Function
  If Dir$(stageFolder & Application.PathSeparator & "node.csv") = vbNullString Then Exit Function
  ReDim seen(0 To NumberOfFreeNode - 1)
  ReDim loadUx(0 To NumberOfFreeNode - 1)
  ReDim loadUy(0 To NumberOfFreeNode - 1)
  ReDim loadX(0 To NumberOfFreeNode - 1)
  ReDim loadY(0 To NumberOfFreeNode - 1)
  fileNumber = FreeFile
  Open stageFolder & Application.PathSeparator & "node.csv" For Input As #fileNumber
  If Not EOF(fileNumber) Then Line Input #fileNumber, lineText
  nRead = 0
  Do While Not EOF(fileNumber)
    Line Input #fileNumber, lineText
    If Len(Trim$(lineText)) = 0 Then GoTo NextNodeLine
    parts = Split(lineText, ",")
    If UBound(parts) < 4 Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseLong(parts(0), nodeId) Then
      Close #fileNumber
      Exit Function
    End If
    origFree = nodeId - 1
    If origFree < 0 Or origFree > NumberOfFreeNode - 1 Then
      Close #fileNumber
      Exit Function
    End If
    If seen(origFree) Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseDouble(parts(1), xValue) Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseDouble(parts(2), yValue) Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseDouble(parts(3), uxValue) Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseDouble(parts(4), uyValue) Then
      Close #fileNumber
      Exit Function
    End If
    seen(origFree) = True
    loadX(origFree) = xValue
    loadY(origFree) = yValue
    loadUx(origFree) = uxValue
    loadUy(origFree) = uyValue
    nRead = nRead + 1
NextNodeLine:
  Loop
  Close #fileNumber
  If nRead <> NumberOfFreeNode Then Exit Function
  For origFree = 0 To NumberOfFreeNode - 1
    If Not seen(origFree) Then Exit Function
    internalNode = FEMIoInternalFreeNode(origFree)
    XXX(2 * internalNode) = loadX(origFree)
    XXX(2 * internalNode + 1) = loadY(origFree)
    TDisp(2 * internalNode) = loadUx(origFree)
    TDisp(2 * internalNode + 1) = loadUy(origFree)
  Next origFree
  FEMIoLoadNode = True
End Function

Private Function FEMIoLoadGauss(ByVal stageFolder As String) As Boolean
  Dim fileNumber As Integer, lineText As String, parts() As String
  Dim k As Long, im As Long, elemId As Long, gpId As Long, nRead As Long, i As Long
  Dim values(0 To 9) As Double, yieldedFlag As Long
  Dim seen() As Boolean
  Dim loadSt0() As Double, loadSt1() As Double, loadSt2() As Double, loadSt16() As Double
  Dim loadPl0() As Double, loadPl1() As Double, loadPl2() As Double, loadPl3() As Double
  Dim loadMult() As Double, loadYieldFn() As Double, loadYielded() As Boolean
  FEMIoLoadGauss = False
  If NumberOfElement <= 0 Then Exit Function
  If Dir$(stageFolder & Application.PathSeparator & "gauss.csv") = vbNullString Then Exit Function
  ReDim seen(0 To NumberOfElement - 1, 0 To 3)
  ReDim loadSt0(0 To NumberOfElement - 1, 0 To 3)
  ReDim loadSt1(0 To NumberOfElement - 1, 0 To 3)
  ReDim loadSt2(0 To NumberOfElement - 1, 0 To 3)
  ReDim loadSt16(0 To NumberOfElement - 1, 0 To 3)
  ReDim loadPl0(0 To NumberOfElement - 1, 0 To 3)
  ReDim loadPl1(0 To NumberOfElement - 1, 0 To 3)
  ReDim loadPl2(0 To NumberOfElement - 1, 0 To 3)
  ReDim loadPl3(0 To NumberOfElement - 1, 0 To 3)
  ReDim loadMult(0 To NumberOfElement - 1, 0 To 3)
  ReDim loadYieldFn(0 To NumberOfElement - 1, 0 To 3)
  ReDim loadYielded(0 To NumberOfElement - 1, 0 To 3)
  fileNumber = FreeFile
  Open stageFolder & Application.PathSeparator & "gauss.csv" For Input As #fileNumber
  If Not EOF(fileNumber) Then Line Input #fileNumber, lineText
  nRead = 0
  Do While Not EOF(fileNumber)
    Line Input #fileNumber, lineText
    If Len(Trim$(lineText)) = 0 Then GoTo NextGaussLine
    parts = Split(lineText, ",")
    If UBound(parts) < 12 Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseLong(parts(0), elemId) Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseLong(parts(1), gpId) Then
      Close #fileNumber
      Exit Function
    End If
    k = elemId - 1
    im = gpId
    If k < 0 Or k > NumberOfElement - 1 Or im < 0 Or im > 3 Then
      Close #fileNumber
      Exit Function
    End If
    If seen(k, im) Then
      Close #fileNumber
      Exit Function
    End If
    For i = 0 To 9
      If Not FEMIoParseDouble(parts(i + 2), values(i)) Then
        Close #fileNumber
        Exit Function
      End If
    Next i
    If Not FEMIoParseLong(parts(12), yieldedFlag) Then
      Close #fileNumber
      Exit Function
    End If
    seen(k, im) = True
    loadSt0(k, im) = values(0)
    loadSt1(k, im) = values(1)
    loadSt2(k, im) = values(2)
    loadSt16(k, im) = values(3)
    loadPl0(k, im) = values(4)
    loadPl1(k, im) = values(5)
    loadPl2(k, im) = values(6)
    loadPl3(k, im) = values(7)
    loadMult(k, im) = values(8)
    loadYieldFn(k, im) = values(9)
    loadYielded(k, im) = (yieldedFlag <> 0)
    nRead = nRead + 1
NextGaussLine:
  Loop
  Close #fileNumber
  If nRead <> NumberOfElement * 4 Then Exit Function
  For k = 0 To NumberOfElement - 1
    For im = 0 To 3
      If Not seen(k, im) Then Exit Function
      Elem(k).Stmat(0, im) = loadSt0(k, im)
      Elem(k).Stmat(1, im) = loadSt1(k, im)
      Elem(k).Stmat(2, im) = loadSt2(k, im)
      Elem(k).Stmat(16, im) = loadSt16(k, im)
      Elem(k).P2PlasticStrain(0, im) = loadPl0(k, im)
      Elem(k).P2PlasticStrain(1, im) = loadPl1(k, im)
      Elem(k).P2PlasticStrain(2, im) = loadPl2(k, im)
      Elem(k).P2PlasticStrain(3, im) = loadPl3(k, im)
      Elem(k).P2PlasticMultiplier(im) = loadMult(k, im)
      Elem(k).P2YieldFunction(im) = loadYieldFn(k, im)
      Elem(k).P2Yielded(im) = loadYielded(k, im)
    Next im
  Next k
  FEMIoLoadGauss = True
End Function

Private Sub FEMIoExportDofLock(ByVal stageFolder As String)
  FEMIoWriteDofLock stageFolder
End Sub

Private Function FEMIoLoadDofLock(ByVal stageFolder As String) As Boolean
  Dim fileNumber As Integer, lineText As String, parts() As String
  Dim origDof As Long, origFree As Long, localDof As Long, intDof As Long, nRead As Long
  Dim swValue As Double, apValue As Double, bdValue As Double, condValue As Long
  Dim seen() As Boolean
  Dim loadSw() As Double, loadAp() As Double, loadBd() As Double, loadCond() As Long
  FEMIoLoadDofLock = False
  If lastDof < 0 Then Exit Function
  If Dir$(stageFolder & Application.PathSeparator & "dof.csv") = vbNullString Then Exit Function
  ReDim seen(0 To lastDof)
  ReDim loadSw(0 To lastDof)
  ReDim loadAp(0 To lastDof)
  ReDim loadBd(0 To lastDof)
  ReDim loadCond(0 To lastDof)
  fileNumber = FreeFile
  Open stageFolder & Application.PathSeparator & "dof.csv" For Input As #fileNumber
  If Not EOF(fileNumber) Then Line Input #fileNumber, lineText
  nRead = 0
  Do While Not EOF(fileNumber)
    Line Input #fileNumber, lineText
    If Len(Trim$(lineText)) = 0 Then GoTo NextDofLine
    parts = Split(lineText, ",")
    If UBound(parts) < 4 Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseLong(parts(0), origDof) Then
      Close #fileNumber
      Exit Function
    End If
    If origDof < 0 Or origDof > lastDof Then
      Close #fileNumber
      Exit Function
    End If
    If seen(origDof) Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseDouble(parts(1), swValue) Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseDouble(parts(2), apValue) Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseDouble(parts(3), bdValue) Then
      Close #fileNumber
      Exit Function
    End If
    If Not FEMIoParseLong(parts(4), condValue) Then
      Close #fileNumber
      Exit Function
    End If
    seen(origDof) = True
    loadSw(origDof) = swValue
    loadAp(origDof) = apValue
    loadBd(origDof) = bdValue
    loadCond(origDof) = condValue
    nRead = nRead + 1
NextDofLine:
  Loop
  Close #fileNumber
  If nRead <> lastDof + 1 Then Exit Function
  ReDim P3LockedSelfWeight(lastDof)
  ReDim P3LockedApplied(lastDof)
  ReDim P3CommittedNodeCond(lastDof)
  ReDim P3CommittedBoundaryDisp(lastDof)
  For origDof = 0 To lastDof
    If Not seen(origDof) Then Exit Function
    origFree = origDof \ 2
    localDof = origDof Mod 2
    intDof = 2 * FEMIoInternalFreeNode(origFree) + localDof
    If intDof < 0 Or intDof > lastDof Then Exit Function
    P3LockedSelfWeight(intDof) = loadSw(origDof)
    P3LockedApplied(intDof) = loadAp(origDof)
    P3BoundaryDisp(intDof) = loadBd(origDof)
    P3CommittedBoundaryDisp(intDof) = loadBd(origDof)
    NodeCond(intDof) = loadCond(origDof)
    P3CommittedNodeCond(intDof) = loadCond(origDof)
  Next origDof
  P3CommittedBoundaryReady = True
  FEMIoLoadDofLock = True
End Function

' Order matters: geometry -> self-weight -> derived principals -> commit. Commit-before-rebuild restores stale Spmat on rollback.
Public Function FEMIoFinishResume() As Boolean
  Dim i As Long
  FEMIoFinishResume = False
  If lastDof >= 0 Then
    For i = 0 To lastDof
      UDisp(i) = TDisp(i)
      P3CommittedDisp(i) = TDisp(i)
    Next i
  End If
  P3HoldOrphanDofs
  femIoResumeHasLock = False
  If Len(femIoResumeFolder) > 0 Then femIoResumeHasLock = FEMIoLoadDofLock(femIoResumeFolder)
  If Not femIoResumeHasLock Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "dof.csv の完全性検査に失敗したため再開できません。", vbObjectError + 3099, -1, -1, 0, 0
    Exit Function
  End If
  P3GravityCommitted = femIoResumeGravity
  If Abs(femIoResumeFs - P3CurrentStrengthFactor) > 0.000000000001 Then
    If Not P3SetMaterialForStrengthFactor(femIoResumeFs) Then Exit Function
  End If
  If femIoResumeFss > 0# Then FSS = femIoResumeFss
  If Not FEMIoRebuildResumeGeometry() Then Exit Function
  P3AssembleSelfWeightOnly
  FEMIoRefreshDerivedStressState
  P3CommitMaterialState
  PlasticLogResumeUnknown
  P3RecountActivePlasticPoints
  P3TrialStateValid = False
  FEMIoFinishResume = True
End Function

' Recompute principals / von Mises / Stmat(8) from sx,sy,txy,sz. Stmat(8) matches Nrf (not 0.5*(sMax-sMin)).
Private Sub FEMIoRefreshDerivedStressState()
  Dim k As Long, im As Long
  Dim sx As Double, sy As Double, sz As Double, txy As Double
  Dim centerValue As Double, radiusValue As Double
  Dim p1 As Double, p2 As Double
  Dim sMax As Double, sMid As Double, sMin As Double
  Dim angleRad As Double
  Dim piValue As Double
  piValue = 3.14159265358979
  For k = 0 To NumberOfElement - 1
    If Elem(k).IsJoint Then GoTo NextDerivedElement
    For im = 0 To 3
      sx = Elem(k).Stmat(0, im)
      sy = Elem(k).Stmat(1, im)
      txy = Elem(k).Stmat(2, im)
      sz = Elem(k).Stmat(16, im)
      centerValue = 0.5 * (sx + sy)
      radiusValue = Sqr((0.5 * (sx - sy)) * (0.5 * (sx - sy)) + txy * txy)
      p1 = centerValue + radiusValue
      p2 = centerValue - radiusValue
      P1SortPrincipal3 p1, p2, sz, sMax, sMid, sMin
      Elem(k).P2PrincipalStress(0, im) = sMax
      Elem(k).P2PrincipalStress(1, im) = sMid
      Elem(k).P2PrincipalStress(2, im) = sMin
      Elem(k).Stmat(3, im) = sMax
      Elem(k).Stmat(4, im) = sMin
      Elem(k).Stmat(5, im) = 0.5 * (sMax - sMin)
      angleRad = 0.5 * P1Atan2(2# * txy, sx - sy)
      Elem(k).Stmat(6, im) = angleRad * 180# / piValue
      Elem(k).Stmat(7, im) = P1VonMises3D(sx, sy, sz, txy)
      If sMin > 0# Then
        Elem(k).Stmat(8, im) = sMax
      ElseIf sMax > 0# Then
        Elem(k).Stmat(8, im) = sMax - sMin
      Else
        Elem(k).Stmat(8, im) = -sMin
      End If
      Elem(k).Stmat(9, im) = Elem(k).P2YieldFunction(im)
    Next im
NextDerivedElement:
  Next k
End Sub

Private Sub FEMIoInferResumeLock()
  Dim i As Long
  If lastDof < 0 Then Exit Sub
  On Error Resume Next
  If UBound(P3LockedSelfWeight) <> lastDof Then ReDim P3LockedSelfWeight(lastDof)
  If UBound(P3LockedApplied) <> lastDof Then ReDim P3LockedApplied(lastDof)
  Err.Clear
  On Error GoTo 0
  If P3PlanHasGravity Then
    P3GravityCommitted = True
    For i = 0 To lastDof
      P3LockedSelfWeight(i) = P3SelfWeightForce(i)
    Next i
  Else
    P3GravityCommitted = False
  End If
  If P3PlanHasApply Then
    For i = 0 To lastDof
      P3LockedApplied(i) = P3AppliedForce(i)
    Next i
  End If
End Sub

Private Sub FEMWriteStageLegend(ByVal ws As Worksheet)
  ws.Cells(1, 8).value2 = "凡例（入力しない）"
  ws.Cells(2, 8).value2 = "種別"
  ws.Cells(2, 9).value2 = "意味"
  ws.Cells(3, 8).value2 = "GRAVITY"
  ws.Cells(3, 9).value2 = "自重。計画に無い場合は載せない。支持は節点の拘束"
  ws.Cells(4, 8).value2 = "LOAD"
  ws.Cells(4, 9).value2 = "載荷シートでその材料の行を載せる（値0の行は無効）"
  ws.Cells(5, 8).value2 = "UNLOAD"
  ws.Cells(5, 9).value2 = "その材料の載荷を外す。要素は残る"
  ws.Cells(6, 8).value2 = "BIRTH"
  ws.Cells(6, 9).value2 = "材料を出す。材料の初期状態をOFFにしておく"
  ws.Cells(7, 8).value2 = "DEATH"
  ws.Cells(7, 9).value2 = "材料を消す。初期状態ONの要素を無効化"
  ws.Cells(8, 8).value2 = "SRM"
  ws.Cells(8, 9).value2 = "途中=パラメータの必要Fs未満なら中断／最後のSRM=Fs探索"
  ws.Cells(9, 8).value2 = "RESET_U"
  ws.Cells(9, 9).value2 = "変位・応力・塑性をゼロ。形状はそのまま。後続SRMの原点"
  ws.Cells(10, 8).value2 = "RESET_STRESS"
  ws.Cells(10, 9).value2 = "座標をX+=uしてから変位・応力をゼロ。後続SRMの原点"
  ws.Cells(11, 8).value2 = "MATSET"
  ws.Cells(11, 9).value2 = "材料番号へコピー、またはパラメータ c=;phi=。後続SRMの原点"
  ws.Cells(13, 8).value2 = "材料番号"
  ws.Cells(13, 9).value2 = "LOAD/UNLOAD/BIRTH/DEATH/MATSETの対象。空ならその種別では使わない"
  ws.Cells(14, 8).value2 = "パラメータ"
  ws.Cells(14, 9).value2 = "SRM=必要Fs。MATSET=コピー元材料番号または c=10;phi=30"
  ws.Cells(15, 8).value2 = "有効"
  ws.Cells(15, 9).value2 = "1=実行 / 0または空=飛ばす"
  ws.Cells(17, 8).value2 = "例1"
  ws.Cells(17, 9).value2 = "GRAVITY→SRM"
  ws.Cells(18, 8).value2 = "例2"
  ws.Cells(18, 9).value2 = "GRAVITY→SRM→RESET_U→MATSET→GRAVITY→SRM"
  ws.Cells(19, 8).value2 = "詳細"
  ws.Cells(19, 9).value2 = "「仕様」シート"
End Sub

Private Sub FEMEnsureJointSheet()
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("接合")
  On Error GoTo 0
  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("載荷"))
    ws.name = "接合"
  End If
  ws.Cells(1, 1).value2 = "材料A"
  ws.Cells(1, 2).value2 = "材料B"
  ws.Cells(1, 3).value2 = "接合材料"
  ws.Cells(1, 4).value2 = "有効"
  ws.Cells(1, 5).value2 = "説明"
  ws.Cells(1, 7).value2 = "凡例（入力しない）"
  ws.Cells(2, 7).value2 = "材料A"
  ws.Cells(2, 8).value2 = "節点を残す側（擁壁・構造）"
  ws.Cells(3, 7).value2 = "材料B"
  ws.Cells(3, 8).value2 = "節点を増やす側（地盤）"
  ws.Cells(4, 7).value2 = "接合材料"
  ws.Cells(4, 8).value2 = "種別JOINTの材料番号"
  ws.Cells(5, 7).value2 = "生成"
  ws.Cells(5, 8).value2 = "要素作成②の後。共有辺を6節点接合にする"
  ws.Cells(6, 7).value2 = "有効"
  ws.Cells(6, 8).value2 = "1=その行を使う / 0または空=使わない"
  ws.Cells(7, 7).value2 = "順序"
  ws.Cells(7, 8).value2 = "材料Aの節点を残し、材料B側に新節点を付ける"
End Sub

Private Sub FEMEnsureLoadSheetLayout(ByVal ws As Worksheet)
  Dim lastRow As Long, rowNo As Long, headerF As String, selectorText As String
  Dim boxParts() As String, noteText As String
  headerF = UCase$(Trim$(CStr(ws.Cells(1, 6).value2)))
  lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
  If lastRow < 2 Then lastRow = 2
  If headerF = "説明" Then
    For rowNo = 2 To lastRow
      noteText = Trim$(CStr(ws.Cells(rowNo, 6).value2))
      If Len(Trim$(CStr(ws.Cells(rowNo, 10).value2))) = 0 Then ws.Cells(rowNo, 10).value2 = noteText
      ws.Cells(rowNo, 6).ClearContents
    Next rowNo
  End If
  For rowNo = 2 To lastRow
    selectorText = UCase$(Trim$(CStr(ws.Cells(rowNo, 2).value2)))
    If left$(selectorText, 4) = "BOX:" Then
      boxParts = Split(mid$(selectorText, 5), ",")
      ws.Cells(rowNo, 2).value2 = "BOX"
      If UBound(boxParts) >= 0 Then If Len(Trim$(CStr(ws.Cells(rowNo, 6).value2))) = 0 Then ws.Cells(rowNo, 6).value2 = Val(boxParts(0))
      If UBound(boxParts) >= 1 Then If Len(Trim$(CStr(ws.Cells(rowNo, 7).value2))) = 0 Then ws.Cells(rowNo, 7).value2 = Val(boxParts(1))
      If UBound(boxParts) >= 2 Then If Len(Trim$(CStr(ws.Cells(rowNo, 8).value2))) = 0 Then ws.Cells(rowNo, 8).value2 = Val(boxParts(2))
      If UBound(boxParts) >= 3 Then If Len(Trim$(CStr(ws.Cells(rowNo, 9).value2))) = 0 Then ws.Cells(rowNo, 9).value2 = Val(boxParts(3))
    End If
  Next rowNo
  ' BOX coordinate columns H:I are input; preserve them during layout refresh.
  ws.Cells(1, 1).value2 = "材料番号"
  ws.Cells(1, 2).value2 = "選択"
  ws.Cells(1, 3).value2 = "方向"
  ws.Cells(1, 4).value2 = "種別"
  ws.Cells(1, 5).value2 = "値"
  ws.Cells(1, 6).value2 = "X1"
  ws.Cells(1, 7).value2 = "Y1"
  ws.Cells(1, 8).value2 = "X2"
  ws.Cells(1, 9).value2 = "Y2"
  ws.Cells(1, 10).value2 = "説明"
  If Len(Trim$(CStr(ws.Cells(2, 1).value2))) = 0 Or UCase$(Trim$(CStr(ws.Cells(2, 1).value2))) = "基礎" Then
    ws.Cells(2, 1).value2 = 1
    ws.Cells(2, 2).value2 = "TOP"
    ws.Cells(2, 3).value2 = "Y"
    ws.Cells(2, 4).value2 = "DISP"
    If Len(Trim$(CStr(ws.Cells(2, 5).value2))) = 0 Then ws.Cells(2, 5).value2 = 0#
  End If
  FEMApplyLoadValidation ws
  FEMWriteLoadLegend ws
End Sub

Private Sub FEMApplyLoadValidation(ByVal ws As Worksheet)
  Dim targetRange As range
  On Error Resume Next
  Set targetRange = ws.range("B2:B100")
  targetRange.Validation.Delete
  targetRange.Validation.Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="TOP,BOTTOM,LEFT,RIGHT,BOX"
  targetRange.Validation.IgnoreBlank = True
  targetRange.Validation.InCellDropdown = True
  Set targetRange = ws.range("C2:C100")
  targetRange.Validation.Delete
  targetRange.Validation.Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="X,Y"
  targetRange.Validation.IgnoreBlank = True
  targetRange.Validation.InCellDropdown = True
  Set targetRange = ws.range("D2:D100")
  targetRange.Validation.Delete
  targetRange.Validation.Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="DISP,FORCE,FIX"
  targetRange.Validation.IgnoreBlank = True
  targetRange.Validation.InCellDropdown = True
  On Error GoTo 0
End Sub

Private Sub FEMWriteLoadLegend(ByVal ws As Worksheet)
  ws.Cells(1, 12).value2 = "凡例（入力しない）"
  ws.Cells(2, 12).value2 = "選択"
  ws.Cells(2, 13).value2 = "どの辺・範囲に載せるか。プルダウン"
  ws.Cells(3, 12).value2 = "TOP"
  ws.Cells(3, 13).value2 = "その材料の上辺（Y最大の辺）"
  ws.Cells(4, 12).value2 = "BOTTOM"
  ws.Cells(4, 13).value2 = "その材料の下辺"
  ws.Cells(5, 12).value2 = "LEFT"
  ws.Cells(5, 13).value2 = "その材料の左辺"
  ws.Cells(6, 12).value2 = "RIGHT"
  ws.Cells(6, 13).value2 = "その材料の右辺"
  ws.Cells(7, 12).value2 = "BOX"
  ws.Cells(7, 13).value2 = "X1,Y1,X2,Y2の矩形。解析座標。その材料の節点だけ"
  ws.Cells(8, 12).value2 = "方向"
  ws.Cells(8, 13).value2 = "X または Y"
  ws.Cells(9, 12).value2 = "DISP"
  ws.Cells(9, 13).value2 = "強制変位。値の単位は長さ"
  ws.Cells(10, 12).value2 = "FORCE"
  ws.Cells(10, 13).value2 = "節点力。値の単位は力"
  ws.Cells(11, 12).value2 = "FIX"
  ws.Cells(11, 13).value2 = "その方向を拘束（変位0）"
  ws.Cells(12, 12).value2 = "値0"
  ws.Cells(12, 13).value2 = "DISP/FORCEはその行を無効。FIXは0でも拘束"
  ws.Cells(13, 12).value2 = "材料番号"
  ws.Cells(13, 13).value2 = "材料データと同じ番号。LOAD/UNLOADの対象"
End Sub

Private Sub FEMEnsureMaterialSheetLayout()
  Dim ws As Worksheet, lastRow As Long, rowNo As Long
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("材料データ")
  On Error GoTo 0
  If ws Is Nothing Then Exit Sub
  ws.Cells(1, 1).value2 = "材料番号"
  ws.Cells(1, 2).value2 = "ヤング率"
  ws.Cells(1, 3).value2 = "ポアソン比"
  ws.Cells(1, 4).value2 = "板厚"
  ws.Cells(1, 5).value2 = "単位体積重量"
  ws.Cells(1, 6).value2 = "内部摩擦角"
  ws.Cells(1, 7).value2 = "粘着力"
  ws.Cells(1, 8).value2 = "ダイレタンシー角"
  ws.Cells(1, 9).value2 = "強度低減"
  ws.Cells(1, 10).value2 = "初期状態"
  ws.Cells(1, 11).value2 = "種別"
  ws.Cells(1, 12).value2 = "せん断剛性"
  ws.Cells(1, 13).value2 = "引張"
  ws.Cells(1, 15).value2 = "凡例（入力しない）"
  ws.Cells(2, 15).value2 = "強度低減"
  ws.Cells(2, 16).value2 = "1=SRMでc・φを下げる / 0=構造物・基盤は下げない"
  ws.Cells(3, 15).value2 = "初期状態"
  ws.Cells(3, 16).value2 = "ON=解析開始から有効 / OFF=BIRTH待ち"
  ws.Cells(4, 15).value2 = "種別"
  ws.Cells(4, 16).value2 = "SOIL=地盤 / STRUCT=構造 / JOINT=接合。空=SOIL"
  ws.Cells(5, 15).value2 = "JOINT"
  ws.Cells(5, 16).value2 = "ヤング率=kn、せん断剛性=ks、引張0=剥離可、引張1=引張も負担"
  ws.Cells(6, 15).value2 = "DAVIS"
  ws.Cells(6, 16).value2 = "FLOW_POLICY=DAVISでもJOINTのc,φは変換しない"
  ws.Cells(7, 15).value2 = "単位"
  ws.Cells(7, 16).value2 = "角度は度。粘着力は応力と同じ単位"
  lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
  If lastRow < 2 Then lastRow = 2
  For rowNo = 2 To lastRow
    If Len(Trim$(CStr(ws.Cells(rowNo, 1).value2))) = 0 Then GoTo NextMaterialRow
    If Len(Trim$(CStr(ws.Cells(rowNo, 9).value2))) = 0 Then ws.Cells(rowNo, 9).value2 = 1
    If Len(Trim$(CStr(ws.Cells(rowNo, 10).value2))) = 0 Then ws.Cells(rowNo, 10).value2 = "ON"
    ' Old legend used L/M (ks/tension). Clear leftover labels on data rows.
    If Not IsNumeric(ws.Cells(rowNo, 12).value2) Then
      If InStr(1, CStr(ws.Cells(rowNo, 12).value2), "強度") > 0 Or InStr(1, CStr(ws.Cells(rowNo, 12).value2), "凡例") > 0 Then
        ws.Cells(rowNo, 12).ClearContents
        If Not IsNumeric(ws.Cells(rowNo, 13).value2) Then ws.Cells(rowNo, 13).ClearContents
      End If
    End If
NextMaterialRow:
  Next rowNo
  FEMApplyMaterialValidation ws
End Sub

Private Sub FEMApplyListValidation(ByVal targetRange As range, ByVal listText As String)
  On Error Resume Next
  targetRange.Validation.Delete
  If Len(Trim$(listText)) > 0 Then
    targetRange.Validation.Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:=listText
    targetRange.Validation.IgnoreBlank = True
    targetRange.Validation.InCellDropdown = True
    targetRange.Validation.ShowError = True
    targetRange.Validation.ErrorTitle = "入力値"
    targetRange.Validation.ErrorMessage = "リストの値を選んでください。"
  End If
  Err.Clear
  On Error GoTo 0
End Sub

Private Function FEMSettingChoiceList(ByVal keyName As String) As String
  Select Case UCase$(Trim$(keyName))
    Case "ACCEL_STEP_FRESH_LU", "ACCEL_STEP_RECOVERY", "ACCEL_ADAPTIVE", "ACCEL_V1_PREDICTOR", "ACCEL_V2A_REUSE", "ACCEL_V2B_COST", "ACCEL_V3_COST_SEARCH", "ACCEL_V4_ANDERSON", "ACCEL_V5_GMRES_LU", "ACCEL_TRACE"
      FEMSettingChoiceList = "0,1"
    Case "MESH_BC_BOTTOM", "MESH_BC_LEFT", "MESH_BC_RIGHT", "MESH_BC_TOP"
      FEMSettingChoiceList = "NONE,ROLLER,PINNED"
    Case "MESH_BC_PIN_CORNER"
      FEMSettingChoiceList = "AUTO,NONE"
    Case "WEIGHT_MODE"
      FEMSettingChoiceList = "TOTAL"
    Case "RCM_POLICY"
      FEMSettingChoiceList = "AUTO,ON,OFF"
    Case "FLOW_POLICY"
      FEMSettingChoiceList = "INCONSISTENT,DAVIS"
    Case "SRM_PSI_POLICY"
      FEMSettingChoiceList = "CAP,REDUCE,KEEP"
    Case "SRM_MODE"
      FEMSettingChoiceList = "REAPPLY"
    Case "OUTPUT_STAGE_MODE"
      FEMSettingChoiceList = "FINAL,ALL"
    Case "OUTPUT_INACTIVE_ELEMENTS"
      FEMSettingChoiceList = "SKIP,INCLUDE"
    Case "DEBUG_MODE"
      FEMSettingChoiceList = "OFF,STAGE,ITER"
    Case "SOLVER"
      FEMSettingChoiceList = "BAND_LU,BAND_LDLT_EXPERIMENTAL"
    Case "BAND_LU_KERNEL"
      FEMSettingChoiceList = "NEW,OLD"
    Case "EXPORT_MODE"
      FEMSettingChoiceList = "OFF,STAGE,ON_FAIL"
    Case "VIEW_SCALE_MODE"
      FEMSettingChoiceList = "AUTO,SYMMETRIC,USER"
    Case "VIEW_DETAIL_MODE"
      FEMSettingChoiceList = "SIMPLE,DETAIL"
    Case "VIEW_RESULT_SOURCE"
      FEMSettingChoiceList = "SMOOTHED,ELEMENT,RAW_GAUSS"
    Case "MESH_GEN_COMPARE"
      FEMSettingChoiceList = "0,1,2"
    Case Else
      FEMSettingChoiceList = vbNullString
  End Select
End Function

Private Sub FEMWriteSettingsChoiceGuide(ByVal ws As Worksheet)
  Dim lastRow As Long, rowNo As Long, keyName As String, listText As String
  lastRow = ws.Cells(ws.rows.count, 5).End(xlUp).row
  On Error Resume Next
  ws.range(ws.Cells(4, 6), ws.Cells(200, 6)).ClearContents
  Err.Clear
  On Error GoTo 0
  ws.Cells(4, 6).NumberFormat = "@"
  ws.Cells(4, 6).value2 = "選択肢"
  For rowNo = 5 To lastRow
    keyName = UCase$(Trim$(CStr(ws.Cells(rowNo, 5).value2)))
    listText = FEMSettingChoiceList(keyName)
    If Len(listText) > 0 Then
      ws.Cells(rowNo, 6).value2 = Replace$(listText, ",", " / ")
    ElseIf keyName = "EXPORT_LOAD" Then
      ws.Cells(rowNo, 6).value2 = "空 / AUTO / フォルダ名"
    ElseIf keyName = "MESH_REPAIR_TERRAIN_COMPARE" Then
      ws.Cells(rowNo, 6).NumberFormat = "@"
      ws.Cells(rowNo, 6).value2 = "1=する  0=しない"
    ElseIf keyName = "ACCEL_PREDICT_BETA" Then
      ws.Cells(rowNo, 6).value2 = "0以上1以下"
    ElseIf keyName = "ACCEL_PREDICT_MAX_RATIO" Then
      ws.Cells(rowNo, 6).value2 = "0より大きく2以下"
    ElseIf keyName = "ACCEL_TANGENT_CHANGE" Then
      ws.Cells(rowNo, 6).value2 = "0より大きく0.1以下"
    ElseIf keyName = "SRM_FIXED_FS" Then
      ws.Cells(rowNo, 6).value2 = "0=通常探索 / 正数=固定Fs"
    End If
  Next rowNo
End Sub

Private Sub FEMApplySettingValidation(ByVal ws As Worksheet)
  Dim lastRow As Long, rowNo As Long, keyName As String, listText As String
  lastRow = ws.Cells(ws.rows.count, 5).End(xlUp).row
  For rowNo = 5 To lastRow
    keyName = UCase$(Trim$(CStr(ws.Cells(rowNo, 5).value2)))
    listText = FEMSettingChoiceList(keyName)
    FEMApplyListValidation ws.Cells(rowNo, 3), listText
    Select Case keyName
      Case "ACCEL_PREDICT_BETA"
        FEMApplyNumericValidation ws.Cells(rowNo, 3), "=AND(ISNUMBER(C" & rowNo & "),C" & rowNo & ">=0,C" & rowNo & "<=1)"
      Case "ACCEL_PREDICT_MAX_RATIO"
        FEMApplyNumericValidation ws.Cells(rowNo, 3), "=AND(ISNUMBER(C" & rowNo & "),C" & rowNo & ">0,C" & rowNo & "<=2)"
      Case "ACCEL_TANGENT_CHANGE"
        FEMApplyNumericValidation ws.Cells(rowNo, 3), "=AND(ISNUMBER(C" & rowNo & "),C" & rowNo & ">0,C" & rowNo & "<=0.1)"
      Case "SRM_FIXED_FS"
        FEMApplyNumericValidation ws.Cells(rowNo, 3), "=AND(ISNUMBER(C" & rowNo & "),C" & rowNo & ">=0)"
      Case "VIEW_COLOR_RESULT"
        FEMApplyListValidation ws.Cells(rowNo, 3), "=FEMViewerSelectionIDs"
    End Select
  Next rowNo
End Sub

Private Sub FEMApplyNumericValidation(ByVal targetRange As range, ByVal ruleText As String)
  With targetRange.Validation
    .Delete
    .Add Type:=7, AlertStyle:=1, Formula1:=ruleText
    .IgnoreBlank = False
    .ShowError = True
    .ErrorTitle = "入力値"
    .ErrorMessage = "F列の範囲内の数値を入力してください。"
  End With
End Sub

Private Sub FEMApplyStageValidation(ByVal ws As Worksheet)
  FEMApplyListValidation ws.range("B2:B100"), "GRAVITY,LOAD,UNLOAD,BIRTH,DEATH,SRM,RESET_U,RESET_STRESS,MATSET"
  FEMApplyListValidation ws.range("E2:E100"), "1,0"
End Sub

Private Sub FEMApplyMaterialValidation(ByVal ws As Worksheet)
  FEMApplyListValidation ws.range("I2:I100"), "1,0"
  FEMApplyListValidation ws.range("J2:J100"), "ON,OFF"
  FEMApplyListValidation ws.range("K2:K100"), "SOIL,STRUCT,JOINT"
  FEMApplyListValidation ws.range("M2:M100"), "0,1"
End Sub

Private Sub FEMSpecPut(ByVal ws As Worksheet, ByRef rowNo As Long, ByVal classText As String, ByVal itemText As String, ByVal valueText As String, ByVal meaningText As String)
  ws.Cells(rowNo, 1).value2 = classText
  ws.Cells(rowNo, 2).value2 = itemText
  ws.Cells(rowNo, 3).value2 = valueText
  ws.Cells(rowNo, 4).value2 = meaningText
  rowNo = rowNo + 1
End Sub

Private Sub FEMEnsureSpecSheet()
  Dim ws As Worksheet, rowNo As Long
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("仕様")
  On Error GoTo 0
  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("設定"))
    ws.name = "仕様"
  End If
  ws.Cells.ClearContents
  ws.Cells(1, 1).value2 = "仕様（入力しない。開くたびに再生成）"
  ws.Cells(2, 1).value2 = "分類"
  ws.Cells(2, 2).value2 = "項目"
  ws.Cells(2, 3).value2 = "値・入力"
  ws.Cells(2, 4).value2 = "意味"
  ws.columns(3).NumberFormat = "@"
  rowNo = 3
  FEMSpecPut ws, rowNo, "このブック", "設定", "C列", "数値と選択肢。F列が候補。G1がプログラムVer"
  FEMSpecPut ws, rowNo, "このブック", "材料データ", "1行=1材料", "地盤・構造・接合を同じ表に書く"
  FEMSpecPut ws, rowNo, "このブック", "ステージ", "上から順", "解析手順。有効=1の行だけ実行"
  FEMSpecPut ws, rowNo, "このブック", "載荷", "LOAD/UNLOAD", "どの材料のどの辺に変位・力・拘束を載せるか"
  FEMSpecPut ws, rowNo, "このブック", "接合", "要素作成②の後", "共有辺をJOINT要素にする"
  FEMSpecPut ws, rowNo, "このブック", "節点・要素", "メッシュ", "拘束は0/1。載荷のFIXでも拘束できる"
  FEMSpecPut ws, rowNo, "このブック", "再開", "EXPORT_*", "成功ステージがブック横 *_out に保存される"
  FEMSpecPut ws, rowNo, "ステージ", "GRAVITY", "", "自重。支持は節点拘束。計画に無ければ載せない"
  FEMSpecPut ws, rowNo, "ステージ", "LOAD", "材料番号", "載荷シートのその材料を有効化"
  FEMSpecPut ws, rowNo, "ステージ", "UNLOAD", "材料番号", "その材料の載荷を外す。要素は残る"
  FEMSpecPut ws, rowNo, "ステージ", "BIRTH", "材料番号", "初期状態OFFの材料を出す"
  FEMSpecPut ws, rowNo, "ステージ", "DEATH", "材料番号", "初期状態ONの材料を消す"
  FEMSpecPut ws, rowNo, "ステージ", "SRM", "パラメータ=必要Fs", "途中=ゲート。最後のSRM=Fs探索（設定の上限Fs・許容差）"
  FEMSpecPut ws, rowNo, "ステージ", "RESET_U", "", "変位・応力・塑性をゼロ。形状そのまま。後続SRMの原点"
  FEMSpecPut ws, rowNo, "ステージ", "RESET_STRESS", "", "形状を変形後座標にしてからゼロ。checkpointはx,yも保存"
  FEMSpecPut ws, rowNo, "ステージ", "MATSET", "コピー元 or c=;phi=", "材料を差し替え。後続SRMの原点"
  FEMSpecPut ws, rowNo, "ステージ", "有効", "1 または 0", "0または空はその行を飛ばす"
  FEMSpecPut ws, rowNo, "設定", "底面・左・右・上面拘束", "NONE / ROLLER / PINNED", "NONE=自由。ROLLER=面に直角方向だけ固定。PINNED=XY固定。要素作成②の後にも再適用"
  FEMSpecPut ws, rowNo, "設定", "剛体モード止め", "AUTO / NONE", "AUTO=底面の最小X節点をXY固定して剛体変位を止める"
  FEMSpecPut ws, rowNo, "設定", "自重", "TOTAL", "全応力の単位体積重量。浮力・水位は未実装"
  FEMSpecPut ws, rowNo, "設定", "RCM", "AUTO / ON / OFF", "帯幅縮小の節点並び替え。checkpointは元の節点番号で持つ"
  FEMSpecPut ws, rowNo, "設定", "非関連の扱い", "INCONSISTENT / DAVIS", "INCONSISTENT=流れ則は非関連、接線は対称近似。DAVIS=等価関連c*,φ*。JOINTのc,φには掛けない"
  FEMSpecPut ws, rowNo, "設定", "ψの扱い", "CAP / REDUCE / KEEP", "SRM時のダイレタンシー。CAP=φ'以下、REDUCE=Fsで低減、KEEP=入力のまま"
  FEMSpecPut ws, rowNo, "設定", "SRMモード", "REAPPLY", "各Fs試行で先行ステージを再実行する"
  FEMSpecPut ws, rowNo, "設定", "結果の残し方", "FINAL / ALL", "FINAL=最後だけ結果シートへ。ALL=各ステージを追記"
  FEMSpecPut ws, rowNo, "設定", "無効要素", "SKIP / INCLUDE", "Death中の出力。SKIP=出さない。INCLUDE=行は残す"
  FEMSpecPut ws, rowNo, "設定", "デバッグログ", "OFF / STAGE / ITER", "ブック横 *_out/run.log"
  FEMSpecPut ws, rowNo, "設定", "途中データの保存", "OFF / STAGE / ON_FAIL", "成功ステージをフォルダへ。ON_FAILは失敗時も保存"
  FEMSpecPut ws, rowNo, "設定", "再開フォルダ", "空 / AUTO / フォルダ名", "空=最初から。AUTO=最後のSUCCESS。RESET後にSRMする再開には origin.csv が必要"
  FEMSpecPut ws, rowNo, "設定", "色スケール", "AUTO / SYMMETRIC / USER", "コンターの凡例"
  FEMSpecPut ws, rowNo, "設定", "表示モード", "SIMPLE / DETAIL", "図の詳細さ"
  FEMSpecPut ws, rowNo, "設定", "色の元", "SMOOTHED / ELEMENT / RAW_GAUSS", "節点平滑 / 要素平均 / Gauss点"
  FEMSpecPut ws, rowNo, "設定", "生成時比較", "0 / 1 / 2", "両方合格のとき。0=品質の良い方。1=角度最適化（デローニ修正）。2=地形追従"
  FEMSpecPut ws, rowNo, "材料", "強度低減", "1 または 0", "1=SRMでc・φを下げる。0=構造物・基盤"
  FEMSpecPut ws, rowNo, "材料", "初期状態", "ON / OFF", "ON=最初から有効。OFF=BIRTH待ち"
  FEMSpecPut ws, rowNo, "材料", "種別", "SOIL / STRUCT / JOINT", "空はSOIL。JOINTはkn=ヤング率、ks=せん断剛性"
  FEMSpecPut ws, rowNo, "材料", "引張（JOINT）", "0 または 1", "0=剥離可。1=引張も負担"
  FEMSpecPut ws, rowNo, "載荷", "選択", "TOP / BOTTOM / LEFT / RIGHT / BOX", "BOXだけX1～Y2を使う。解析座標、その材料の節点だけ"
  FEMSpecPut ws, rowNo, "載荷", "方向", "X / Y", "作用方向"
  FEMSpecPut ws, rowNo, "載荷", "種別", "DISP / FORCE / FIX", "強制変位 / 節点力 / 拘束。DISPとFORCEは値0で無効"
  FEMSpecPut ws, rowNo, "接合", "材料A", "番号", "節点を残す側（構造・擁壁）"
  FEMSpecPut ws, rowNo, "接合", "材料B", "番号", "新節点を付ける側（地盤）"
  FEMSpecPut ws, rowNo, "接合", "接合材料", "JOINT番号", "種別JOINTの材料"
  FEMSpecPut ws, rowNo, "節点", "X/Y拘束条件", "0 または 1", "0=自由、1=固定"
  FEMSpecPut ws, rowNo, "再開ファイル", "node.csv", "x,y,ux,uy", "RESET_STRESS後はx,yが新しい基準形状"
  FEMSpecPut ws, rowNo, "再開ファイル", "gauss.csv", "応力と塑性", "主応力は再開時に再計算する"
  FEMSpecPut ws, rowNo, "再開ファイル", "dof.csv", "拘束とロック荷重", "必須。欠けると再開しない"
  FEMSpecPut ws, rowNo, "再開ファイル", "origin.csv", "RESET/MATSET原点", "後続SRMのreplay用。無いとSRM再開は停止する"
  FEMSpecPut ws, rowNo, "再開", "BIRTHの再現", "履歴は消さない", "gauss.csvを読んだあと skipped BIRTH は active だけ戻す"
  On Error Resume Next
  ws.columns("A:D").AutoFit
  ws.Move After:=ThisWorkbook.Worksheets("設定")
  Err.Clear
  On Error GoTo 0
End Sub

Private Sub P6EnsureSettingRow(ByVal ws As Worksheet, ByVal keyName As String, ByVal categoryText As String, ByVal itemText As String, ByVal defaultValue As Double, ByVal noteText As String)
  Dim lastRow As Long, rowNo As Long, foundRow As Long
  foundRow = 0
  lastRow = ws.Cells(ws.rows.count, 5).End(xlUp).row
  For rowNo = 1 To lastRow
    If UCase$(Trim$(CStr(ws.Cells(rowNo, 5).value2))) = UCase$(keyName) Then
      foundRow = rowNo
      Exit For
    End If
  Next rowNo
  If foundRow = 0 Then
    foundRow = lastRow + 1
    ws.Cells(foundRow, 5).value2 = keyName
    ws.Cells(foundRow, 3).value2 = defaultValue
  End If
  ws.Cells(foundRow, 1).value2 = categoryText
  ws.Cells(foundRow, 2).value2 = itemText
  ws.Cells(foundRow, 4).value2 = noteText
End Sub

Private Sub P6EnsureTextSettingRow(ByVal ws As Worksheet, ByVal keyName As String, ByVal categoryText As String, ByVal itemText As String, ByVal defaultText As String, ByVal noteText As String)
  Dim lastRow As Long, rowNo As Long, foundRow As Long
  foundRow = 0
  lastRow = ws.Cells(ws.rows.count, 5).End(xlUp).row
  For rowNo = 1 To lastRow
    If UCase$(Trim$(CStr(ws.Cells(rowNo, 5).value2))) = UCase$(keyName) Then
      foundRow = rowNo
      Exit For
    End If
  Next rowNo
  If foundRow = 0 Then
    foundRow = lastRow + 1
    ws.Cells(foundRow, 5).value2 = keyName
    ws.Cells(foundRow, 3).NumberFormat = "@"
    ws.Cells(foundRow, 3).value2 = defaultText
  End If
  ws.Cells(foundRow, 1).value2 = categoryText
  ws.Cells(foundRow, 2).value2 = itemText
  ws.Cells(foundRow, 4).value2 = noteText
End Sub

Public Sub FEMResetRunLog()
  Dim ws As Worksheet
  Dim lastRow As Long
  On Error Resume Next
  Set ws = GetIntegratedResultSheet()
  On Error GoTo 0
  If ws Is Nothing Then
    P3RunLogRow = 0
    Exit Sub
  End If
  lastRow = ws.Cells(ws.rows.count, 4).End(xlUp).row
  If lastRow < 2 Then lastRow = 2
  ws.range(ws.Cells(1, 4), ws.Cells(lastRow, 14)).ClearContents
  ws.Cells(1, 4).value2 = "解析ログ（再実行で消去） Ver=" & FEM_BUILD_STAMP
  ws.Cells(2, 4).value2 = "順"
  ws.Cells(2, 5).value2 = "場所"
  ws.Cells(2, 6).value2 = "ステージ"
  ws.Cells(2, 7).value2 = "種別"
  ws.Cells(2, 8).value2 = "状態"
  ws.Cells(2, 9).value2 = "Fs"
  ws.Cells(2, 10).value2 = "増分"
  ws.Cells(2, 11).value2 = "反復"
  ws.Cells(2, 12).value2 = "相対残差"
  ws.Cells(2, 13).value2 = "補正比"
  ws.Cells(2, 14).value2 = "メッセージ"
  P3RunLogRow = 2
  P3RunLogStageNo = 0
  P3RunLogKind = vbNullString
End Sub

Public Sub FEMAppendRunLog(ByVal placeText As String, ByVal statusText As String, ByVal messageText As String)
  Dim ws As Worksheet
  Dim stageNo As Long
  Dim kindText As String
  On Error Resume Next
  Set ws = GetIntegratedResultSheet()
  If ws Is Nothing Then Exit Sub
  If P3RunLogRow < 2 Then FEMResetRunLog
  P3RunLogRow = P3RunLogRow + 1
  stageNo = P3RunLogStageNo
  kindText = P3RunLogKind
  If stageNo <= 0 And P3ActivePrefix >= 1 And P3ActivePrefix <= P3StageN Then
    stageNo = P3StageId(P3ActivePrefix)
    kindText = P3StageKind(P3ActivePrefix)
  End If
  ws.Cells(P3RunLogRow, 4).value2 = P3RunLogRow - 2
  ws.Cells(P3RunLogRow, 5).value2 = placeText
  If stageNo > 0 Then ws.Cells(P3RunLogRow, 6).value2 = stageNo
  ws.Cells(P3RunLogRow, 7).value2 = kindText
  ws.Cells(P3RunLogRow, 8).value2 = statusText
  ws.Cells(P3RunLogRow, 9).value2 = P3CurrentStrengthFactor
  ws.Cells(P3RunLogRow, 10).value2 = CurrentIncrement
  ws.Cells(P3RunLogRow, 11).value2 = CurrentIteration
  ws.Cells(P3RunLogRow, 12).value2 = RelativeResidualFree
  ws.Cells(P3RunLogRow, 13).value2 = P3MaxCorrection
  ws.Cells(P3RunLogRow, 14).value2 = messageText
  If P3RunLogRow > 80 Then
    ' 診断シートは要約。本文は *_out/run.log
    P3RunLogRow = 80
  End If
  FEMIoLog placeText, statusText, messageText
End Sub

Public Function GetIntegratedResultSheet() As Worksheet
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets(FEM_INTEGRATED_RESULT_SHEET)
  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets("解析結果")
    If Not ws Is Nothing Then ws.name = FEM_INTEGRATED_RESULT_SHEET
  End If
  On Error GoTo 0
  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.name = FEM_INTEGRATED_RESULT_SHEET
  End If
  Set GetIntegratedResultSheet = ws
End Function

Public Function GetDiagnosticSheet() As Worksheet
  Dim ws As Worksheet
  Set ws = GetIntegratedResultSheet()
  Set GetDiagnosticSheet = ws
End Function


Private Sub FEMIoFlushPolicy(ByVal statusText As String)
  Dim lines As String, pathText As String, fileNum As Integer
  If femIoMode = "OFF" Then Exit Sub
  lines = PolicyLogDrain()
  If Len(lines) > 0 Then FEMIoAppendPolicyCsv "adaptive_v2a.csv", lines
  lines = PolicyLogGateRows(statusText)
  If Len(lines) > 0 Then FEMIoAppendPolicyCsv "adaptive_v2a_gates.csv", lines
  lines = IncrementLogDrain()
  If Len(lines) > 0 Then FEMIoAppendPolicyCsv "increment_summary.csv", lines
End Sub

Private Sub FEMIoAppendPolicyCsv(ByVal fileName As String, ByVal lines As String)
  Dim pathText As String, fileNum As Integer
  On Error GoTo WriteFailed
  pathText = FEMIoOutDir() & Application.PathSeparator & fileName
  fileNum = FreeFile
  Open pathText For Append As #fileNum
  Print #fileNum, lines
  Close #fileNum
  Exit Sub
WriteFailed:
  On Error Resume Next
  Close #fileNum
  Err.Clear
End Sub
