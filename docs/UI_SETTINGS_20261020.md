# UI版の設定一覧（2026-10-20）

UI_03では旧簡易メッシュの `MESH_POINT*` / `MESH_NX` / `MESH_NY` を非表示・不使用にしています。互換性のため値は保持します。現在の操作は[UI_03手順](SHAPE_MESH_UI_20261020.md)を参照してください。

配布XLSMから読み取った初期入力値とG列の推奨・目安です。材料・荷重・支持・Fs等の条件を一括変更するものではありません。
設定値の変更はC列だけで行います。再表示・再読込はユーザーの設定値を保持します。
[操作・対応機能・検証記録](WORKBOOK_UI_20261020.md)。自動出力は編集しません。

## メッシュ設定

| 入力セル | 項目 / KEY | 配布保存値 | 推奨・目安 |
| --- | --- | --- | --- |
| C9 | 生成要素の材料番号<br>`MESH_MATERIAL` | 1 | モデル・表示による |
| C10 | 底面拘束<br>`MESH_BC_BOTTOM` | PINNED | モデル・表示による |
| C11 | 左面拘束<br>`MESH_BC_LEFT` | ROLLER | モデル・表示による |
| C12 | 右面拘束<br>`MESH_BC_RIGHT` | ROLLER | モデル・表示による |
| C13 | 上面拘束<br>`MESH_BC_TOP` | NONE | モデル・表示による |
| C14 | 剛体モード止め<br>`MESH_BC_PIN_CORNER` | AUTO | モデル・表示による |
| C17 | 点1 X<br>`MESH_POINT1_X` | 0 | モデル・表示による |
| C18 | 点1 Y<br>`MESH_POINT1_Y` | 0 | モデル・表示による |
| C19 | 点2 X<br>`MESH_POINT2_X` | 0 | モデル・表示による |
| C20 | 点2 Y<br>`MESH_POINT2_Y` | 4 | モデル・表示による |
| C21 | 点3 X<br>`MESH_POINT3_X` | 3 | モデル・表示による |
| C22 | 点3 Y<br>`MESH_POINT3_Y` | 4 | モデル・表示による |
| C23 | 点4 X<br>`MESH_POINT4_X` | 3 | モデル・表示による |
| C24 | 点4 Y<br>`MESH_POINT4_Y` | 0 | モデル・表示による |
| C25 | 横方向の分割数 nx<br>`MESH_NX` | 25 | モデル・表示による |
| C26 | 縦方向の分割数 ny<br>`MESH_NY` | 25 | モデル・表示による |
| C29 | 形状メッシュの生成方式<br>`MESH_GEN_COMPARE` | 2 | モデル・表示による |

## 解析設定

| 入力セル | 項目 / KEY | 配布保存値 | 推奨・目安 |
| --- | --- | --- | --- |
| C9 | 自重<br>`WEIGHT_MODE` | TOTAL | TOTAL |
| C10 | 非関連の扱い<br>`FLOW_POLICY` | INCONSISTENT | INCONSISTENT |
| C11 | アワーグラス係数<br>`Q8_HOURGLASS_FACTOR` | 0.05 | 0.05 |
| C12 | 節点の並び替え<br>`RCM_POLICY` | AUTO | AUTO |
| C15 | Fs探索の上限<br>`SRM_FMAX` | 3 | モデル依存 |
| C16 | 成功・失敗Fsの許容幅<br>`SRM_TOL` | 0.02 | 現値を維持 |
| C17 | ψの扱い<br>`SRM_PSI_POLICY` | KEEP | モデル依存 |
| C18 | モード<br>`SRM_MODE` | REAPPLY | REAPPLY |
| C19 | 固定Fsで試験する<br>`SRM_FIXED_FS` | 0 | 0 |

## 高速化設定

| 入力セル | 項目 / KEY | 配布保存値 | 推奨・目安 |
| --- | --- | --- | --- |
| C9 | DLLの使用<br>`NATIVE_DLL` | AUTO | AUTO |
| C10 | DLLの演算方式（倍精度）<br>`NATIVE_PRECISION` | AUTO | AUTO |
| C11 | 状態配列のDLL転送<br>`NATIVE_STATE_COPY` | 1 | 1 |
| C12 | 行列組立のDLL処理<br>`NATIVE_ASSEMBLY` | 1 | 1 |
| C13 | GMRESのCPU拡張命令<br>`NATIVE_GMRES_SIMD` | 1 | 1 |
| C14 | GMRESの転送・更新最適化<br>`NATIVE_GMRES_TUNED` | 1 | 1 |
| C15 | GMRESの作業領域再利用<br>`NATIVE_GMRES_PREPARED` | 1 | 1 |
| C16 | ログの整形・バッファ処理<br>`IO_FAST_PATH` | 1 | 1 |
| C17 | 診断イベントの整形処理<br>`EVENT_LOG_FAST` | 1 | 1 |
| C20 | 候補方式の自動選別<br>`ACCEL_ADAPTIVE` | 1 | 1 |
| C21 | V1 変位予測<br>`ACCEL_V1_PREDICTOR` | 1 | 1 |
| C22 | V2a 増分間の接線再利用<br>`ACCEL_V2A_REUSE` | 1 | 1 |
| C23 | V2b 費用による接線更新<br>`ACCEL_V2B_COST` | 0 | 0 |
| C24 | V3 Fs探索の調整<br>`ACCEL_V3_COST_SEARCH` | 0 | 0 |
| C25 | V4 Anderson補正<br>`ACCEL_V4_ANDERSON` | 1 | 1 |
| C26 | V5 旧LUを使うGMRES<br>`ACCEL_V5_GMRES_LU` | 0 | 0 |
| C27 | 小さくなった増分幅の回復<br>`ACCEL_STEP_RECOVERY` | 0 | 0 |
| C28 | 回復直後のfresh LU試験<br>`ACCEL_STEP_FRESH_LU` | 0 | 0 |
| C29 | 高速化詳細ログ<br>`ACCEL_TRACE` | 0 | 0 |
| C30 | 予測子係数beta<br>`ACCEL_PREDICT_BETA` | 1 | 1 |
| C31 | 予測子刻み比上限<br>`ACCEL_PREDICT_MAX_RATIO` | 1.5 | 1.5 |
| C32 | 接線変化の上限<br>`ACCEL_TANGENT_CHANGE` | 0.05 | 0.05 |

## 出力・表示

| 入力セル | 項目 / KEY | 配布保存値 | 推奨・目安 |
| --- | --- | --- | --- |
| C9 | 結果の残し方<br>`OUTPUT_STAGE_MODE` | FINAL | FINAL |
| C10 | 無効要素<br>`OUTPUT_INACTIVE_ELEMENTS` | SKIP | SKIP |
| C11 | ログを保存する範囲<br>`DEBUG_MODE` | STAGE | STAGE |
| C12 | ログ書き出し間隔<br>`DEBUG_FLUSH` | 5 | 5 |
| C13 | 途中状態を保存する範囲<br>`EXPORT_MODE` | STAGE | STAGE |
| C14 | 途中結果から再開する<br>`EXPORT_LOAD` | （空欄） | 空欄（新規） |
| C17 | 変位倍率<br>`VIEW_DISP_SCALE` | 1 | モデル・表示による |
| C18 | 表示する結果の番号<br>`VIEW_COLOR_RESULT` | 13 | モデル・表示による |
| C19 | 色スケール<br>`VIEW_SCALE_MODE` | SYMMETRIC | SYMMETRIC |
| C20 | 色付けする結果の集め方<br>`VIEW_RESULT_SOURCE` | SMOOTHED | SMOOTHED |
| C21 | 表示モード<br>`VIEW_DETAIL_MODE` | DETAIL | SIMPLE |
| C22 | スケール<br>`VIEW_SCALE` | 35 | モデル・表示による |
| C23 | 原点X<br>`VIEW_MIN_X` | 1000 | モデル・表示による |
| C24 | 原点Y<br>`VIEW_MAX_Y` | 1000 | モデル・表示による |

## 詳細設定

| 入力セル | 項目 / KEY | 配布保存値 | 推奨・目安 |
| --- | --- | --- | --- |
| C9 | 帯域ソルバ<br>`SOLVER` | BAND_LU | BAND_LU |
| C10 | Band LUカーネル<br>`BAND_LU_KERNEL` | NEW | NEW |
| C11 | 物理破壊監視<br>`PHYSICAL_FAILURE_MODE` | SHADOW | 固定・未使用 |
| C14 | 最小面積比<br>`MESH_QUALITY_MIN_AREA_RATIO` | 1e-06 | 1e-06 |
| C15 | 最小Jacobian比<br>`MESH_QUALITY_MIN_JACOBIAN_RATIO` | 1e-06 | 1e-06 |
| C16 | 最小角度<br>`MESH_QUALITY_MIN_ANGLE_DEG` | 45 | 品質目標による |
| C17 | 最大角度<br>`MESH_QUALITY_MAX_ANGLE_DEG` | 135 | 品質目標による |
| C18 | 最大アスペクト<br>`MESH_QUALITY_MAX_ASPECT_RATIO` | 5 | 5 |
| C21 | 最大反復<br>`MESH_REPAIR_MAX_ITERATIONS` | 5 | 5 |
| C22 | 移動係数<br>`MESH_REPAIR_DAMPING` | 0.5 | 0.5 |
| C23 | 改善判定<br>`MESH_REPAIR_IMPROVEMENT_TOLERANCE` | 0.0001 | 0.0001 |
| C24 | 角度最適化反復<br>`MESH_REPAIR_ANGLE_ITERATIONS` | 2000 | 2000 |
| C25 | 地形追従比較<br>`MESH_REPAIR_TERRAIN_COMPARE` | 1 | 1 |
| C26 | 格子要素数下限比<br>`MESH_TERRAIN_COUNT_MIN` | 0.55 | 0.55 |
| C27 | 格子要素数上限比<br>`MESH_TERRAIN_COUNT_MAX` | 1.8 | 1.8 |
| C28 | 要素数倍率上限<br>`ADAPT_MAX_ELEMENT_RATIO` | 2 | 2 |
| C29 | 細分化割合<br>`ADAPT_MARK_FRACTION` | 0.15 | 0.15 |
| C32 | MESH_GEN_LAST_METHOD<br>`MESH_GEN_LAST_METHOD` | （自動出力） | 自動出力 |
| C33 | 生成比較の結果（自動更新）<br>`MESH_GEN_LAST_COMPARE` | （自動出力） | 自動出力 |
| C34 | BAND_LU_BENCH<br>`BAND_LU_BENCH` | 0 | モデル・表示による |
| C35 | BAND_LU_BENCH_N<br>`BAND_LU_BENCH_N` | 20 | モデル・表示による |
