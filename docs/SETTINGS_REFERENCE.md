# 設定一覧とKEY対応

[READMEへ](../README.md) · [操作説明書](USER_GUIDE.md) · [入力仕様](MODEL_INPUTS.md) · [結果・対処法](RESULTS_AND_TROUBLESHOOTING.md)

対象は `2DSoilFEM_20261008_practical.xlsm`（`20261008_BEST_03`）です。以下の74項目の「配布保存値」は配布XLSMの用途別シートC列から確認した値です。プログラムがKEY欠落時に使う既定値や、すべての案件に推奨する値とは区別してください。

## 目次

- [設定の保存場所とKEY](#設定の保存場所とkey)
- [旧C37〜C46の対応](#旧c37c46の対応)
- [メッシュ設定](#メッシュ設定)
- [解析設定](#解析設定)
- [高速化設定](#高速化設定)
- [出力・表示](#出力表示)
- [詳細設定](#詳細設定)
- [設定が入らないとき](#設定が入らないとき)

## 設定の保存場所とKEY

編集する場所は「メッシュ設定」「解析設定」「高速化設定」「出力・表示」「詳細設定」の **C列** です。隠れているF列に設定KEYがあります。内部の旧「設定」シートはE列にKEY、C列に用途別シートを読む参照式を持ち、計算部はこの対応を通じて値を取得します。[設定連携](../src/vba/FEMUi.bas#L91)

旧設定C列へ値を直接入力する運用は避けてください。連携を再構築すると参照式へ戻ります。設定を特定するときはセル番号だけでなくKEYも確認します。旧版と現在版では、同じC37等でも用途が異なる場合があります。

0/1項目には数値の0または1、文字の選択項目には表示された選択肢を入力します。高速化の1を文字ONで代用しないでください。入力リストと数値制限は用途別セルへ設定されています。保護は見出し・KEY・数式等を守り、KEYのあるC列の入力欄を編集可能にしています。

## 旧C37〜C46の対応

旧「設定」の該当範囲を使っていた場合は、現在は次の入力セルを使います。この対応は配布版のKEY照合によるものです。

|旧設定セル|KEY|現在の入力先|入力|
|---|---|---|---|
|C37|`ACCEL_V2A_REUSE`|高速化設定 C13|0 / 1|
|C38|`ACCEL_V2B_COST`|高速化設定 C14|0 / 1|
|C39|`ACCEL_V3_COST_SEARCH`|高速化設定 C15|0 / 1|
|C40|`ACCEL_V4_ANDERSON`|高速化設定 C16|0 / 1|
|C41|`ACCEL_V5_GMRES_LU`|高速化設定 C17|0 / 1|
|C42|`ACCEL_TRACE`|高速化設定 C24|0 / 1|
|C43|`ACCEL_PREDICT_BETA`|高速化設定 C25|0以上1以下|
|C44|`ACCEL_PREDICT_MAX_RATIO`|高速化設定 C26|0より大きく2以下|
|C45|`ACCEL_TANGENT_CHANGE`|高速化設定 C27|0より大きく0.1以下|
|C46|`SRM_FIXED_FS`|解析設定 C19|0=探索、正数=固定Fs|

数値の3項目はドロップダウンで0/1を選ぶ設定ではなく、範囲を持つ数値入力です。固定Fsも非負の数値です。V1、適応選別、増分幅回復はこの旧10セルの外にあります。

## メッシュ設定

|入力セル|項目|KEY|配布保存値|選択肢・条件|用途|
|---|---|---|---|---|---|
|C9|生成要素の材料番号|`MESH_MATERIAL`|1|数値（用途欄参照）|要素作成①②で付ける|
|C10|底面拘束|`MESH_BC_BOTTOM`|PINNED|NONE / ROLLER / PINNED|Ymin面。NONE=自由 / ROLLER=Y固定 / PINNED=XY固定|
|C11|左面拘束|`MESH_BC_LEFT`|ROLLER|NONE / ROLLER / PINNED|Xmin面。NONE=自由 / ROLLER=X固定 / PINNED=XY固定|
|C12|右面拘束|`MESH_BC_RIGHT`|ROLLER|NONE / ROLLER / PINNED|Xmax面。NONE=自由 / ROLLER=X固定 / PINNED=XY固定|
|C13|上面拘束|`MESH_BC_TOP`|NONE|NONE / ROLLER / PINNED|Ymax面。NONE=自由 / ROLLER=Y固定 / PINNED=XY固定|
|C14|剛体モード止め|`MESH_BC_PIN_CORNER`|AUTO|AUTO / NONE|AUTO=底面の最小X節点をXY固定 / NONE=しない|
|C17|点1 X|`MESH_POINT1_X`|0|数値（用途欄参照）|左下X（要素作成①）|
|C18|点1 Y|`MESH_POINT1_Y`|0|数値（用途欄参照）|左下Y|
|C19|点2 X|`MESH_POINT2_X`|0|数値（用途欄参照）|左上X|
|C20|点2 Y|`MESH_POINT2_Y`|4|数値（用途欄参照）|左上Y|
|C21|点3 X|`MESH_POINT3_X`|3|数値（用途欄参照）|右上X|
|C22|点3 Y|`MESH_POINT3_Y`|4|数値（用途欄参照）|右上Y|
|C23|点4 X|`MESH_POINT4_X`|3|数値（用途欄参照）|右下X|
|C24|点4 Y|`MESH_POINT4_Y`|0|数値（用途欄参照）|右下Y|
|C25|横方向の分割数 nx|`MESH_NX`|25|数値（用途欄参照）|X方向分割|
|C26|縦方向の分割数 ny|`MESH_NY`|25|数値（用途欄参照）|Y方向分割|
|C29|形状メッシュの生成方式|`MESH_GEN_COMPARE`|2|0 / 1 / 2|0=両方合格なら品質の良い方 / 1=デローニ修正優先 / 2=地形格子（生格子。角度最適化しない）|

ROLLERは底面・上面ではY固定、左右ではX固定です。PINNEDはXY固定、NONEは拘束追加なしです。AUTOの剛体モード止めは底面の最小X節点をXY固定します。生成後の個別拘束は節点データで確認します。点座標・X/Y分割数は簡易メッシュ①に使い、形状②は形状入力の境界・格子間隔を使います。

## 解析設定

|入力セル|項目|KEY|配布保存値|選択肢・条件|用途|
|---|---|---|---|---|---|
|C9|自重|`WEIGHT_MODE`|TOTAL|TOTAL|TOTALのみ。全応力のγ。浮力・水位は未実装|
|C10|非関連の扱い|`FLOW_POLICY`|DAVIS|INCONSISTENT / DAVIS|INCONSISTENT=流れ則は非関連・接線は対称近似 / DAVIS=等価関連c*,φ*（JOINTには掛けない）|
|C11|アワーグラス係数|`Q8_HOURGLASS_FACTOR`|0.05|数値（用途欄参照）|0～1|
|C12|節点の並び替え|`RCM_POLICY`|AUTO|AUTO / ON / OFF|AUTO=帯幅が減るときだけ並び替え / ON=必ずRCM / OFF=入力順|
|C15|Fs探索の上限|`SRM_FMAX`|3|数値（用途欄参照）|探索の上限|
|C16|成功・失敗Fsの許容幅|`SRM_TOL`|0.02|数値（用途欄参照）|PASSとFAILのFs幅の終了許容値。配布保存値0.02。|
|C17|ψの扱い|`SRM_PSI_POLICY`|KEEP|CAP / REDUCE / KEEP|CAP=ψをφ'以下に制限 / REDUCE=ψもFsで低減 / KEEP=入力のまま|
|C18|モード|`SRM_MODE`|REAPPLY|REAPPLY|REAPPLYのみ。各Fs試行で先行ステージを再実行|
|C19|固定Fsで試験する|`SRM_FIXED_FS`|0|0=探索／正数=固定Fs|0=全探索 / 正=そのFsだけ独立再載荷。F<1も可|

INCONSISTENTとDAVIS、ψのCAP/REDUCE/KEEPは計算の材料条件に関わります。単なる速度オプションとして結果比較の途中で変更しないでください。DAVISの等価関連処理はJOINTには掛かりません。RCMは内部の帯幅縮小用です。SRM_FIXED_FSを正数にした実行は全探索の計測と分けて記録します。

## 高速化設定

|入力セル|項目|KEY|配布保存値|選択肢・条件|用途|
|---|---|---|---|---|---|
|C9|候補方式の自動選別|`ACCEL_ADAPTIVE`|1|0 / 1|0=固定フラグ試験。1=ONの候補を適格性・進行・採算で選別。|
|C12|V1 変位予測|`ACCEL_V1_PREDICTOR`|1|0 / 1|0=基準 / 1=同一ステージの確定変位増分から予測|
|C13|V2a 増分間の接線再利用|`ACCEL_V2A_REUSE`|1|0 / 1|0=基準 / 1=同一塑性集合・小接線変化で最初の補正に旧LU|
|C14|V2b 費用による接線更新|`ACCEL_V2B_COST`|0|0 / 1|0=基準 / 1=実測費用と改善実績で限定再利用。上限3は維持|
|C15|V3 Fs探索の調整|`ACCEL_V3_COST_SEARCH`|0|0 / 1|0=二分 / 1=保護付き探索。最終実幅は基準二分以下|
|C16|V4 Anderson補正|`ACCEL_V4_ANDERSON`|1|0 / 1|0=基準 / 1=凍結接線・未減衰区間のみ。悪化候補は棄却|
|C17|V5 旧LUを使うGMRES|`ACCEL_V5_GMRES_LU`|0|0 / 1|0=Band LU / 1=最新帯行列＋旧LU。真の残差を確認しLUへ復帰|
|C20|小さくなった増分幅の回復|`ACCEL_STEP_RECOVERY`|0|0 / 1|追加の増分幅回復試験。配布値0。適応選別が1の場合の試験用。|
|C21|回復直後のfresh LU試験|`ACCEL_STEP_FRESH_LU`|0|0 / 1|追加回復試験中に最初の補正をfresh LUにする。回復・適応選別をONにした試験用。|
|C24|高速化詳細ログ|`ACCEL_TRACE`|0|0 / 1|0=追加詳細OFF / 1=正確な更新理由・全増分・候補/復帰をsolver_events.csvへ|
|C25|予測子係数beta|`ACCEL_PREDICT_BETA`|1|0以上1以下|検証用パラメータ 0～1。収束条件は変更しない|
|C26|予測子刻み比上限|`ACCEL_PREDICT_MAX_RATIO`|1.5|0より大きく2以下|検証用パラメータ 0超～2。上限超は予測しない|
|C27|接線変化の上限|`ACCEL_TANGENT_CHANGE`|0.05|0より大きく0.1以下|検証用パラメータ 0超～0.1。要素ごとの相対Frobenius差|

配布値は実測ADAPT_03の選別方式を使う候補構成です。すべてのモデルで最速を保証する組合せではありません。V1とV2aは適用条件を別に評価します。V2b/V3/V5および追加回復試験は配布値OFFです。比較条件と実測の評価は [高速化履歴](ACCELERATION_HISTORY.md) にあります。

## 出力・表示

|入力セル|項目|KEY|配布保存値|選択肢・条件|用途|
|---|---|---|---|---|---|
|C9|結果の残し方|`OUTPUT_STAGE_MODE`|FINAL|FINAL / ALL|FINAL=最後だけ結果シートへ / ALL=各ステージを追記|
|C10|無効要素|`OUTPUT_INACTIVE_ELEMENTS`|SKIP|SKIP / INCLUDE|SKIP=Death中は出さない / INCLUDE=行は残しステージ番号だけ書く|
|C11|ログを保存する範囲|`DEBUG_MODE`|STAGE|OFF / STAGE / ITER|OFF=書かない / STAGE=ステージ境界 / ITER=反復も。*_out/run.log|
|C12|ログ書き出し間隔|`DEBUG_FLUSH`|5|数値（用途欄参照）|秒。長時間計算の監視用|
|C13|途中状態を保存する範囲|`EXPORT_MODE`|STAGE|OFF / STAGE / ON_FAIL|OFF=保存しない / STAGE=成功ステージ / ON_FAIL=失敗時も。*_out へ|
|C14|途中結果から再開する|`EXPORT_LOAD`|空欄|空欄／AUTO／保存フォルダー名|空=最初から / AUTO=最後のSUCCESS / stg01_GRAVITY などフォルダ名|
|C17|変位倍率|`VIEW_DISP_SCALE`|1|数値（用途欄参照）|変形図|
|C18|表示する結果の番号|`VIEW_COLOR_RESULT`|13|1〜68の整数|現行の保存結果ビューアの選択番号。結果要素E:BTの積分点別列。|
|C19|色スケール|`VIEW_SCALE_MODE`|SYMMETRIC|AUTO / SYMMETRIC / USER|AUTO=選択値の範囲、SYMMETRIC=正負対称。USERは縮尺・作図原点の手入力。|
|C20|色付けする結果の集め方|`VIEW_RESULT_SOURCE`|SMOOTHED|SMOOTHED / ELEMENT / RAW_GAUSS|従来P1処理の結果源。現在のFEMViewResultは選択した保存積分点列を読む。|
|C21|表示モード|`VIEW_DETAIL_MODE`|DETAIL|SIMPLE / DETAIL|SIMPLE=簡易 / DETAIL=節点番号など詳細|
|C22|スケール|`VIEW_SCALE`|35|数値（用途欄参照）|座標→図形|
|C23|原点X|`VIEW_MIN_X`|1000|数値（用途欄参照）|作図原点X|
|C24|原点Y|`VIEW_MAX_Y`|1000|数値（用途欄参照）|作図原点Y|

DEBUG/EXPORTと図の選択番号は [結果・ログ説明](RESULTS_AND_TROUBLESHOOTING.md) を参照してください。現在の保存結果ビューアでUSERが手入力するのは図の縮尺・原点です。色の上下限入力欄はありません。

## 詳細設定

|入力セル|項目|KEY|配布保存値|選択肢・条件|用途|
|---|---|---|---|---|---|
|C9|帯域ソルバ|`SOLVER`|BAND_LU|BAND_LU / BAND_LDLT_EXPERIMENTAL|BAND_LU=標準 / BAND_LDLT_EXPERIMENTAL=影のLDLTをサンプル比較（本番はLUのまま）|
|C10|Band LUカーネル|`BAND_LU_KERNEL`|NEW|NEW / OLD|NEW=パックド1D（本番） / OLD=従来2D（回帰）。ログは BandLUKernel=|
|C11|物理破壊監視|`PHYSICAL_FAILURE_MODE`|SHADOW|SRMはSHADOW監視|通常操作の切替用ではない。GRAVITY監視なし、SRMの力学診断はSHADOW。|
|C14|最小面積比|`MESH_QUALITY_MIN_AREA_RATIO`|1e-06|数値（用途欄参照）|代表寸法^2に対する比|
|C15|最小Jacobian比|`MESH_QUALITY_MIN_JACOBIAN_RATIO`|1e-06|数値（用途欄参照）|代表寸法^2に対する比|
|C16|最小角度|`MESH_QUALITY_MIN_ANGLE_DEG`|45|数値（用途欄参照）|deg|
|C17|最大角度|`MESH_QUALITY_MAX_ANGLE_DEG`|135|数値（用途欄参照）|deg|
|C18|最大アスペクト|`MESH_QUALITY_MAX_ASPECT_RATIO`|5|数値（用途欄参照）|長辺/短辺|
|C21|最大反復|`MESH_REPAIR_MAX_ITERATIONS`|5|数値（用途欄参照）|1～50|
|C22|移動係数|`MESH_REPAIR_DAMPING`|0.5|数値（用途欄参照）|0～1|
|C23|改善判定|`MESH_REPAIR_IMPROVEMENT_TOLERANCE`|0.0001|数値（用途欄参照）|相対|
|C24|角度最適化反復|`MESH_REPAIR_ANGLE_ITERATIONS`|2000|数値（用途欄参照）|1～5000|
|C25|地形追従比較|`MESH_REPAIR_TERRAIN_COMPARE`|1|0 / 1|1=地形追従と角度最適化を比較 / 0=しない|
|C26|格子要素数下限比|`MESH_TERRAIN_COUNT_MIN`|0.55|数値（用途欄参照）|デローニ要素数に対する比|
|C27|格子要素数上限比|`MESH_TERRAIN_COUNT_MAX`|1.8|数値（用途欄参照）|デローニ要素数に対する比|
|C28|要素数倍率上限|`ADAPT_MAX_ELEMENT_RATIO`|2|数値（用途欄参照）|再メッシュの増加上限|
|C29|細分化割合|`ADAPT_MARK_FRACTION`|0.15|数値（用途欄参照）|0～1|
|C32|MESH_GEN_LAST_METHOD|`MESH_GEN_LAST_METHOD`|地形追従格子|処理記録|最後のメッシュ生成方法の処理記録。通常は編集不要。|
|C33|生成比較の結果（自動更新）|`MESH_GEN_LAST_COMPARE`|候補比較の記録（複数行）|処理記録|最後の候補比較の処理記録。本文は配布ブックC33に保存。|
|C34|BAND_LU_BENCH|`BAND_LU_BENCH`|0|通常0（比較試験用）|帯LU比較試験用。通常は0。1のベンチ専用実行では自重後にSRMを実行しない。|
|C35|BAND_LU_BENCH_N|`BAND_LU_BENCH_N`|20|数値（用途欄参照）|帯LU比較試験の反復回数設定。通常利用では変更不要。|

BAND_LUが通常の本番ソルバーです。BAND_LDLT_EXPERIMENTALは接線LDLTの比較を行う実験設定で、本番のLU解法を置き換える指定ではありません。品質目標はメッシュ生成・修正条件であり、目標に合格しただけで解析精度が保証されるものではありません。

## 設定が入らないとき

1. 現在の用途別シートC列を編集しているか確認します。旧「設定」C列は連携用の式です。
2. 0/1の項目には数値を入力し、文字選択には候補の綴りを使います。数値欄の先頭にアポストロフィを付けて文字にしないでください。
3. KEY・見出し・シート名や、C列以外の式を貼り付けで上書きしていないか確認します。表全体の貼り付けではなく値の入力欄を選びます。
4. マクロを有効にして開き直します。起動時に入力連携・ボタン等が構築されます。破損したKEYを手で推測して直すより、配布ブックのコピーと入力を比較します。
5. 設定を変更した解析は再実行します。旧結果が表に残っていても現在の設定の結果ではありません。

設定の読込側は [CONTROL.bas](../src/vba/CONTROL.bas#L91)、一覧・入力制限は [FEMIo.bas](../src/vba/FEMIo.bas#L3320)、入力連携は [FEMUi.bas](../src/vba/FEMUi.bas#L91) にあります。保存ブックの74KEY、リスト・数値制限・保護の確認は [動作確認報告](VALIDATION.md) に記録しています。
