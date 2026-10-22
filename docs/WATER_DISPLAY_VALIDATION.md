# 水理表示・工程別入力の更新と検証

水理表示と入力の更新識別は `20261022_WATER_VIEW_03`、配布ブックは `2DSoilFEM_ui_20261023.xlsm` です。計算部の識別 `20261022_WATER_02`、Double、材料則、収束条件、SRM探索条件、DLLは変更していません。地形格子・デローニ境界辞書の更新も保持しています。

## 入力と表示

- 水条件の共通入力へモード説明と移動リンクを追加しました。
- 既存ステージの工程IDに対応する L:T 表を追加しました。L:N は現行入力、O:S は将来の時間圧密用で灰色です。現行版では予約欄への入力や FIXED_STRESS を検証エラーにします。[入力仕様と将来拡張](WATER_STAGE_INPUT_PLAN.md)。
- 橙破線は登録入力水位、紫菱形は固定水頭節点、紫実線は計算全水頭の等値線、緑矢印は計算Darcy流束、青点はGP水圧、橙矢印は力学への等価節点水圧荷重です。HYDROSTATICには計算水頭・流束を描きません。
- 等値線はQ8の保存節点水頭を表示用に三角形へ補間します。ソルバーや数値表を変更しません。非透水節点の未計算水頭0は描画対象外です。一定水頭では丸め誤差程度の流束を大きな流向矢印へ正規化しません。未知自由水面の推定線ではありません。
- 水頭範囲と凡例の重なりを実画面確認で修正しました。既存のExcel標準図形による描画方式を維持しています。

## 実測

同一入力、同一DLLで計算し、[全成分比較](water-display-evidence/20261022/water_display_check.json)を行いました。

|例|更新前→更新後の時間 [s]|全変位・応力・診断値の最大差|
|---|---:|---:|
|遮水構造体cutoff|5.891 → 5.374|0|
|内水ありbox|4.493 → 5.145|0|
|水位変更・掘削・再施工stages|6.622 → 6.033|0|
|湿潤斜面INCONSISTENT SRM|137.598 → 146.785|0|
|湿潤斜面DAVIS SRM|20.419 → 21.390|0|

時間は各1回の壁時計実測で、表示・I/Oを含み、速度改善の判定ではありません。INCONSISTENTと最初の3例の基準は地下水受入20、DAVISの基準は今回の表示位置修正前の同条件実行です。DAVISとINCONSISTENTを同じ解とみなしていません。

湿潤斜面のFs区間はINCONSISTENT 1.5500–1.5625、DAVIS 1.5000–1.5125、最終自由DOF残差はそれぞれ2.483e-14、1.117e-10です。成功Fsでの最大節点変位ノルムは1.43819 m、6.72372 mと大きいため、PDFには元の座標への変位着色図を載せます。小ひずみモデルの変形予測として採用する図ではありません。別途固定Fs=1.3も実計算し、変形倍率1の例を保存しました（最大変位0.08382 m、残差1.317e-8、5.896 s）。

工程別N列で水位線を指定した `one_way` と、従来C列指定の `stages` は全変位・全応力・診断値が完全一致しました。保存再開も照合しました。内水ありboxの水圧合力85 kN・モーメント127.5 kN m・反力107 kN、JOINTの有効接触力28 kN＋水圧支持力20 kN＝48 kN、鉛直反力88 kNを確認しました。

## 確認範囲と再現

Excel 16.0 build 20430でVBEコンパイル、保存・再読込、全33モジュールの埋込みソース照合、配布元入力の保持を確認しました。これは現在の実行環境の確認であり、Excel 2016実機での新たな再試験を行ったという意味ではありません。

乾燥・静水圧・水没・全応力・定常浸透の既存試験、遮水体、空洞内水、工程、JOINTと接線、両流れ則SRMを実行しました。FIXED_STRESS、予約時間入力、工程番号未入力、C/N矛盾、孤立ID、重複ID、未登録水位線の7入力拒否も確認しました。[独立Darcy・沈下検算](water-display-evidence/20261022/independent.json)／[JOINT検算](water-display-evidence/20261022/joint_check.json)。

専用コピーで共通入力・工程別表・施工順・プレビュー・結果・解析前メッシュ・応力図・保存再表示を実画面で確認し、入力変更後の旧結果拒否を確認しました。入力変更拒否の最初のUI試験はテスト側のVBA例外受け取り方法で停止したため、受け取り用関数を修正し、専用コピーで再試験して期待したエラーを確認しています。配布ブックの計算不具合とは区別しています。

```powershell
pwsh -NoProfile -File tests/run_water_display_regression.ps1 -Book workbook/ui_20261023/2DSoilFEM_ui_20261023.xlsm -Out tests/tmp/water_view_rerun
python tests/check_groundwater.py --basic tests/tmp/water_view_rerun/basic --cutoff tests/tmp/water_view_rerun/cutoff
python tests/check_joint_water.py --fixture tests/fixtures/joint_water_wet.json --run tests/tmp/water_view_rerun/joint --report tests/tmp/water_view_rerun/joint_check.json
```

Excelが起動中なら開始しません。長時間SRMの計算を時間で打ち切る試験ではありません。[計算済み例題ブック集](manuals/examples/water_display_examples.zip)／[生CSV・ログ](water-display-evidence/20261022/calculation_logs.zip)／[コンパイル・再開照合](water-display-evidence/20261022/package.json)／[ソース照合](water-display-evidence/20261022/sources.json)。自由水面・不飽和・時間圧密・fixed-stress連成は未実装です。
