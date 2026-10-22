# 1024版説明書の確認

対象は [操作説明書PDF](manuals/2DSoilFEM_Manual_20261024.pdf) と [原稿](manuals/2DSoilFEM_Manual_20261024.md)。対応ブックはmainの [2DSoilFEM_ui_20261024.xlsm](../workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm)（PR #41、merge `5581c061af2983a38eacd69fce3cf3d260a8e352`）です。

## 掲載範囲と操作

1023版の全31章と解析例1～18・11-Bの19節を保持しています。各例の「確認値」表131行は原稿間で完全一致します。理論図、保存された実計算の結果図、擁壁断面力、支持力、トンネル、アンカー相当力、JOINT、水理、折れ線水位、湿潤斜面SRM、出力整理、巻末用語集を継承しました。操作は7画面への移動、C4での切替、青い入力欄、工程編集、表示設定を説明します。元表が必要な操作は詳細表への経路を明示しています。

実Excelの専用コピーで76画面を確認しました。掲載用画像は69点（実画面の切り出し・結合と実測値の整理表）で、入力欄に赤枠と番号を付けています。画像の撮影元・SHA-256・入力値は [screen_checks.json](manual24-evidence/20261024/screen_checks.json)、撮影順は [capture_plan.json](manual24-evidence/20261024/capture_plan.json)、赤枠番号は [image_annotations.json](manual24-evidence/20261024/image_annotations.json) に保存しました。既存の計算済みブックは読み取り専用で開き、入力値と保存結果を撮影コピーへ移しています。図の表示倍率1/50、塑性状態選択、工程の日本語表示、水位線・材料水理特性・水理境界を実画面と照合しています。

## 基本斜面の実計算と独立検算

配布1024ブックそのものを元にした専用コピーを、Excel 64bit 16.0 build 20430で計算しました。261節点・72 Q8要素、c=10 kPa、φ=30°、ψ=0°、INCONSISTENT、自重→SRM、FMAX=3、SRM_TOL=0.02です。倍精度、材料則、収束条件、SRM条件は変更していません。

|項目|結果|
|---|---:|
|成否|自重・SRMともPASS|
|Fs成功側／失敗側|2.6250000000000013／2.637500000000001|
|Fs区間幅|0.012499999999999734|
|計算時間|72.8523451秒|
|最終自由DOF相対残差|1.2822499905154632×10⁻¹⁰|
|鉛直支持反力|1187.999999999991 kN|
|独立積分した自重|1188.0 kN（66 m²×18 kN/m³×1 m）|
|支持反力＋自由DOF残差－自重|−2.2737367544323206×10⁻¹³ kN|
|最終変位ノルムの最大値|0.48249146758059497 m|
|拘束自由度の最大変位|0 m|
|全要素11×11点の最小Jacobian（元形状／倍率1）|0.14999999999999647／0.12098280848750796|

Fs区間・自重反力は従来の基本例と同じです。時間はこの実行の実測値で、速度改善の比較試験ではありません。変位ノルムと工程表の「試行中の最大変位成分」は定義が異なります。倍率1の図には反転・潰れがありません。倍率50の図は、過大な表示倍率の影響を説明する図として明示しています。

[入力・全変位/応力・節点表・生ログ・計算済みブック](manual24-evidence/20261024/README.md)。既存の応用例の数値結果は保存記録を保持し、全例を1024版で再計算したとは扱いません。

## PDFの検査

全212ページを画像化し、全ページの配置を確認しました。目次31項目・しおり・ページ番号・図番号・参照範囲・用語集のページ参照を検査し、変更された操作ページと主要図を拡大確認しています。図のキャプションと画像の分離、ページ外の文字、空白ページはありません。水条件の指定優先順は、工程H列（元の水条件N列）→工程C列→既定値B10（元の水条件B5）で一致し、H/Cの矛盾は停止します。

PDF中の配布ブックと検証資料のファイルリンク11件は、すべてmainのGitオブジェクト存在を確認しました。作者PCのfile:リンク、機能ブランチや特定コミットの配布先はありません。表紙の斜面入力メッシュと既存の書式を保ち、用語集は巻末です。作業指示・依頼の確認文は本文へ追加していません。

[構造・全ページ描画検査](manual24-evidence/20261024/pdf_checks.json)／[リンク・ページ内配置](manual24-evidence/20261024/pdf_link_bounds_checks.json)／[目視確認と最終ハッシュ](manual24-evidence/20261024/visual_review.json)／[掲載範囲の一致](manual24-evidence/20261024/coverage.json)。

## 再現

リポジトリ直下で実行します。

```powershell
$env:MANUAL_PUBLIC_REF='main'
$env:MANUAL_PDF_TITLE='2DSoilFEM 操作説明書 20261024'
python -X utf8 docs/manuals/build_manual_pdf_v2.py docs/manuals/2DSoilFEM_Manual_20261024.md docs/manuals/2DSoilFEM_Manual_20261024.pdf
python -X utf8 tests/check_manual24_coverage.py
python -X utf8 tests/review_water_manual.py docs/manuals/2DSoilFEM_Manual_20261024.pdf tests/tmp/manual24/review --source docs/manuals/2DSoilFEM_Manual_20261024.md --public-ref main --book-version 20261024
python -X utf8 tests/check_manual24_pdf.py
```

Chrome、Pythonのmarkdown・PyMuPDF・Pillow・NumPyを使用します。実計算の再現は受入記録の手順を参照してください。XLSM/VBA/DLLと1023版の原稿・PDFは、この説明書更新では変更していません。
