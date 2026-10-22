# 1024版説明書の受入記録

- [全検証の説明](../../MANUAL_20261024_VALIDATION.md)
- [基本斜面の計算済みブック・DLL1個・入力JSON](../../manuals/examples/manual24_basic.zip)
- [入力JSON](../../../tests/fixtures/manual24_basic.json)
- [全変位・全応力・診断値・実行時間](basic_result.json)
- [節点の保存結果](basic_nodes.json)／[工程結果](basic_stages.json)
- [計算の全ログ・工程チェックポイント](basic_calculation_logs.zip)
- [独立検算](independent_checks.json)
- [76画面の受入値](screen_checks.json)／[撮影順](capture_plan.json)／[赤枠番号](image_annotations.json)
- [本文掲載範囲](coverage.json)／[PDF検査](pdf_checks.json)／[リンク](pdf_link_bounds_checks.json)／[目視確認](visual_review.json)

## 基本例の再現

Excelが終了した状態で、リポジトリ直下から専用コピーを実行します。出力先は新しいフォルダーを指定します。

```powershell
./tests/run_groundwater_fixture.ps1 -Fixture tests/fixtures/manual24_basic.json -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -Out tests/tmp/manual24_reproduce -Flow INCONSISTENT
python -X utf8 tests/check_manual24_basic.py tests/fixtures/manual24_basic.json tests/tmp/manual24_reproduce tests/tmp/manual24_reproduce/independent_checks.json
```

配布原版は変更せず、倍精度・収束条件も保持します。既存の出力フォルダーへ重ねません。`basic_result.json`の計算対象ブックSHA-256は `16161736280992174173430F9A163FCA72CC4F3D408683CDCE1990461910B560` です。計算済み基本例ブックのSHA-256は `C4FE7B99B1103A8C42B2D2CCD09591DBAAB99CD6660AE3D3942AA5D02B545604` です。

## 画面の再現

`tests/capture_manual24.ps1` は配布ブックの専用コピーを作り、ホームで初期化した後、`command.json`の指定で保存済み例題の入力値・結果を表示します。元の例題は読み取り専用です。Excelのユーザー作業中は起動しません。`capture_plan.json`のcaseパスはリポジトリ直下基準です。水理の例題はmainの`water_display_examples.zip`と`water_polyline_example.zip`を、計画の`tests/tmp/manual24/cases/`配下へ展開します。基本例は上の実行、または基本例ZIPを`tests/tmp/manual24/basic/`へ展開して用意します。

各指定が`status.txt`へ出力された後、Computer Useの実ウィンドウキャプチャを`raw/名前.jpg`として保存します。最後に`{"action":"quit"}`を送ると専用コピーだけを閉じます。`tests/prepare_manual24_images.py --base tests/tmp/manual24`で切り出しと赤枠を再現できます。`capture_plan.json`をbase直下へ置きます。入力値の記録は数式の再実装ではなく、実ExcelのValue2です。図の倍率変更には元設定表の再結線・設定変更通知を実行します。

節点表JSONは末尾の空白行・空白列だけを除外し、見出し＋2工程×261節点の523行・10列を保存しています。全数値はExcel出力と一致します。元の完全なブックとチェックポイントもZIPに保存しています。
