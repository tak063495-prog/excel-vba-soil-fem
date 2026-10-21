# 説明書レビュー対応とUI_05

SRM凡例はさらに「最後の有効行のSRM」へ訂正しました。現行ZIP/配布SHAはこの文言を含みます。今回の解析再実行はなく、保存結果の不変性とコンパイル・再生成・再読込を確認。[追加訂正記録](../../../../reviews/2DSoilFEM_srm_legend_followup_20261021.md)。PR29時点の配布SHAは`../srm-legend/pr29_release.json`に保持しています。

PR28後のN01-N05補訂は[追加検証](../manual-rereview/README.md)を参照。現行の配布原版・道路実習ZIPはFEMIoの凡例3文言を訂正したUI_05です。以下の38項目と道路比較記録はPR28時点の履歴として保持し、今回の94/143項目・同一セッション2回解析と区別します。補訂前の配布SHAは`../manual-rereview/ui05_pr28_release.json`に保存しました。

レビューR01–R16の対応先・未実施は[改訂対応表](../../../../reviews/2DSoilFEM_manual_response_20261021.md)を参照してください。既存レビューの評価と今回の実行結果を混ぜません。

- [最初の道路実習：UI_05計算済みブックZIP](road_ui05_workbook.zip)。XLSMと単一DLLを別の新規フォルダーへ展開し、メッシュは再生成せず確認から始めます。328節点・95 Q8・11工程、工程1/11の支持反力1440/1842kN。
- [UI_05配布SHAと6モジュールSHA](release.json)、[38項目のExcel確認](ui_checks.json)、[保存・再読込・再表示](roundtrip.json)。入力規則はValidation.ValueとSTOPのメタデータを実Excelで検査。物理的なキー入力とエラーダイアログの初心者試験とは区別します。
- [道路の2条件の完全一致・手計算](review_summary.json)、[γ変更入力](road_gamma20.json)。元とγ変更の各11工程について入力表・全節点結果・全GP応力/ひずみ・工程表を修正前と完全一致で照合。時間値は性能保証ではありません。
- road_ui05 / road_gamma20_ui05：修正版の計算記録・独立検算・実行CSVログ（gzip）。road_gamma20：比較元UI_04のγ変更記録。元のγ19の比較元は既存20261021/road01です。
- actual_A.svg / actual_B.svg：保存済みwall_davis08の同じFs=1.7・同じ最終収束増分を、色/矢印を修正した表示処理でExcelから再描画。再解析ではありません。元の結果表は変更していません。
- tunnel_dimensions.png：寸法の説明図。新しいトンネル解析結果ではありません。
- manual_section_source.py：既存組版へ追加したページ定義。全冊の単独ビルダーではありません。

## 再集計

リポジトリのルートで`python tests/check_manual_review.py docs/manuals/construction-evidence/20261021/manual-review --check-only`。gzipの公開値を直接読み、元の道路記録との完全一致・γだけの変更・1440/1864kN・一様弾性自重沈下64.405714mmを確認します。Excelは起動しません。

実Excelの再確認は`pwsh -File tests/verify_manual_review.ps1 -Book <専用候補コピー.xlsm>`。元ブックは残します。このドライバーは指定候補へ6モジュールを更新するため、配布原版を直接指定しません。
道路の再計算は`tests/run_manual_case.ps1 -SourceWorkbook <UI_05コピー> -Fixture tests/manual_cases/road_widening_shape.json -OutDirectory <未作成フォルダー> -Flow INCONSISTENT`。γ変更はFixtureに本フォルダーのroad_gamma20.jsonを指定します。既存Excelがあれば開始せず、同時には実行しません。
`tests/render_manual_review.ps1`は計算済みブックのコピーに描画モジュールとテスト専用SVG書出しフックを入れ、A/Bを再描画します。公開アプリのFEMViewer自体にはフックを入れていません。

## 制約

Issue #26のゼロ局所変位・非ゼロ閾値・材料配列非変更・観測スナップショットは合成テスト。DLL/VBAによる実塑性保持工程や全SRM探索の新規回帰を網羅したものではありません。Issue #27は数式参照を自動で書き換える修正ではなく、危険な配置変更と数式が併存する場合の変更前停止です。開いていない別ブックの外部参照まで自動保全しません。JOINT/RESETの新規実計算や初心者による独力再現テストは実施していません。
