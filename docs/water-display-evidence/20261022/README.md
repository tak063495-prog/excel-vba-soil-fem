# 水理表示・工程別入力の受入記録

条件と結論は [WATER_DISPLAY_VALIDATION.md](../../WATER_DISPLAY_VALIDATION.md) を参照してください。各ケースの `result.json` は全自由度変位と積分点出力・Fs診断・実行時間です。日本語名のJSONはExcel結果表の保存値で、空セルも含みます。大きな `結果節点.json` は内容を変えずgzip圧縮しています。

`calculation_logs.zip` に実計算のCSV・ログ・チェックポイント、[実計算ブック集](../../manuals/examples/water_display_examples.zip) に計算済みXLSMとDLL1個を保存しています。`ui_*/*.json` は専用コピーの実画面検証時の図形名・文字・位置で、画面画像は説明書の `images/water_display` にあります。`ui_hydro/stale_error.txt` は工程別入力の変更後に旧結果を拒否した記録です。

`manual_prose_preservation.json` は旧第5・6・8～24章（新第5・6・9～25章）の本文照合、`manual_page_mapping.json` は改ページ後の用語集参照の照合です。章参照・図参照・ページ番号・空行を除いた本文を保持しました。PDFの最終ページ・リンク・全ページ画像の確認記録は別途 `manual_checks.json` へ保存します。
