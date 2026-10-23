# 1024版 時間連成説明書の再現資料

- [PDF](../../manuals/2DSoilFEM_Manual_20261024.pdf)／[検証本文](../../MANUAL_FIXED_STRESS_VALIDATION.md)
- [入力画面原寸とセル値](actual_excel_captures.zip)：`screens`／`screens2`／`screens3`。最終配布ブックの専用コピーに合格済みケースの入力・保存表を移し、現行UIで描画して撮影。計算・採用状態は元の合格記録に対応する。
- [画像注釈](../../manuals/images/fixed_stress/annotation_manifest.json)：原画像SHA、セル範囲、切抜き範囲と赤枠番号。
- [作図値](../../manuals/images/fixed_stress/figure_data.json)：元CSV／JSONのSHA、同一物理GPの理論値、MAINLINEのみの全時刻歴。
- [全ページ一覧](all_page_review_sheets.zip)／[構造検査](structure_checks.json)／[確認値保持](coverage.json)／[ページ参照](printed_page_mapping.json)／[公開リンクとレイアウト](pdf_link_bounds_checks.json)
- [保存結果の直接再表示：本線](frame_reopen_c12/frame_reopen.json)／[SRM枝](frame_reopen_s11/frame_reopen.json)。新規Excelセッションから事前ロードなしでボタン経路を実行。

数値の元記録は [fixed-stress-evidence](../../FIXED_STRESS_VALIDATION.md) と [xlsm-issue-evidence](../../xlsm-issue-evidence/20261024/README.md)。追加の再計算ではなく、受入済みの生データを抽出して作図する。実行はリポジトリルートから：

```powershell
python -X utf8 tests/prepare_fixed_stress_manual_figures.py
python -X utf8 tests/annotate_fixed_stress_manual.py
$env:MANUAL_PUBLIC_REF='main'
$env:MANUAL_PDF_TITLE='2DSoilFEM 操作説明書 20261024'
python -X utf8 docs/manuals/build_manual_pdf_v2.py docs/manuals/2DSoilFEM_Manual_20261024.md docs/manuals/2DSoilFEM_Manual_20261024.pdf
python -X utf8 tests/review_water_manual.py docs/manuals/2DSoilFEM_Manual_20261024.pdf tests/tmp/manual_fs/review --source docs/manuals/2DSoilFEM_Manual_20261024.md --public-ref main --book-version 20261024 --last-example 23
python -X utf8 tests/check_manual24_coverage.py
python -X utf8 tests/check_manual24_pdf.py
```

作図スクリプトは公開ZIPから必要な2ケースの受入生データを専用一時フォルダーへ安全に展開し、入力fixtureのSHAとPASS記録を照合する。Mandelの理論値は最大誤差から逆算せず、独立級数をGP座標で評価する。注釈スクリプトは原画面ZIPから展開する。Pythonのnumpy／scipy／matplotlib／Pillow／PyMuPDF／markdown、ChromeまたはEdgeが必要。注釈の数字はWindowsのArialを使用。
