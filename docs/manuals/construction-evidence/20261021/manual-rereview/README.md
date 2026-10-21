# 説明書再レビュー N01-N05の追加検証

[対応表](../../../../reviews/2DSoilFEM_manual_rereview_response_20261021.md)。旧PR28の結果はmanual-reviewフォルダーに保持しています。

- `generic_checks.json`：汎用UI_05候補の94項目。コンパイル、凡例/入力規則、保存・再読込、埋込モジュール、空ステージ生成。
- `road_checks.json`：道路11工程を同一セッションで2回計算した143項目と実測時間。全9入力/結果表が更新前保存値と完全一致。
- `road_baseline_tables.json.gz` / `road_run1_tables.json.gz` / `road_run2_tables.json.gz`：比較した9表。各値はPowerShellのExcel Value2行列のJSON文字列。生値を丸めず、文字列単位の一致で照合。
- `run1_logs/` / `run2_logs/`：それぞれの計算直後に退避したCSV/ログ（gzip）。数値表の結果と対応する実行記録。
- `ui05_pr28_release.json`：補訂前UI_05の配布SHAを保存した履歴。現行配布SHAは`../manual-review/release.json`または`workbook/ui_20261020/BUILD_RECORD.json`を参照。

Excelを起動しない公開値の照合：

```powershell
python tests/check_manual_rereview.py
python tests/check_manual_review.py docs/manuals/construction-evidence/20261021/manual-review --check-only
```

実Excelでの再現（Windows ExcelとVBAプロジェクトへの参照許可が必要）：

```powershell
# 同梱DLLが隣にある元ブックを指定。出力先は未作成のファイルにする。
pwsh -NoProfile -File tests/verify_manual_rereview.ps1 -SourceBook <展開済み道路例.xlsm> -DestinationBook <専用フォルダー/road.xlsm> -Out <専用フォルダー/road_checks.json> -RunRoad
```

このドライバーは元ブックをコピーし、コピーのFEMIoだけを更新します。起動中のExcelがあれば開始せず、起動した専用Excelだけを終了します。元ブックのSHA不変も確認します。汎用ブックの検査は`-RunRoad`を付けずに行います。入力規則はプログラム検査で、物理的キー入力による初心者試験とは区別します。二つの解析が同じ結果であることは確認していますが、新しいJOINT/RESET・SRM全探索の検証ではありません。
