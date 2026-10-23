# 内側CSVライタの中断・保存回帰

[修正と実測](../../ISSUE3_INNER_WRITERS.md)／[集計](validation.json)。`records.zip`はbefore_*の改修前再現、faultsの24条件、fullの通常連成17入力と再開、keyboardの実キー記録を含む。失敗再現を合格結果として数えない。XLSMと重複DLLは収録しない。`source_snapshot.zip`は配布ブックと照合した52ソース。

```powershell
pwsh -NoProfile -File tests/run_issue3_inner_regression.ps1 -Book <配布XLSM絶対パス> -Out <新規出力先>
python -X utf8 tests/check_issue3_inner.py --root <出力先> --book <配布XLSM> --out <集計.json>
```

入力は`tests/fixtures/fixed_stress/column_inner_io.json`。通常のcolumn_smokeへ後続WAITを追加し、工程保存の中断後に進まないことを検査する。実キーは`tests/run_groundwater_fixture.ps1 -CancelAt FSCSV_TIME_UI -CheckLedger`を使い、keyboard_ready.txtが出た専用コピーでCtrl+Breakを送る。自分のExcel以外を中断しない。
