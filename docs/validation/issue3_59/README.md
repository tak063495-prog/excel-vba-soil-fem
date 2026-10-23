# Issue #3・#59の合格記録

[修正内容と結果](../../ISSUE3_59_VALIDATION.md)／[集計と配布SHA256](validation.json)

- `records.zip`：full7（連成17）、mixed6（混在17）、input597（形状9＋直接検査）、final_fault7（中断16＋通常エラー4）、keyboard7（実キー）、subdivision7（正当な曲辺の細分認証）の生記録。
- `before_fs`と`before_ledger6b`は修正前の不具合再現記録。中断を消費してPASSまで続いた記録と、台帳差が残った記録で、合格件数には含めない。before_fsの旧水圧差列は配列順の収集方法が異なるため比較に使わず、PASS・時刻・確定回数で不具合を確認する。
- `source_snapshot.zip`：組込み照合した52モジュールとSHA256。検証用に注入した関数を含めない。
- `saved_states.json`：改修前の受理状態CSV1,853個との完全一致。比較元は[既存記録](../issue49_57/records.zip)のfull3。
- `package.json`と`inputs.json`：コンパイル・保存再読込・ソース・入力全セルの照合。
- `independent_geometry.json`：別計算のdetJ評価。未インストールのsympyに依存せずNumPyで実行できる。

生記録内の実行時パスは来歴として保持する。検証コピーのXLSMと重複DLLは収録せず、配布ブックと固定入力から再現する。

```powershell
pwsh -NoProfile -File tests/run_issue59_regression.ps1 -Book <配布XLSM絶対パス> -Out <新規出力先>
pwsh -NoProfile -File tests/run_issue3_cancel_regression.ps1 -Book <配布XLSM絶対パス> -Out <新規出力先>
python -X utf8 tests/check_issue59_geometry.py
python -X utf8 tests/check_issue59_excel.py --root <形状出力先> --book <配布XLSM> --out <新規集計.json>
python -X utf8 tests/check_issue3_cancel.py --root <中断出力先> --book <配布XLSM> --out <新規集計.json>
```

実キー試験は`tests/run_groundwater_fixture.ps1`の`-CancelAt FS_UI`を用い、keyboard_ready.txt生成後にその専用ExcelへCtrl+Breakを送る。標準の自動回帰は実キー入力を送らない。
