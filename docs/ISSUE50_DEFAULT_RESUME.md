# Issue50残件：時間詳細行を省略したWATERの再開

## 原因と修正

[追加指摘](https://github.com/tak063495-prog/excel-vba-soil-fem/issues/50#issuecomment-6365833340)はmain `deeeb453`の配布ブックでも残存していた。工程別水理・時間設定はあるが、時間詳細AC:ALの行がないWATERで、通常実行だけがSTARTへ上書きし、読込側はHOLDのまま開始プロファイルを復元していた。

`FSGetOptions`へ工程種別を渡し、実効既定値を共通化。詳細行なしWATERはSTART、明示HOLDおよび行あり・水位変化空欄はHOLDとする。通常実行・保存再開・時刻別読込が同じ判定を使う。LINEAR／時系列の時刻補間と、WATER以外のHOLDは維持。

倍精度・材料式・材料定数・力学/連成収束条件・SRM条件・保存形式は変更しない。計算識別07を維持し、旧main07の正しいWATER保存状態からの再開も確認した。入力欄とDLL1個、PDFは維持する。

## 実Excelでの再現

初期W（4 m）→対象W2（5 m）、幅1 m・地表Y=2 m、γw=9.81 kN/m³。GRAVITY→WAIT→WATER40 s→WAIT20 s。WATER完了保存からの再開で、mainはW2→Wへ戻り、外水圧の鉛直合計は−29.43→−19.62 kN、支持反力は69.43→59.62 kNとなった。最大全変位差は1.398291761e-05 m。保存GP水圧が読めても、境界と荷重の不整合は残る。[再現の生記録](validation/issue50_default/baseline_reproduction.json)。

修正版の外水圧は連続・復元直後・次WAITで−29.43 kNを保持する。反力は自重40 kN＋外水圧29.43 kN＝69.43 kNと照合した。

## 検証

- strict buildと配布コピーの実Excelコンパイル、52モジュールの保存前／再オープン後照合に合格。ブックSHA256 `597a168818ccc73ee46278a430f261a3d8d173de83264adcd2dc1c25935ddaf8`。
- 詳細行なし、明示HOLD、明示START、LINEAR、行あり空欄モード、時系列の6条件×途中／完了再開＝AUTO12条件。VBA単独OFFでも省略時の途中／完了2条件に合格。
- 復元直後の水位プロファイル・全境界圧・外水圧全自由度・全節点水頭・全GP流束/水圧・境界流量と、後続WAITを比較。AUTO12条件の最大全変位差は4.931e-13 m。省略時と明示STARTの連続結果は全変位・全水理観測値が完全一致。
- 時系列の折点間（補間率0.5）からの再開を追加確認。別Excelでの直接再開は途中・完了・旧main07の保存状態の3条件。時刻別結果の読込2条件でも境界・荷重を確認。画面の描画試験と、この保存状態読込の数値確認は区別する。
- 観測コードなしの対照で、新規／再開の全変位・全応力・診断値が観測ありと完全一致。試験コードは配布へ含めない。
- 既存17条件（正常計算・保存再開と期待拒否）では、正常系の全変位・全応力・Fs区間・反力・残差が前回配布と完全一致し、拒否系の停止理由・保持状態も一致。受理状態CSV 1853個がバイト一致。入力/設定9表10115セルが一致。

[入力・生ログ・独立照合・配布識別](validation/issue50_default/README.md)。既に旧版の誤った境界で再開した後続結果は自動補正せず、元の正しいWATER保存状態から再開するか、入力から再解析する。

## 再現手順

リポジトリルートで、別のExcelを起動していない状態で専用コピーを使う。

```powershell
python -X utf8 tests/generate_issue50_fixtures.py --source tests/fixtures/issue49_57/water_start.json --out tests/fixtures/issue50
pwsh -NoProfile -File tests/run_issue50_resume.ps1 -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -Out tests/tmp/issue50_auto_new
pwsh -NoProfile -File tests/run_issue50_resume.ps1 -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -Out tests/tmp/issue50_off_new -Native OFF -Cases default_start_no_detail
```

途中はC000003（工程3の経過20 s）、完了はstg03_WATERを使う。折点間の時系列確認は同じfixtureをC000002（経過10 s）から再開する。`run_issue_cold_resume.ps1 -WaterObservables`で別Excelから直接再開でき、`-ViewOnly`で時刻別状態だけを読み込む。入力・観測・比較ドライバーは同梱する。
