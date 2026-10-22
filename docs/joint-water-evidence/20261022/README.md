# JOINT・水圧検証証跡（2026-10-22）

このフォルダーは、後の説明書PDFで再利用できるように、source22最終ビルドのJOINT（水圧を含む）検証に合格した成果物を保存するための証跡です。入力済みブック、設定条件、計算結果、独立検算、画面キャプチャ、コンパイル・パッケージ記録を同じ検証世代に固定します。最終ビルドで全ステージを再開した後に保存し、ノード・要素結果表と水圧結果表を再オープンして確認できた状態だけを受入成果物とします。

## 受入範囲

|区分|内容|証跡|
|---|---|---|
|正のケース|14ケース（乾燥、湿潤、PRESCRIBED、CONTINUOUS、定常水頭、部分湿潤、滑り、開口、工程、SRM、構造体対、DAVIS）|`ACCEPTED_CASES.md`、`comparison.json`、個別JSON|
|入力拒否|4ケース（モード欠落、圧力不整合、水頭競合、SOIL指定誤り）|`comparison.json` の `rejected`|
|既存機能照合|従来の地下水8ケースを旧版と比較し、変位・応力・指標の差0|`comparison.json` の `unchanged_groundwater`|
|保存再読込|最終受入で確認できた正例ブックを最終状態で保存し、再オープン時の結果署名と結果表を確認|`comparison.json`、`saved_results.json`、`records.zip`|

`docs/manuals/examples/joint_water_examples.zip` には、14正例ブック（ケース名を重複なく改名）と同一DLL（1個）を収録します。ZIP内のブックは最終状態を保存済みで、入力と結果を再現するための例として使用できます。途中のステージ・増分の履歴は最終再開表には残さず、初回からの履歴は `records.zip` のログに保持します。

## 再現・照合

主な記録は次のとおりです。

- `comparison.json`：最終受入の正例、4拒否、既存8ケースの比較結果。
- `compile.log` / `package.json` / `package.log`：31モジュールの厳密コンパイル、保存再読込、配布確認。
- `provenance.json`：ソース基準コミット、ブック・DLL・記録ZIPのSHA-256、Excel 16.0 (Build 20430)。
- `records.zip`：実行ログ、数値CSV・JSON、入力条件、チェックポイント。保存ブックは別の例題ZIP、UI確認は`ui_checks.json`です。
- `joint_input.jpg` / `joint_result.jpg` / `water_input.jpg`：入力・結果画面のキャプチャ。

数値の基準は倍精度のままです。水圧・接合ばね力・反力、全変位、残差、Fsを各個別JSONと実行記録に保存しています。部分湿潤の流体合力は区間積分を基準とし、3点積分点の圧力列だけで合力を再構成しないでください。ノード水圧確認表のQ:U列は外周面とJOINT面を合わせた水圧節点力です。

## 適用上の注意

- JOINT-SRMは固定 `Fs=1.2` の照査例です。JOINTを含む斜面の全探索結果を示す資料ではありません。
- DAVIS例は材料則の既存確認と接合滑りの確認です。湿潤斜面の長時間SRM性能比較を意味しません。
- 自由水面、不飽和、過渡圧密、JOINT沿いの透水抵抗・開口幅流れ、双方向連成はこの証跡の受入範囲外です。
- 開発途中の失敗試行や未保存の診断状態は、この受入ZIPへ含めません。説明書の図・数値には受入ブックと対応する検算・再読込記録を使用します。

## 初回ログと再開ログ

初回の受理数値は各result.json・出力表JSON・チェックポイントCSVです。通常の_outログは再開開始時に更新されます。wet・stages・wet_srmは初回詳細ログを再取得し、全数値が受入22と一致したことをretention_checks.jsonに記録しました。初回詳細ログはrecords.zipのretention/*/initial_run_logs、その他の通常_outログは再開側です。その他の再開ケースの初回詳細反復ログはこのZIPに残っていません。公開ドライバーは以後の再開試験で初回ログも退避します。

## 版の固定

実装基準、ビルド名、ブック・DLL・記録ZIPのSHA-256は、最終 source22 実行後に更新された `provenance.json` と `ACCEPTED_CASES.md` を正とします。source18/21の開発途中の結果は最終受入ZIPに含めません。

[入力済み14ブックを取得](../../manuals/examples/joint_water_examples.zip)／[ケース台帳](ACCEPTED_CASES.md)／[検証概要](../../JOINT_WATER_VALIDATION.md)／[保存再読込](saved_results.json)／[画面確認](ui_checks.json)／[画面確認の元データ](ui_checks_raw.zip)。
