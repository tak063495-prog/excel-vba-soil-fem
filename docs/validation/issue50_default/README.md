# Issue50：時間詳細行なしWATERの再開

[修正と実測](../../ISSUE50_DEFAULT_RESUME.md)／[配布・ソース識別](validation.json)／[入力・実計算ログ](records.zip)／[52モジュールのソース](source_snapshot.zip)／[mainでの再現](baseline_reproduction.json)／[AUTO12条件](matrix_auto.json)／[VBA単独2条件](matrix_off.json)／[別Excel・時刻別読込](extra.json)／[通常17条件](regression.json)／[全入力・設定](inputs.json)／[配布同一性](release.json)

records.zipには期待どおり不一致を再現したbefore_completedも含む。これは合格ケースではない。修正版の合格と混同しない。試験専用コピーの観測コードはtests/issue50_observer.basとrun_groundwater_fixture.ps1の-WaterObservablesで再現でき、配布ブックには含めていない。観測あり／なしの変位・応力・診断値が新規／再開とも完全一致することを確認。

旧版の正しいWATER保存状態を修正版で復元する互換確認を行った。旧版の誤った再開条件で既に計算した後続結果の自動修復ではない。元の正しいWATER保存状態から再開するか、入力から計算し直す。
