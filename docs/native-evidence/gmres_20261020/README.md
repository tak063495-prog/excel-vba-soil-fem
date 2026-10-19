# GMRES倍精度SIMD 検証記録 2026-10-20

`benchmark_summary.json` と `pair{1,2,3}_enabled{0,1}.json` が配布版の全探索比較です。旧/新は同一DLLの設定0/1。中央値8.45%短縮で、10%達成とは扱いません。
Fs・全170変位・最終残差・主要回数は完全一致します。時間変動と範囲は集計へ残しています。

`native_profile_before*`、`native_profile03*` は別の計測用DLLによる階層QPCです。全体速度の比較へ混ぜません。`gmres03_profile_names.json` は集計配列の対応です。
`native_profile02*` と `pilot/` は採用前の加算順序保持案、`pilot03/` はSIMD案の初期単発試行で、主3組とは別です。

`log_verification.json` は独立したPythonで実CSVのSHA、件数、残り全バイトを検証した結果です。
GMRES_*行とREGULARIZED_CORRECTION_REQUESTのlinear_relresだけを比較から除外し、それ以外の10,869イベントの一致を確認しています。配布版のログは一切省略しません。
`solver_events_pair1.zip` に両側の実CSVを保存しました。全ログのバイト一致は主張しません。

INCONSISTENT/DAVIS/builtinの各0/1とsummaryは固定Fsの17組です。これらはGMRES呼出0で、非対象経路の回帰確認です。
`tuned03_unit.json` の87ケースは成功するGMRESも含む独立真残差・既知解・拒否時出力保護です。帯域行列でILU SIMD経路を確認しています。
`legacy_gmres_check.json` は前版DLLと旧エクスポートの120ケース比較で、全解・作業領域・イベントバイトが一致します。再現はリポジトリ直下から `python docs/native-evidence/gmres_20261020/legacy_export_check.py`。
dispatchの2記録は新DLL/旧DLL各5条件です。状態コピーと既存SIMD、比較器の結果も同梱します。

`final03_book_compile.log` は実VBAコンパイルと保存・再オープン後の8モジュール照合、`native03_compile.log` は警告をエラーにするDLLコンパイルです。
`verification.json` と `environment.json` に配布物・ソース・入力のSHA-256と環境を記録しています。
計測用DLLと一時ブックは配布しません。配布ZIPはXLSM、DLL2個、README、build.jsonの5個です。

initial_cli_failure.log は初期試験スクリプトのCLI引数不足で起動できなかった記録です。修正後の最終試験結果は 	uned03_unit.json / .log（87ケース合格）です。
