# 行列組立高速化の検証記録

基点は `87d4f13ddb80584d0a43fecad9030a8a6430122f`、採用ビルドは `20261020_INCO_ASSEMBLY_01`。
Double・収束条件・材料則・SRM探索条件を維持しています。

- `discovery.*` / `discovery_summary.json` / `discovery_injection.json`: 前版の61項目QPC計測。最大のCSR散布表作成4.913秒を特定しました。
- `pair{1,2,3}_enabled{0,1}.json`: 同一ブック・同一DLLでNATIVE_ASSEMBLYだけを変更した全探索の生データ。全変位・丸め前Fs区間・残差・主要回数・ログSHA-256・高精度時間・Excelの単独使用確認を含みます。
- `benchmark_summary.json`: 全3組で短縮し、中央値53.718秒→45.298秒、15.67%短縮。数値・ログの完全一致と、高速側で実際にDLL散布表が使われたことを必須にします。
- `profile{0,1}.json` / `.csv` / `_summary.json` / `profile_comparison.json`: 同一の62項目計測ブックによる0/1比較。散布表4.937秒→0.528秒、行列組立9.181秒→5.364秒。計測ブック全体は54.328秒→52.113秒。主な速度比較の3組へ混ぜません。排他的時間の総和と全体時間が一致します。
- `profile_injection.json` / `profile_build.log`: 計測境界と実際のVBAコンパイル・保存前後のソース照合記録。
- `vba_arrays.json` / `legacy_arrays.json`: 新DLL・旧DLL各54条件。実VBAの散布表全配列と要素UDT全バイト、キャッシュ再利用を比較。要素数1/2/20、ゼロ/疎/密行列、非活動要素、設定欠落/0/1、AUTO/OFFを含みます。
- `INCONSISTENT_{0,1}.json` / `DAVIS_{0,1}.json` と各 `_summary.json`: Case1〜4×off/adaptiveの固定Fsで16組・32解析。全170変位・Fs・残差・主要回数・ログが一致。INCONSISTENTではDLL呼出しが正、DAVISでは0を必須にします。
- `builtin_{0,1}.json` / `builtin_summary.json`: 同梱306要素、固定Fs=1。全2,010変位とログが完全一致、残差1.05571945918321e-13。時間2.854秒→3.196秒で短縮していないことも記録します。
- `previous_release_check.json`: 前回公開版から数値結果が変わらず、ビルド番号列を除く109,847イベントが一致。ビルド文字数の増加によるログサイズ差を速度改善に含めません。同一ビルドの0/1比較では全バイト一致を必須にします。
- `assembly_unit.log`: 独立ctypes試験3メソッド。24入力構成でcount-only/fill、26の無効パラメータ・NULL要求、両パスの非有限値・無効CSR位置を検証。入力不変・拒否時の出力保持・出力末尾ガードを含みます。
- `native_compile.log`: MSVC /fp:strict /W4 /WXでのDLLビルド記録。
- `simd_checks.json`: 既存の因子分解・解法の確認。STRICTの因子・右辺・解が前版DLLと全バイト一致。
- `gmres_checks.json` / `state_checks.json`: 既存GMRESの解・残差・失敗時出力保持と、材料状態コピー19ケース。
- `evidence_unit.log` / `inco_evidence_unit.log`: 比較器の成功例と、数値・ログ・設定・経路の改変を拒否する試験。
- `environment.json` / `verification.json`: CPU・Excel・タイマー・ソース、コンパイル、保存設定、ブック/DLL/ZIPのハッシュとZIP内容確認。

JSONの絶対パスは実行時の所在です。使い捨てブックと生CSVはGit管理外に保存します。
配布物は検証したブック・DLLと同一バイトで、計測・テスト用の追加コードを含みません。
短い固定Fs解析での速度改善や、任意モデルで同じ短縮率を保証するものではありません。
