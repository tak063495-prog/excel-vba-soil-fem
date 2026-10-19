# INCONSISTENT高速化の検証データ

2026-10-19夜〜20日（日本時間）、Windows 11・64bit Excelで測定しました。
解釈・条件・限界は [比較報告](../../INCONSISTENT_NATIVE_PERF.md) を参照してください。

|ファイル|内容|
|---|---|
|`environment.json`|CPU/OS/Excel、コンパイラ、QPC周波数、ソースのSHA256|
|`pair1～3_enabled0/1.json`|計測コードなしの同一ブック/DLLによる全探索。170自由度の全変位を含む|
|`benchmark_summary.json`|3組の時刻・値の検査と中央値|
|`profile_discovery*`|初回36手続きの内訳調査|
|`profile_before*`, `profile_after*`|39手続きの階層型QPC計測とexclusive時間の集計|
|`profile_injection.json`, `profile_comparison.json`|計測注入対象と前後の解の比較|
|`final_release*`, `package_settings.json`|最終ソースを埋め込んだ配布ブックの全探索、保存された設定とブックハッシュ|
|`strict_to_release.json`|従来STRICT/x87と配布AUTOの丸め差|
|`matrix_*`, `regression_summary.json`|INCONSISTENT/DAVIS、Case1〜4、off/adaptiveの16比較|
|`builtin_0/1.json`|同梱306要素モデル、INCONSISTENT、固定Fs=1の比較|
|`kernel_band.json`, `kernel_gmres.json`, `kernel_state.json`|DLLの独立残差・ガード・失敗時無変更の試験|
|`state_vba_integration.json`|実VBA構造体全体の30比較。形状不適合・設定欠落・OFFも含む|
|`gmres_dispatch_new/legacy.json`|設定欠落/0/1、STRICT/SSE2/AUTO、旧DLLへの分岐と独立解の10比較|
|`solver_event_counts.json`|全探索の再始動回数・CSVサイズ・SHA256。再始動数は丸め差により変化|

第1組の記録時点では丸め前の `Bracket` フィールドをまだ追加していなかったため、
両側ともそのフィールドがありません。Fs表示と数値の中点は一致しています。
第2・3組、最終配布ブック、39手続きの前後比較では丸め前の下端・上端・中点・幅を直接記録して一致を確認しました。
この保存データの再集計には `summarize_inco_benchmark.py <本フォルダ> --out <新規JSON> --allow-legacy-bracket` を使います。
新たに計測する場合はこのオプションを付けず、両側の丸め前Fs情報を必須としてください。

3組の計測後、旧ブック互換のため設定欠落時の既定値を1→0に戻し、コメントを修正しました。
ペア計測はすべて設定値を明示しており、その変更の影響を受けません。
最終配布ブックは設定値1を保存し直した状態をファイル内から確認し、全探索を別途再実行しました。
DLLは全ペア・最終配布とも同じSHA256です。
