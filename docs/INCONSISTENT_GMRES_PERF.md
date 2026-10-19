# INCONSISTENTのGMRES内部計測と倍精度SIMD化（2026-10-20）

基点は行列組立版 `9dc8e8a`。ビルドは `20261020_INCO_GMRES_03` です。
[XLSMとDLL2個の一括ZIP](../workbook/SoilFEM_INCONSISTENT_GMRES_20261020.zip) を展開して同じフォルダに置いてください。
保存済み結果は元ブックのものです。入力を確認して再解析してください。

## 計測と採用した変更

前版の階層QPCでは、GMRES計算が全体の21.1%、材料評価が14.7%でした。
材料側のVBAテーブル再構築は0.182秒（79,993回）で、材料評価の大半はDLL内部です。
テーブルのキャッシュ化だけでは全体10%に届かないため、最大のGMRESを先に分解しました。
材料DLLと材料則は今回変更していません。

使い捨てDLLに関数ごとのQPCを追加した実測です。親子の時間は階層的に分離しています。

|処理|呼出回数|排他的な秒数|
|---|---:|---:|
|ILU前処理|1,660,000|4.818|
|CSR行列ベクトル積|1,759,808|3.865|
|内積|35,811,584|1.989|
|ベクトル更新|35,811,584|1.106|
|真残差の行列積以外|99,808|0.015|
|GMRESの残り|830|1.743|

GMRES全体は13.537秒。タイマーを多数呼ぶ費用を含むため、配布版の速度比較には使いません。
残差の親手続きは0.236秒で、その内側の行列積をCSR行へ含めています。

最初は積4個を先に計算し、加減算順序を完全に保持する案を試しました。
単発の全探索は45.614秒→45.163秒で約1%しか変わらず、これを確定した高速化とは扱いません。
補助関数をインライン化した試作02でも、ILUは4.849秒とほぼ変わりませんでした。
その結果を踏まえ、採用版03ではILUの三角解法とKrylov側のCSR行列積の内積をSSE2の倍精度2レーンで計算します。
複数の累積器で積を加算し、最後にまとめるため加算順序が変わります。
既存のAUTO/AVX2の内積・ベクトル更新、STRICTの旧演算選択は維持します。

維持した条件:

- 浮動小数点の変数・配列・SIMD演算はDouble（binary64）。Singleへの変換やfast-mathを使いません。
- 直交化は従来の2回の修正Gram–Schmidt。
- 真残差は従来のスカラーCSR計算を使い、再スタート時と終了時に同じ許容値で判定。
- 反復上限、再スタート幅、ILU対角判定、ゼロ判定、非有限値・失敗時の出力保護を維持。
- 材料則、非線形収束条件、SRM探索条件は変更しません。

`NATIVE_GMRES_TUNED=1`、`NATIVE_DLL=AUTO`、`NATIVE_GMRES_SIMD=1`かつSSE2/AUTOのときに使用します。
キー欠落・0・新エクスポートがない旧DLLでは従来のGMRESへ戻ります。
STRICTは従来どおりx87参照経路へ戻ります。配布ブックには新キー1を保存し、開き直して確認します。

## 全探索の比較

|組|従来（秒）|高速側（秒）|短縮率|
|---|---:|---:|---:|
|1|51.769252|41.319738|20.18%|
|2|48.452008|44.995849|7.13%|
|3|45.167073|44.359437|1.79%|
|中央値|48.452008|44.359437|**8.45%**|

3組とも短縮しましたが、今回の中央値では10%に届きません。従来側45.17～51.77秒、高速側41.32～45.00秒と時間変動があり、短縮率はこの環境・モデルでの観測値です。

同一ブック・同一DLLで新キーだけを0/1へ変更しました。
Case1/off、INCONSISTENT、精度AUTO、既存の状態コピー・GMRES SIMD・I/O・イベント・行列組立の高速経路は両側で1です。
入力準備後の解析全体を高精度Stopwatchで測定します。各回は独立したExcelで、順序は0→1、1→0、0→1。
計測用DLLを使った時間や初期試作の結果をこの3組へ混ぜていません。

3組ともFs下限2.200000000000001、上限2.2125000000000012、中央値2.206250000000001、探索17回が完全一致。
全170変位の最大絶対差は0、最終残差は両側とも **6.5595364013624e-13** です。
成功増分7、分解6,713、再利用4,078、塑性点76も一致。高速側は全830回で新GMRESを使用し、fallback=0でした。
前版配布DLLの旧GMRESエクスポートとの独立比較120ケースでも、解・作業領域・イベントの全バイトが一致しました。

GMRES再スタート行は98,978→99,048、全イベントは109,847→109,917へ増えます。
比較用に除外する診断値以外の10,869イベントは、3組とも順序・全バイトが一致しました。
これはイベントの省略による高速化ではありません。

途中のGMRES残差と再スタート回数は加算順序の変更により差が出ます。
失敗した線形解法の残差を表示する `REGULARIZED_CORRECTION_REQUEST` の `linear_relres` にも差が伝わります。
ログ全体のバイト一致を主張しません。全CSVのSHA-256・件数を両側で残します。
比較器はGMRESイベント行と上記の残差表示値を除いた、ステージ・Fs・増分・他のイベント・詳細をハッシュ比較します。
この正規化はテスト上の比較だけで、配布版の実ログは省略・改変しません。
最終残差・全変位・Fsと主要回数は別途比較し、この範囲の微小差で隠しません。

## 内訳の再計測と検証

|処理|従来（秒）|高速側（秒）|
|---|---:|---:|
|ILU前処理|4.818030|4.322229|
|CSR行列積（真残差内を含む）|3.864776|2.347597|
|内積|1.989152|1.983249|
|ベクトル更新|1.106126|1.228215|
|真残差の行列積以外|0.015137|0.015133|
|GMRESの残り|1.743407|1.750369|
|GMRES全体|13.536628|11.646792|

親子の排他時間の合計をQPC総時間と照合済みです。行列積と前処理を合わせて約2.01秒短縮しました。
GMRES呼出830、総反復1,660,000は両側で同じです。
各GMRESはこの全探索モデルでは反復上限に達し、既存の代替解法へ進みます。上限・代替条件は変えていません。
途中経路の丸め差により内積・ベクトル更新の呼出は35,811,584→35,800,044、真残差確認は99,808→99,878へ変わります。
QPCの計測負荷と実行時間の変動があるため、この単発内訳を全体短縮秒数の厳密な分解とは扱いません。

- 配布DLL: MSVC x64 `/O2 /fp:strict /W4 /WX` でコンパイル成功。
- 配布ブック: VBAの実コンパイル成功。保存・開き直し後も8モジュールのSHA-256と設定を照合。
- 倍精度GMRES: 独立真残差・既知解・失敗時出力保護など87ケース合格。成功解の従来との差は最大4.440892098500626e-16。
- SSE2/AVX2両モード、STRICT拒否、ILUの下三角・上三角とも4要素以上ある帯域行列を含めて確認。
- VBA切替: SIMDキー欠落/0、STRICT、SSE2、AUTOの5条件を新DLL・旧DLLで各確認。旧DLLでは新関数へ入らず従来関数を使用。
- 状態コピー19ケース、既存の対称・非対称SIMDとSTRICT参照比較、比較器の否定試験も合格。

固定Fsの回帰比較は次の17組です。全組でFs・全変位・最終残差・主要回数・CSV全体が一致しました。
これらの短い解析ではGMRES呼出0で、新経路の速度試験ではなく他経路への影響確認です。

|流れ則・モデル|条件|従来（秒）|高速側（秒）|最終残差（両側同値）|
|---|---|---:|---:|---:|
|INCONSISTENT|Case1/off|2.176|2.186|9.44351e-16|
|INCONSISTENT|Case1/adaptive|1.569|1.584|5.44705e-15|
|INCONSISTENT|Case2/off|1.581|1.560|1.16075e-13|
|INCONSISTENT|Case2/adaptive|1.562|1.620|1.25044e-13|
|INCONSISTENT|Case3/off|1.613|1.649|4.72993e-12|
|INCONSISTENT|Case3/adaptive|1.698|1.589|3.83478e-11|
|INCONSISTENT|Case4/off|1.577|1.585|7.54343e-10|
|INCONSISTENT|Case4/adaptive|1.591|1.561|9.04295e-10|
|DAVIS|Case1/off|2.177|2.133|1.04569e-15|
|DAVIS|Case1/adaptive|1.578|1.563|4.22510e-15|
|DAVIS|Case2/off|1.558|1.621|2.71604e-09|
|DAVIS|Case2/adaptive|1.591|1.601|5.62935e-08|
|DAVIS|Case3/off|1.605|1.590|1.46030e-09|
|DAVIS|Case3/adaptive|1.567|1.572|5.35674e-10|
|DAVIS|Case4/off|1.556|1.575|1.66834e-10|
|DAVIS|Case4/adaptive|1.555|1.553|1.23204e-10|
|builtin|builtin/default|3.098|2.846|1.05572e-13|

同梱306要素モデルは全2,010変位の最大差0。短い固定Fs解析の時間差は測定変動を含むため、高速化の根拠へ使いません。

この確認は小モデル全探索と同梱モデルなどの範囲です。任意モデルで同じ短縮率になる保証はありません。
材料評価DLL内部の最適化は今回の変更範囲に含めていません。
生データ・環境・比較仕様は [検証記録](native-evidence/gmres_20261020/README.md) を参照してください。

## 再現手順

64bit Windows Excel、Visual Studio C/C++ x64、Python、PowerShellを使用します。
VBAプロジェクトへのアクセスを信頼する設定が必要です。
他のExcelが起動中なら開始を拒否し、テスト用コピーと自分が起動したExcelだけを扱います。
既存の出力を上書きしないため、新しい出力先を使ってください。

```powershell
cmd /c native\build.cmd
pwsh -NoProfile -File tests/native/build_gmres.ps1 -OutputDirectory "$PWD/tests/tmp/gmres_rebuild"
python tests/native/test_gmres_tuned.py --out tests/results/gmres_unit_new.json -v
$book="$PWD/tests/tmp/gmres_rebuild/2DSoilFEM_inco_gmres_20261020.xlsm"
pwsh -NoProfile -File tests/native/run_inco_benchmark.ps1 -Book $book -OutputDirectory "$PWD/tests/results/gmres_pairs_new" -GmresOnly -Pairs 3
python tests/native/summarize_io_benchmark.py tests/results/gmres_pairs_new --scope gmres --out tests/results/gmres_pairs_new/summary.json
pwsh -NoProfile -File tests/native/run_gmres_regression.ps1 -Book $book -OutputDirectory "$PWD/tests/results/gmres_regression_new"
pwsh -NoProfile -File tests/native/check_gmres_dispatch.ps1 -Book $book -OutputDirectory "$PWD/tests/results/gmres_dispatch_new" -TunedExport present
```

内部計測は `instrument_gmres.py` で使い捨てソースを作り、`build_gmres_profile.cmd` で別フォルダへDLLを構築します。
`run_native_compare.ps1 -NativeProfile` はQPCをリセットし、同一Excelスレッドから集計値を取得します。
`summarize_gmres_profile.py` はノード名一覧と集計値から親子の合計を検証します。
配布ブックと配布DLLにQPC計測機能・追加テストコードは含めません。
