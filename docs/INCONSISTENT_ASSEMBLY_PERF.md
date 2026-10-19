# INCONSISTENTの行列組立高速化（2026-10-20）

前版 `87d4f13` の行列組立をQPCで分解しました。
Double・収束条件・反復上限・材料則・SRM探索条件は変更していません。
ビルドは `20261020_INCO_ASSEMBLY_01`。
[XLSMとDLL2個の一括ZIP](../workbook/SoilFEM_INCONSISTENT_ASSEMBLY_20261020.zip) を展開し、同じフォルダに置いてください。
保存済み解析結果は元ブックのものです。入力を確認して再解析してください。

## 内訳と変更

小モデルSRM全探索の61項目のQPC計測では、行列組立は7,804回、9.521秒でした。
主な内訳は次のとおりです。親手続きと子手続きの時間を重複加算していません。

|処理|秒|
|---|---:|
|CSR散布表の作成|4.913|
|要素行列の対称性検査|2.395|
|CSR値のコピー|0.484|
|メッシュ・RCM方針の確認|0.474|
|CSR値のゼロ初期化|0.363|
|解法選択|0.241|

表は主な処理の抜粋です。このほかCSR対角検査、ワーク領域の確認、制御処理などが含まれます。

最大の散布表作成を計算DLLの `sfn_csr_scatter_list` へ移しました。
接線行列が変わると非ゼロの位置も変わるため、幾何形状だけで表を使い回す変更はしていません。
従来どおり、その時点の全要素の16×16行列から非ゼロ項を列挙します。
要素→局所行→局所列の順序、CSR位置・値インデックスを維持します。
非活動要素も散布表の作成対象に含め、有効性の適用は既存の数値組立側で行います。

DLLはbinary64のビット表現でゼロと非有限値を判定します。
正負のゼロを除外し、最小の非正規化数も非ゼロとして保持します。
この処理に浮動小数点の加減乗除はなく、実際の行列加算や対称性検査は既存のままです。
最初に件数・配置を検証し、VBAが必要数の配列を確保した後で表を作成します。
容量不足・非有限値・位置異常などの拒否時はDLLの出力を書き換えません。
元の世代番号・要素数・パターン署名によるキャッシュ再利用条件も維持します。

`NATIVE_ASSEMBLY=1`、`NATIVE_DLL=AUTO`で使用します。
設定欠落・0・旧DLLでのエクスポート欠落・不適合な配置では従来のVBA経路へ戻ります。
配布ブックは1を保存し、開き直して確認しています。材料DLLは前版と同一です。

## 全探索の比較

|組|従来VBA（秒）|DLL散布表（秒）|
|---|---:|---:|
|1|53.718|46.924|
|2|53.559|42.288|
|3|54.079|45.298|
|中央値|53.718|45.298|

中央値で **15.67%短縮（1.186倍）**。全3組で短縮しました。
基準の範囲は53.559〜54.079秒、高速側は42.288〜46.924秒です。

同一ブック・同一DLLで `NATIVE_ASSEMBLY` だけを0/1に切り替えます。
`IO_FAST_PATH=EVENT_LOG_FAST=NATIVE_STATE_COPY=NATIVE_GMRES_SIMD=1`、精度AUTO、
INCONSISTENT、Case1/offの全探索です。順序は0→1、1→0、0→1。
入力設定の準備後の解析全体を高精度Stopwatchで測定し、各回は独立したExcelで実行します。
QPC計測用ブックの時間は、この速度比較へ混ぜていません。

|比較項目|両経路の結果|
|---|---|
|Fs下端 / 上端|2.200000000000001 / 2.2125000000000012|
|Fs中点 / 幅|2.206250000000001 / 0.012500000000000178|
|試行数|17|
|全170変位の最大差|0|
|最終自由残差|6.5595364013624e-13|
|成功増分 / 分解 / 再利用 / 塑性点|7 / 6713 / 4078 / 76|
|散布表の最終件数|5,120（全配列の順序も別試験で一致）|
|最終行列対称性誤差|8.588448757284068e-17|
|DLL散布表作成回数|0 → 7,804|
|イベントログ|13,541,361バイト、全バイト一致|

イベントログのSHA-256は `471C64B436C1A307197CA01908F232F129FAD500364C299860319DBE8452BB34` です。

## 内訳の再計測と回帰確認

同一の62項目の計測ブックで、設定0/1を比較しました。

|処理|従来（秒）|高速側（秒）|
|---|---:|---:|
|SetTotalMat|9.181|5.364|
|P6EnsureCSRScatterList|4.937|0.528|
|FEMNativeCSRScatterList|0.063|0.494|
|AssemblySymmetryScan|2.430|2.433|
|AssemblyZeroValues|0.369|0.363|
|AssemblyCopyValues|0.489|0.484|
|P6SelectSolverMode|0.208|0.632|

上2行は子手続きを含む時間です。表の行を合算しません。散布表作成は89.30%、行列組立全体は41.57%短縮しました。
計測ブック全体は54.328秒→52.113秒で、他の処理・I/O時間にも変動がありました。計測負担を含むこの1組を、全探索3組の中央値へ混ぜていません。
Fs・全変位・残差・主要回数・イベントCSVはこの計測対でも完全一致し、QPCの排他的時間の総和は全体時間に一致しました。

- DLLはMSVC `/O2 /fp:strict /W4 /WX` で警告をエラーとしてビルド成功。
- 実際の64bit ExcelでVBAをコンパイルし、保存前と開き直し後に必須8モジュールをソース照合。計測用の62項目版もコンパイル成功。
- 実VBA配列・要素UDTの全バイト比較は、新DLL54条件・旧DLL54条件で一致。要素数1/2/20、ゼロ/疎/密行列、非活動要素、設定欠落/0/1、AUTO/OFF、キャッシュ再利用を確認。
- ctypesの独立試験3メソッド（24入力構成、26の無効パラメータ・NULL要求、両パスの非有限値・無効CSR位置）を通過。非対称配置、正負ゼロ、最小の非正規化数、容量不足、件数がNNZより多い配置、拒否時の出力保持・入力不変を検査。
- 新DLLの既存計算も確認し、STRICTの因子・右辺・解は前版DLLと全バイト一致。材料状態コピー19ケース、GMRESの解・残差・失敗時の出力保持、比較器の改変拒否試験も通過。
- 前回公開版との比較でFs・全変位・残差・主要回数が一致し、ビルド番号列を除く全109,847イベントも一致。

固定FsのCase1〜4×off/adaptive×INCONSISTENT/DAVISで16組、同梱306要素モデルで1組を比較しました。
全組でFs・全変位・残差・主要回数・ログが完全一致。DAVISではCSR経路を使わず、新DLL散布表の呼出回数0も確認します。

|流れ則・モデル|条件|従来（秒）|高速側（秒）|残差（両者同値）|
|---|---|---:|---:|---:|
|INCONSISTENT|Case1/off|2.257|2.403|9.44351e-16|
|INCONSISTENT|Case1/adaptive|1.603|1.707|5.44705e-15|
|INCONSISTENT|Case2/off|1.685|1.743|1.16075e-13|
|INCONSISTENT|Case2/adaptive|1.580|1.699|1.25044e-13|
|INCONSISTENT|Case3/off|1.657|1.653|4.72993e-12|
|INCONSISTENT|Case3/adaptive|1.640|1.697|3.83478e-11|
|INCONSISTENT|Case4/off|1.610|1.678|7.54343e-10|
|INCONSISTENT|Case4/adaptive|1.595|1.800|9.04295e-10|
|DAVIS|Case1/off|2.146|2.141|1.04569e-15|
|DAVIS|Case1/adaptive|1.628|1.571|4.22510e-15|
|DAVIS|Case2/off|1.619|1.578|2.71604e-09|
|DAVIS|Case2/adaptive|1.649|1.603|5.62935e-08|
|DAVIS|Case3/off|1.668|1.627|1.46030e-09|
|DAVIS|Case3/adaptive|1.622|1.613|5.35674e-10|
|DAVIS|Case4/off|1.591|1.569|1.66834e-10|
|DAVIS|Case4/adaptive|1.596|1.579|1.23204e-10|
|builtin|builtin/default|2.854|3.196|1.05572e-13|

**短い固定Fs解析では短縮を確認できない条件があります。** 同梱306要素の単発測定は2.854秒→3.196秒で、速くなっていません。全2,010変位の最大差は0、残差は1.05571945918321e-13です。
全探索の短縮率をこの固定Fs解析へ適用しません。各回の時間変動を含めて記録しています。

小モデルと同梱モデルの確認範囲の結果です。任意のメッシュ・材料・施工履歴で同じ短縮率になることを保証するものではありません。
コード・配列比較・全変位を含む生データは [検証記録](native-evidence/assembly_20261020/README.md) に収録します。

## 再現手順

64bit Windows Excel、Visual Studioのx64 C/C++ビルド環境、Python、PowerShellを使用します。
テストは他のExcelが起動中なら開始を拒否し、コピーしたブックと自分が起動したExcelだけを扱います。
VBAプロジェクトへのアクセスを信頼する設定が必要です。
既存の出力は上書きしないため、新しい出力先を指定してください。

```powershell
cmd /c native\build.cmd
pwsh -NoProfile -File tests/native/build_assembly.ps1 -OutputDirectory "$PWD/tests/tmp/assembly_rebuild"
python -m unittest discover -s tests/native -p test_assembly.py -v
$book="$PWD/tests/tmp/assembly_rebuild/2DSoilFEM_inco_assembly_20261020.xlsm"
pwsh -NoProfile -File tests/native/check_assembly.ps1 -Book $book -OutputDirectory "$PWD/tests/results/assembly_arrays_new"
pwsh -NoProfile -File tests/native/check_assembly.ps1 -Book $book -OutputDirectory "$PWD/tests/results/assembly_legacy_new" -NativeDll "$PWD/workbook/inco_event_20261020/soilfem_native_x64.dll" -LegacyDll
pwsh -NoProfile -File tests/native/run_inco_benchmark.ps1 -Book $book -OutputDirectory "$PWD/tests/results/assembly_pairs_new" -AssemblyOnly -Pairs 3
python tests/native/summarize_io_benchmark.py tests/results/assembly_pairs_new --scope assembly --out tests/results/assembly_pairs_new/summary.json
pwsh -NoProfile -File tests/native/run_assembly_regression.ps1 -Book $book -OutputDirectory "$PWD/tests/results/assembly_regression_new"
```

`instrument_profile.py --assembly-sections` は使い捨てソースにメッシュ方針・対称性・ゼロ初期化・コピーと関連手続きのQPC境界を追加します。
`make_native_workbook.ps1 -Modules FEMIncoProfile` で計測ブックを構築し、
`run_native_compare.ps1 -Profile -AssemblyFast 0` / `1` と同じ条件を指定して実行します。
`summarize_profile.py` は全体時間と排他的時間の総和が一致することを検査します。
