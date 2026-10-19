# INCONSISTENTのDLL高速化（2026-10-20）

配布ビルドは `20261020_INCO_NATIVE_01`。対象はWindowsの64bit Excelです。
[実行ブック](../workbook/inco_native_20261020/2DSoilFEM_inco_native_20261020.xlsm) と
[計算DLL](../workbook/inco_native_20261020/soilfem_native_x64.dll)、
[材料DLL](../workbook/inco_native_20261020/soilfem_mat_x64.dll) を同じフォルダに置いてください。
保存済み結果は元ブックのものなので、入力を確認して再解析してください。

## 何を維持し、何を変更したか

Double、材料則、Newton・線形求解の許容値、反復上限、直線探索、カットバック、SRM探索条件を維持しました。
「高速化設定」の「DLLの演算方式（倍精度）」は `AUTO` が配布保存値です。
対応CPUとOSではAVX2/FMA、それ以外ではSSE2を選びます。SSE2を直接指定することもできます。
加算の集約順序とFMAによる丸め差を許容する設定です。Singleへの変更や収束条件の緩和はありません。
`STRICT` は従来のx87経路を使い、今回のGMRES経路を無効にします。
材料DLLは従来バイナリをそのまま同梱しました。

以前の帯LU SIMD試作に加え、次を実装しました。

- ILU(0)前処理付きGMRESを倍精度SSE2/AVX2へ移植。二重の修正Gram–Schmidt、真の残差での受理、再始動、失敗時の戻り値・イベントを維持。
- ソート済みCSRの対角位置で前進・後退代入の走査範囲を限定。CSRの整列、対角、配列寸法と作業領域を呼び出し前に検査。
- 確定材料状態の9フィールドをまとめて復元。計算を行わずバイトコピーし、塑性点フラグとカウントは元処理と同じ規則で更新。

新規エクスポートがないDLL、対象外の配列形状、`NATIVE_DLL=OFF/BAND/MAT` では既存の経路へ戻ります。
新機能用の内部KEY `NATIVE_STATE_COPY`、`NATIVE_GMRES_SIMD` は配布ブックでは1です。
KEY欠落時は0なので、旧設定を持つブックへのソース取り込みだけでは有効になりません。
比較・切り分け時には各KEYを0にできます。演算方式のKEY欠落時も従来どおり `STRICT` です。

ログ出力フォルダのキャッシュも調査しましたが、書き込み失敗時の挙動変更を含むため採用しませんでした。
本版の `FEMIo.bas` は基準コミットと同じです。

## 処理時間の内訳

基準は `abf037bf557d11d119426f44b1c0463691ebf82b` からの帯LU SIMD版。
まず一時コピーに階層型 `QueryPerformanceCounter` を注入しました。
表は子の時間を差し引いたexclusive時間です。親子のinclusive時間を足していません。
exclusive合計が解析ルートの実時間と一致することを検査しています。

初回の36手続き計測（Case1、加速off、INCONSISTENT、SRM全探索、AUTO）:

|処理|時間|全体の割合|
|---|---:|---:|
|CSR求解の未分類部分（主にGMRES。次の計測で分離）|84.449 s|53.56%|
|出力フォルダの解決 `FEMIoOutDir`|16.936 s|10.74%|
|確定材料状態の復元|10.244 s|6.50%|
|全体接線組立（子を除く）|8.763 s|5.56%|
|材料点更新 `FEMMatEval`|7.680 s|4.87%|
|解析全体|157.665 s|100%|

この結果に従って、最大のCSR/GMRES経路を最優先に改善しました。
GMRES・ILU境界を追加した39手続き計測の結果は、下記の生データと集計を参照してください。
QPC計測を入れたブックは内訳調査専用で、速度比較には計測を入れていないブックを使います。

39手続きの同一計測ブックで、新機能OFF/ONを比較したexclusive時間:

|処理|OFF|ON|OFFの全体比|
|---|---:|---:|---:|
|GMRES DLL呼び出し境界（830回）|81.331 s|10.922 s|51.44%|
|確定材料状態の復元（45,600回）|10.383 s|1.427 s|6.57%|
|出力フォルダの解決|17.079 s|16.108 s|10.80%|
|全体接線組立（子を除く）|8.885 s|8.967 s|5.62%|
|材料点更新（子を除く）|7.593 s|7.631 s|4.80%|
|解析ルート全体|158.098 s|76.303 s|100%|

GMRES部分は約7.45倍、状態復元は約7.28倍になりました。
出力フォルダの呼び出し回数は再始動イベントの減少に伴って117,562→109,884回になりましたが、
ログ処理そのものは変更していません。残る最大項目はI/O関連です。

## 同条件の速度・精度比較

|比較項目|基準（新機能OFF）|高速化ON|
|---|---:|---:|
|3回の中央値|155.811 s|70.548 s|
|実時間の範囲|151.700–176.088 s|69.198–71.978 s|
|Fs下端／上端|2.2000 / 2.2125|同一|
|区間中点／幅|2.20625 / 0.0125|同一|
|探索回数|17|17|
|最終自由残差|6.5595364013624e-13|同一|
|全170成分の最大変位差|—|0（全3組）|
|成功増分／分解／再利用／塑性点|7 / 6713 / 4078 / 76|同一|

中央値同士で **2.209倍、54.72%短縮**。全3組で高速化しました。
実行順はOFF→ON、ON→OFF、OFF→ONです。
第1組はFs表示・中点を比較し、第2・3組では丸め前のFs下端・上端・幅も直接取得して完全一致を確認しました。
QPCの内訳調査と比較条件を一致させ、上記時間に計測コードのオーバーヘッドを含めていません。

別途、従来のSTRICT/x87結果と最終配布AUTOを比較しました。
Fsと主要回数は一致し、全変位の最大絶対差は `3.0984e-12`、最大変位に対する比は `2.3363e-12`。
最終自由残差はSTRICT `7.1557e-13`、AUTO `6.5595e-13` でした。
こちらは丸め差を含む比較です。3組の同一AUTO比較で差が0だったこととは区別しています。
最終ソースを埋め込み直した配布ブックも全探索を再実行し、75.507秒、同じ変位・残差・丸め前Fs区間・主要回数でした。

GMRES内部の数値経路まで完全一致するという意味ではありません。
基準の再始動イベントは106,656件、最終配布版では98,978件でした。
加算順序・FMAの丸め差によって内側の推定残差と再始動タイミングが変わります。
真の残差による受理判定、許容値、反復上限は同じです。
この比較では外側の解・Fs区間・残差・主要回数が一致しました。

測定環境はAMD Ryzen 5 5600X、Windows 11 Home、64bit Excel `16.0.20430.20182`。
各解析の前後でExcelプロセスが測定用の1つだけであることを確認しました。
最終比較は同じブック・同じDLL・AUTOで新機能2個だけを0/1にし、3組を交互の順序で実行します。
全体時間は高分解能の `.NET Stopwatch` で測定し、従来のVBA `Timer` 値も保存します。
UI描画・I/Oも含むエンドツーエンドの時間です。

入力は追跡済みの `tests/fixtures/smoke_mesh.json`（85節点、20要素）です。
Case1は座標を0.25倍、c=5、φ=10°。残りの入力とSRM条件は比較ブックで共通です。
変位の比較ゲートは `max|u_new-u_ref| <= 1e-12 + 1e-9*max|u_ref|`。
これは検証用の比較基準であり、解析の収束許容値を書き換えるものではありません。
各自由度の変位、最終自由残差、Fs、成否、主要回数をJSONに保存します。

この小モデルでの改善率を、すべてのモデルやCPUへ一般化しません。
SIMDの丸め差は収束境界の近傍で経路の違いに発展する可能性があるため、案件ごとの比較を続けてください。
SRMの表示区間は数値的な収束境界であり、物理破壊や設計安全率を確定するものではありません。
既存の標準版や過去の検証記録は保持しています。

## 検証記録と再現

[生データ・集計・環境情報](native-evidence/inco_20261020/) を同梱します。
プロファイル、全変位を含む計測JSON、単体試験、実VBA構造体との統合試験、回帰結果を参照してください。
実行用ブックには計測コードを埋め込んでいません。

Case1〜4 × off/adaptive × INCONSISTENT/DAVISの16固定Fs比較と、同梱306要素モデルの固定Fs=1も、
全変位・丸め前Fs情報・最終残差・主要回数が一致しました。
状態復元のC試験19条件と、実VBA構造体の30比較、GMRESのSSE2/AUTO・再始動・複数許容値・
ゼロ右辺・非有限値・特異前処理・無効配列の試験を通過しています。
旧DLL、設定欠落、STRICTへの切り替えを含むGMRES経路の10比較も通過しました。
大規模な文献ケースの全探索を今回すべて再実行したものではありません。

64bit Excelの「VBAプロジェクト オブジェクトモデルへのアクセスを信頼する」を有効にし、
手元の作業を保存してExcelを閉じてから実行します。スクリプトは他のExcelが存在すると解析を開始しません。
Windows PowerShell 7、Python 3、Visual Studio 2022 C++ Build Toolsを使用します。

```powershell
pwsh -NoProfile -File tests/native/build_simd.ps1 -OutputDirectory tests/tmp/inco_rebuild -Inconsistent
python tests/native/test_simd.py --dll tests/tmp/inco_rebuild/soilfem_native_x64.dll --baseline workbook/soilfem_native_x64.dll
python tests/native/test_native_state.py --dll tests/tmp/inco_rebuild/soilfem_native_x64.dll
python tests/native/test_gmres_simd.py --dll tests/tmp/inco_rebuild/soilfem_native_x64.dll
pwsh -NoProfile -File tests/native/check_state_restore.ps1 -SourceWorkbook tests/tmp/inco_rebuild/2DSoilFEM_inco_native_20261020.xlsm -OutputRoot tests/tmp/state_check
pwsh -NoProfile -File tests/native/run_inco_benchmark.ps1 -Book tests/tmp/inco_rebuild/2DSoilFEM_inco_native_20261020.xlsm -OutputDirectory tests/results/inco_repeated
python tests/native/summarize_inco_benchmark.py tests/results/inco_repeated --out tests/results/inco_summary.json
```

出力先が既にある場合は上書きせず、新しい名前を使用します。
ペア計測の中断後に完了済みJSONを保持して続ける場合だけ `-Resume` を使います。
手続き別の再計測は `instrument_profile.py --src src/vba --dst <新規フォルダ>`、
`make_native_workbook.ps1 -Modules FEMIncoProfile,FEM` で計測モジュールを取り込み、
`run_native_compare.ps1 -Profile -Suite smokesrm -CaseName Case1 -Variants off -Modes AUTO -Flow INCONSISTENT -Precision AUTO` を使います。
`-StateCopy 0/1 -GmresSimd 0/1` で比較し、CSVを `summarize_profile.py` で検査します。
