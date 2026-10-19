# INCONSISTENTのILU前処理・材料評価内部の改善（2026-10-20）

前版 `b2014b7`（PR #15）を基点とする追加改善です。ビルドは `20261020_INCO_ILUMAT_01`。
[実行ブックとDLL2個の一括ZIP](../workbook/SoilFEM_INCONSISTENT_ILUMAT_20261020.zip) を展開し、3個を同じフォルダに置いてください。
保存済み結果は元ブックのものです。入力を確認して再解析してください。

## 採用した変更

ILUは1回のGMRES中に同じ因子を繰り返し使います。対角逆数を1回だけDouble配列へ準備し、後退代入の毎回の除算を乗算にしました。
CSRの形状・添字を検証してから準備します。ピボット閾値 `1e-12` は各適用でも従来と同じです。
真残差は従来のスカラーCSR計算で確認し、2回の修正Gram–Schmidt、許容値、再スタート幅、反復上限、成功時だけ解を出力する契約を維持します。
確保した逆数配列は成功・失敗とも解放し、メモリ確保失敗時は前版のSIMD経路へ戻ります。

`NATIVE_GMRES_PREPARED=1` は `NATIVE_GMRES_SIMD=1`、`NATIVE_GMRES_TUNED=1`、AUTO/SSE2で有効です。
キー欠落・0・新エクスポートがない旧DLLでは前版GMRESへ戻ります。STRICTは従来どおりx87参照経路です。
配布ブックには新キー1を保存し、開き直して確認しました。

材料DLLでは次の同一値の再計算だけを削減しました。

- 塑性分岐で、既にDoubleへ格納した試行応力・降伏判定・角度を再利用。
- 同じ角度の `sin` / `cos` を接線計算へ渡して再利用。
- 同じ `pow(cos,2)` / `pow(sin,2)` の結果を再利用。

x87、Doubleへの丸め、MSVCR100の数学関数、材料則、接線の有限値確認、decline判定を維持します。
材料点20万件をx87制御語0x027f/0x037fで比較し、処理した点の全出力200バイト、処理可否、浮動小数点エラー判定が一致しました。
全体としてDouble、非線形収束条件、SRM探索条件、ログ出力の内容・再試行条件は変更していません。

## 配布版の全探索比較

|組|前版条件（秒）|新条件（秒）|短縮率|
|---|---:|---:|---:|
|1|47.050575|46.582910|0.99%|
|2|42.395004|40.212163|5.15%|
|3|43.551804|41.559898|4.57%|
|中央値|43.551804|41.559898|**4.57%**|

3組とも短縮しました。今回の追加改善は4.57%であり、追加10%の達成とは扱いません。

Case1/off、20要素・85節点、INCONSISTENT、精度AUTO。同一のコンパイル済みブック・同一native DLLを両側で使用し、
前版材料DLL＋PREPARED=0と、新材料DLL＋PREPARED=1を比較しました。他の既存高速化設定はすべて1です。
解析入力の準備後から終了までを高精度Stopwatchで測定し、各回は独立したExcelです。順序は旧→新、新→旧、旧→新。
計測用DLLや初期単発結果はこの3組へ含めません。実行時間には変動があり、任意のモデルで同率の短縮を保証しません。

全3組でFs下限2.200000000000001、上限2.2125000000000012、中央値2.206250000000001、探索17回が完全一致。
全170変位の最大絶対差は0、最終残差は両側とも **6.5595364013624e-13** でした。
成功増分7、分解6,713、再利用4,078、塑性点76、材料decline499、fallback=0も一致。
新条件の全830回で新エクスポートを使用し、別の内部計測でもILU prepared適用1,660,000回を確認しました。
GMRESイベント行は99,048→98,851へ変わります。
それ以外の10,869行は、失敗線形解法のlinear_relres表示だけを正規化すると、3組とも全バイト一致しました。

ILU逆数乗算では途中の丸めが変わるため、GMRESの途中残差・再スタート回数・失敗線形解法の `linear_relres` 表示には差があります。
それ以外のステージ・Fs・増分・イベント・詳細は実CSVを独立に比較しています。
この除外は検証上だけで、配布版のログを省略・改変するものではありません。最終残差・Fs・全変位・主要回数は別途完全一致を要求します。

## 内部計測

ILUの独立計測は現行SIMDを基準とし、逆数の準備費用を反復費用から分離しました。
各10,000反復・5標本の中央値で、169未知数・帯域7では14.66%、676未知数・帯域20では12.46%短縮。
独立した `L*(U*x)` の相対残差は約2.6e-16以下、解の差は約3.9e-16以下でした。
これは人工因子のマイクロ計測で、Excel全体の短縮率には使用していません。

|処理|旧条件（秒）|新条件（秒）|
|---|---:|---:|
|ILU適用|4.345300|3.756597|
|GMRES全体|11.882345|10.942343|
|材料評価＋要素接線の合計|12.412928|10.864871|

材料側の排他時間と回数です。親子の時間を重複加算しません。

|処理|旧回数|新回数|旧排他秒|新排他秒|
|---|---:|---:|---:|---:|
|sfm_eval|56,583|56,583|3.849800|3.885093|
|mp_update|4,469,142|4,469,142|0.542030|0.423976|
|spectral|4,213,541|4,213,541|2.670604|2.414509|
|eval_yield|12,895,750|8,682,209|1.047987|0.710546|
|c_pow|43,153,922|34,727,788|2.120227|1.712660|
|c_sqrt|22,387,201|18,173,660|0.460112|0.371240|
|c_atan|12,895,750|8,682,209|0.419280|0.260569|
|c_sin|8,426,583|4,213,541|0.198489|0.101489|
|c_cos|8,426,583|4,213,541|0.214367|0.103225|
|sfm_tangent|7,803|7,803|0.906493|0.897949|

材料ルート以外からの数学関数呼出は旧0.016460秒、新0.016386秒で、集計で分離しています。

階層QPCは使い捨てのnative DLL・材料DLLで同じ解析を計測したものです。
多数のQPC呼出とインライン化の抑止による費用を含むため、時間差を配布版の全体短縮秒数へ直接換算しません。
共有数学関数の計測には材料評価以外のDLLエクスポートからの呼出も含まれます。この時間は集計で分けています。

## 検証と採用しなかった案

- DLL: nativeはMSVC `/O2 /fp:strict /W4 /WX`、材料はGCC 14.2.0のx87・no-CRT・警告エラー設定でコンパイル成功。
- 材料DLLの無変更再ビルドも、前版配布DLLと全探索の数値・CSV全体が一致。
- 配布XLSM: VBA実コンパイル成功。保存・開き直し後の8モジュールSHA-256と設定を照合。
- GMRES prepared 87ケース・前版tuned 87ケース合格。独立真残差、既知解、非有限値、失敗時出力保護、広いL/U帯域を含みます。
- 旧GMRESエクスポート120ケースは、前版DLLと解・作業領域・イベントの全バイト一致。
- 新DLL・前版DLLでSIMD設定欠落/0、STRICT、SSE2、AUTOの各5条件を確認。
- 状態コピー19ケース、既存対称/非対称SIMD・STRICT参照、比較器の否定試験も合格。

固定Fs回帰17組の時間も省略せず記録します。全組でFs・全変位・最終残差・主要回数・CSV全体が一致しました。
これらはGMRES呼出0で、短時間の解析には測定変動もあるため、この表を高速化の根拠には使いません。

|流れ則・モデル|条件|旧（秒）|新（秒）|最終残差（両側同値）|
|---|---|---:|---:|---:|
|INCONSISTENT|Case1/off|2.564|2.609|9.44351e-16|
|INCONSISTENT|Case1/adaptive|1.987|1.983|5.44705e-15|
|INCONSISTENT|Case2/off|2.112|1.992|1.16075e-13|
|INCONSISTENT|Case2/adaptive|2.026|1.975|1.25044e-13|
|INCONSISTENT|Case3/off|1.985|2.009|4.72993e-12|
|INCONSISTENT|Case3/adaptive|1.980|1.981|3.83478e-11|
|INCONSISTENT|Case4/off|2.005|1.969|7.54343e-10|
|INCONSISTENT|Case4/adaptive|1.986|2.028|9.04295e-10|
|DAVIS|Case1/off|2.574|2.583|1.04569e-15|
|DAVIS|Case1/adaptive|1.957|1.958|4.22510e-15|
|DAVIS|Case2/off|1.977|1.950|2.71604e-09|
|DAVIS|Case2/adaptive|1.971|1.968|5.62935e-08|
|DAVIS|Case3/off|1.981|1.956|1.46030e-09|
|DAVIS|Case3/adaptive|1.999|2.018|5.35674e-10|
|DAVIS|Case4/off|1.987|2.252|1.66834e-10|
|DAVIS|Case4/adaptive|2.012|1.972|1.23204e-10|
|builtin|builtin/default|2.921|2.832|1.05572e-13|

同梱306要素モデルの全2,010変位も最大差0、最終残差1.05571945918321e-13です。

`pow(x,2)` をSSE2のDouble乗算へ変える案も別ソースで試しました。
20万点で処理可否・FPエラー判定の差は0でしたが、出力に丸め差が出ました。
主応力がほぼ等しいケースでは表示角度の差が最大約29.76度に拡大するため、今回の配布版には採用していません。
材料側は全出力が一致した重複処理の削減に限定しました。

GMRESの明らかな不成立をゼロ行から判定できるかも計測しました。両側の全830呼出で完全なゼロ行は0件で、この方法による反復省略はできませんでした。
反復を減らす変更や収束条件を緩める変更は採用していません。

生データ、配布物SHA-256、環境、実CSV、コンパイル記録は [検証記録](native-evidence/ilumat_20261020/README.md) に保存しています。

## 再現

64bit Windows Excel、MSVC x64、MinGW-w64 GCC、PowerShell、Pythonを使用します。
VBAプロジェクトへのアクセスを信頼する設定が必要です。他のExcelが起動中なら開始を拒否します。
元ブックを変更せず、テスト用コピーと自分が起動したExcelだけを扱います。出力先は新しいフォルダを指定してください。

```powershell
cmd /c native\build.cmd
New-Item -ItemType Directory -Force tests/results | Out-Null
# GCCのある絶対パスを指定。build_material.ps1はbuild_mat.shと同じx87/no-CRTオプションです。
pwsh -NoProfile -File tests/native/build_material.ps1 -Compiler C:/path/to/gcc.exe -OutputDirectory "$PWD/tests/tmp/mat_rebuild_new"
pwsh -NoProfile -File tests/native/build_ilumat.ps1 -MaterialDll "$PWD/tests/tmp/mat_rebuild_new/soilfem_mat_x64.dll" -OutputDirectory "$PWD/tests/tmp/ilumat_rebuild_new"
python tests/native/test_gmres_tuned.py --entry sfn_gmres_prepared --out tests/results/prepared_new.json -v
```

比較用フォルダへ同じ新ブック・同じnative DLLと、前版 `workbook/inco_gmres_20261020/soilfem_mat_x64.dll` を配置してください。

```powershell
pwsh -NoProfile -File tests/native/run_ilumat_benchmark.ps1 -ReferenceBook C:/path/to/reference.xlsm -CandidateBook C:/path/to/candidate.xlsm -OutputDirectory "$PWD/tests/results/ilumat_pairs_new" -Pairs 3
python tests/native/summarize_ilumat.py tests/results/ilumat_pairs_new --out tests/results/ilumat_pairs_new/summary.json
python tests/native/check_gmres_logs.py tests/results/ilumat_pairs_new --out tests/results/ilumat_pairs_new/log_check.json
pwsh -NoProfile -File tests/native/run_ilumat_regression.ps1 -ReferenceBook C:/path/to/reference.xlsm -CandidateBook C:/path/to/candidate.xlsm -OutputDirectory "$PWD/tests/results/ilumat_regression_new"
```

内部計測には `instrument_gmres.py` と `instrument_material.py` で使い捨てソースを作成します。
両方の計測DLLを同じフォルダへ置き、`run_native_compare.ps1 -NativeProfile -BothProfiles` で別々の集計配列を取得します。
`--diagnose` のnative配列29～31はゼロ行の追加診断です。配布DLL・配布ブックには計測器・点試験エクスポート・テストマクロを含めません。
