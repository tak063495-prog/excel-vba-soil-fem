# INCONSISTENT 要素行列積・ILU構築の追加検証（2026-10-20）

前版 `6292076`（PR #16）との比較です。ビルド `20261020_INCO_ELEMENT_01`。
[実行ブックとDLL2個の一括ZIP](../workbook/SoilFEM_INCONSISTENT_ELEMENT_20261020.zip)。

## 変更

- 材料DLLの接線×B、および要素接線の3項内積を固定長へ展開。先頭の0加算を含め、項の順序と各加算後のDoubleへの丸めを維持します。
- ILU(0)構築で、各上三角成分の列位置を毎回行頭から探す処理を、単調に進む2本の走査へ変更。行内が厳密な昇順かを一度確認し、未整列・重複列は旧検索を使用します。数値更新の順序・ピボット閾値・失敗情報は維持します。
- GCCの最適化をO3へ変更。`-mfpmath=387 -fexcess-precision=standard -ffp-contract=off -fno-fast-math -fno-builtin`、MSVCR100数学関数、no-CRTを維持。AVX/FMAや特定CPU向けのビルド指定は追加していません。

GMRES本体・真残差判定・直交化回数・反復上限・収束条件・材料則・SRM探索条件は変更していません。
両側とも前版のPREPARED=1と同じnative DLLを使います。今回の追加変更によるGMRES途中値の差も認めず、全イベントCSVの一致を確認します。

## 配布候補の全探索

|組|前版（秒）|候補（秒）|短縮率|
|---|---:|---:|---:|
|1|44.582079|40.452694|9.26%|
|2|40.933868|40.024546|2.22%|
|3|39.508724|38.906376|1.52%|
|4|40.802164|45.962986|-12.65%|
|5|44.695913|41.377589|7.42%|
|中央値|40.933868|40.452694|**1.18%**|

5組中4組で短縮しました。追加10%の達成とは扱いません。

同じコンパイル済みブックのコピーを使い、材料DLLだけを前版/候補へ入れ替えました。
INCONSISTENT、Case1/off、20要素・85節点、精度AUTO、既存の高速化設定とPREPAREDは両側1。
準備後の解析開始から終了までをStopwatchで計測。各実行は独立Excel、順序は旧→新、新→旧、旧→新、新→旧、旧→新です。最初の3組の後、回帰検証を挟んで追加2組を計測し、全5組を集計しました。
元ブックは変更せず、既存Excelが起動中ならテストを拒否します。

全5組でFs区間 **[2.200000000000001, 2.2125000000000012]**、中央値2.206250000000001、探索17回が完全一致。
全170変位の最大絶対差は0、最終残差は両側とも **6.5595364013624e-13**。
成功増分7・分解6,713・再利用4,078・塑性点76も一致しました。
GMRES途中値を含む全イベントCSV **13,415,684バイト**が各組で完全一致し、実CSVを独立に読み直して検証しています。

今回の比較はこのモデル・このPCでの結果であり、他のモデルやPCで同じ短縮率を保証しません。
各回のばらつきも表へ残しています。
全体の差は測定変動に比べて小さく、1組では遅くなりました。1.18%は今回の中央値の観測値で、統計的に確定した短縮率としては扱いません。

## 内訳と不採用・探索段階の結果

QPCによる材料DLL内部の排他時間です。数学関数を細かく計測する前版プロファイラとは計測範囲が異なります。
親と子の時間は二重加算しません。QPC呼出やインライン抑止の影響を含むため、全体の短縮率には使いません。

|処理|回数（両側同じ）|旧・排他秒|新・排他秒|
|---|---:|---:|---:|
|sfm_eval|56,583|0.516063|0.527806|
|sfm_tangent|7,803|0.897907|0.809463|
|strain|4,469,142|0.812384|0.716386|
|mp_update|4,469,142|3.072667|2.961996|
|state_display|4,468,643|0.774241|0.772144|
|stress_B|4,468,643|0.861757|0.537251|
|internal_force|1,122,226|0.892892|0.729163|
|commit|1,122,226|0.232670|0.135112|
|sfm_ilu0|830|0.146045|0.063192|
|sfm_regularized|1,752|2.358765|2.341842|


初期のO2版（行列積の展開＋ILU検索変更）も3組測定しました。
中央値41.869533秒→42.374825秒で1.21%遅く、全体高速化を確認できませんでした。
ILU検索変更だけの単発も40.362524秒→41.603564秒で3.07%遅く、この単発を改善の根拠には使いません。
これらも生データへ残し、配布候補のO3版の成績へ混ぜていません。

ILU構築単体は、人工帯行列169未知数/帯幅8と676未知数/帯幅20で計測しました。
169未知数: 41.361→30.786µs（25.57%短縮） / 676未知数: 1755.424→652.212µs（62.85%短縮）。
準備済みバッファで直接DLLを呼び、割当・出力シリアライズは計測から除外。これはExcel全体の改善率ではありません。

GMRESの残差ベクトル計算とノルム集計を1回の走査へまとめる案も、別のスカラー試験器で比較しました。
169未知数: 30.53%のカーネル単体短縮 / 676未知数: 22.67%のカーネル単体短縮でした。いずれも2つの人工行列での探索試験です。
前版内部計測では真残差処理全体が約0.234秒で、全体時間への寄与は小さいため、今回は製品GMRESへ組み込みません。
2回のMGSをまとめる案や加算順序を変える案、反復省略、収束条件を緩める案も採用していません。
既存の行列分解再利用は維持し、行列値が変わった後の因子を新たに使い回す変更はしていません。

## 数値・コンパイル確認

- 材料点20万件: x87制御語0x027f/0x037fで、成功点の全200バイト・処理可否・FPエラー判定が一致。
- 要素処理5,200ケース: 両制御語で要素全体・内力・塑性点フラグ・診断・done配列が完全一致。ゼロ増分、接合、未対応材料、サブステップ、NaN、オーバーフロー、範囲外DOF、非活動要素、hourglass、接線のみも含みます。
- ILU構築2,024ケース: 両制御語で戻り値・因子全体・失敗行/列・最小ピボットが完全一致。別途、配布用DLL同士1,012ケースも一致。
- 材料DLLのx87 GMRES 480ケース: 両制御語で解・作業領域・イベントが完全一致し、成功時の真残差を独立に確認。
- 正則化回復: 独立28行列×4経路の112照合、8容量境界、メモリ不足時の代替、ユーザー中断の伝播を確認。
- モデル別回帰17組: Fs・全変位・最終残差・主要回数・CSV全体が完全一致。
- 比較器の否定試験、Python構文、PowerShell構文、配布物の内容・ハッシュを照合。

材料DLLはGCC 14.2.0、警告をエラーとしてコンパイル。配布native DLLは前版と同一です。
VBAは実Excelでコンパイルし、保存・開き直し後の8モジュールをソースと照合しました。
FEMCoreの変更はビルド番号だけです。計測器・制御語変更用プローブは配布DLLへ含めません。

固定Fsの回帰17組です。短い解析で測定変動があるため、この表は全体高速化の根拠には使いません。これらはGMRES呼出0です。

|流れ則|モデル|条件|旧秒|新秒|最終残差（同値）|
|---|---|---|---:|---:|---:|
|INCONSISTENT|Case1|off|2.164|2.131|9.443509e-16|
|INCONSISTENT|Case1|adaptive|1.592|1.573|5.447052e-15|
|INCONSISTENT|Case2|off|1.591|1.570|1.160747e-13|
|INCONSISTENT|Case2|adaptive|1.606|1.576|1.250442e-13|
|INCONSISTENT|Case3|off|1.611|1.623|4.729932e-12|
|INCONSISTENT|Case3|adaptive|1.575|1.575|3.834779e-11|
|INCONSISTENT|Case4|off|1.578|1.647|7.543433e-10|
|INCONSISTENT|Case4|adaptive|1.574|1.561|9.042951e-10|
|DAVIS|Case1|off|2.174|2.314|1.045686e-15|
|DAVIS|Case1|adaptive|1.575|1.682|4.225099e-15|
|DAVIS|Case2|off|1.576|1.573|2.716041e-09|
|DAVIS|Case2|adaptive|1.562|1.559|5.629346e-08|
|DAVIS|Case3|off|1.590|1.565|1.460302e-09|
|DAVIS|Case3|adaptive|1.576|1.570|5.356740e-10|
|DAVIS|Case4|off|1.555|1.570|1.668341e-10|
|DAVIS|Case4|adaptive|1.664|1.569|1.232036e-10|
|INCONSISTENT|builtin|default|3.043|2.780|1.055719e-13|

同梱306要素モデルの全2,010変位も最大絶対差0です。

生データ・環境・SHA-256・コンパイル記録は [検証記録](native-evidence/element_20261020/README.md) に保存します。

## 再現

64bit Windows Excel、MSVC、GCC 14.2.0、Python、PowerShellを使用します。
VBAプロジェクトへのアクセスを信頼する設定が必要です。既存Excelを閉じて実行してください。

```powershell
pwsh -NoProfile -File tests/native/build_material.ps1 -Compiler C:/path/to/gcc.exe -Optimization O3 -OutputDirectory tests/tmp/element_rebuild_new
pwsh -NoProfile -File tests/native/build_element.ps1 -MaterialDll tests/tmp/element_rebuild_new/soilfem_mat_x64.dll -MaterialBuildRecord tests/tmp/element_rebuild_new/build.json -OutputDirectory tests/tmp/element_book_new
# referenceフォルダへ同じコンパイル済みブックとnative DLL、前版の材料DLLを配置します。
pwsh -NoProfile -File tests/native/run_element_benchmark.ps1 -ReferenceBook C:/path/to/reference.xlsm -CandidateBook C:/path/to/candidate.xlsm -OutputDirectory tests/results/element_pairs_new -Pairs 5
python tests/native/summarize_element.py tests/results/element_pairs_new --pairs 5 --out tests/results/element_pairs_new/summary.json
pwsh -NoProfile -File tests/native/run_element_benchmark.ps1 -ReferenceBook C:/path/to/reference.xlsm -CandidateBook C:/path/to/candidate.xlsm -OutputDirectory tests/results/element_regression_new -Regression
python tests/native/summarize_element.py tests/results/element_regression_new --regression --out tests/results/element_regression_new/summary.json
pwsh -NoProfile -File tests/native/check_packed48.ps1 -SourceWorkbook C:/path/to/candidate.xlsm -OutputRoot tests/results/element_recovery_new
```

内部計測用ソースは `instrument_element.py --src native/soilfem_mat.c --out tests/tmp/element_profile_new/material.c` で作成します。
同じ最適化・浮動小数点設定で別DLLを作り、`run_native_compare.ps1 -NativeProfile -MaterialProfile` で取得します。
`summarize_element_profile.py` は材料評価・接線・ILU構築・正則化回復の4ルートと子処理の排他和を照合します。
