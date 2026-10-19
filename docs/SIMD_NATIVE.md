# 倍精度SIMD DLL（2026-10-19）

基準は main の `abf037b`（候補48＋DLL統合版）。試験版のビルド番号は
`20261019_SIMD_NATIVE_01` です。
[実行用ブック](../workbook/simd_20261019/2DSoilFEM_simd_20261019.xlsm) と
[帯行列DLL](../workbook/simd_20261019/soilfem_native_x64.dll)、
[材料DLL](../workbook/simd_20261019/soilfem_mat_x64.dll) を同じフォルダに置きます。
旧標準のブックとDLLは履歴として保持しています。

## 演算と精度の方針

配列・入出力は引き続き64bitのDoubleです。高速化は、x87の拡張精度による
VBAとのビット一致を任意に解除し、倍精度SSE2、AVX2・FMAを使うことで行います。
材料則、ピボットしきい値、反復改良、非線形収束許容値、SRMの探索条件は変更していません。
単精度化、fast-math、FTZ/DAZの変更は行いません。

「高速化設定」の「DLLの演算方式（倍精度）」（C29）で `NATIVE_PRECISION` を指定できます。
保存済みの旧版結果は、再解析してから利用してください。

|設定|実際の演算|
|---|---|
|STRICT|従来のx87演算。VBAとのビット一致を検査する基準|
|SSE2|倍精度SSE2。連続配列更新を2成分ずつ処理|
|AUTO|CPUとOSが対応すれば連続更新をAVX2・FMAで4成分ずつ処理。それ以外はSSE2|

新ブックの保存値はAUTO。KEYがない既存ブックを新ソースから再構築した場合はSTRICTです。
コード上のAVX2指定はAUTOと同じ自動フォールバックです。
ログに `precision=STRICT` / `SSE2` / `AVX2_FMA` が出ます。
`NATIVE_DLL=OFF` は引き続き純VBAです。

CPUIDでAVX、OSXSAVE、FMA、AVX2を確認し、XGETBVでOSのXMM/YMM保存対応を確認します。
AVX2部分だけ別オブジェクトを `/arch:AVX2` でコンパイルし、基本DLLはSSE2でビルドします。
非対応CPUでAVX命令を実行する構成にはしていません。ただしAVX非対応実機での試験は未実施です。

対象は `soilfem_native_x64.dll` の対称・非対称帯LUと前進・後退代入です。
連続帯行列の更新と内積をSIMD化し、非対称帯行列のストライド更新にはSSE2を使います。
このストライド更新はAVX2の集約・分散処理よりSSE2が速い実測に基づく選択です。
材料点・GMRES・正則化などを含む `soilfem_mat_x64.dll` は今回変更していません。

精度選択はスレッド単位です。新VBAは各呼出し前に選択を確認するので、新版ブック間で
演算設定が混ざるのを防ぎます。旧VBAとの同時利用では旧VBA側にこの確認がないため、
厳密な旧版比較はExcelを閉じて別セッションで行ってください。

ビット一致の解除による差を、一律に機械イプシロン以下と保証するものではありません。
悪条件の行列、降伏面・カットバック・限界Fsに近い判定では差が増幅されることがあります。
比較テストの暫定受入値は、変位ベクトルの最大成分を基準として
`max|u_fast-u_ref| <= 1e-12 + 1e-9*max|u_ref|`、成否とFs区間の一致です。
これは比較用のゲートであり、解析の収束許容値を緩めた値ではありません。

## 検証と測定

測定機はAMD Ryzen 5 5600X、Windows、64bit Excel。Excel試験は1本ずつ実行しました。
以下の全探索時間は各1回の測定で、異なるPCでの速度保証ではありません。
最終DAVIS測定のSTRICT実行中に別のExcelプロセスが起動したため、最後の時間は参考値です。
既に起動していたそのExcelは操作・終了していません。追加の小回帰は起動前の検査で停止しました。

最終配布版の同梱モデルのDAVIS・SRM全探索では、AUTO→STRICTの順に測定し、
STRICT 39.671875秒 → AUTO 31.6171875秒、約20.3%短縮（1.25倍）でした。
Fs区間2.2375～2.25、17試行、成功増分83、
分解1019、再利用893、塑性点697が一致しました。
全2010成分の変位最大差は `1.7764e-15`、最大変位で正規化した差は `6.5136e-15`。
最終相対残差は `4.3455692461e-6` / `4.3455692480e-6` でした。
調整前の初回測定（STRICT→AUTOの順）は38.71875秒 → 31.5390625秒でした。
初回と最終配布版を混ぜた平均や中央値は示していません。

最終DLLのカーネル試験は、短いベクトルと端数、ldの余白、行交換、前進再生、
不正モード・引数、特異行列、スレッド設定の分離、独立残差を検査します。STRICTの対称・非対称LUの
因子・右辺・解は旧配布DLLとバイト一致しました。
検査行列の最大絶対残差は約 `1.51e-14` でした。
カーネル測定はn=600、帯幅64、ウォームアップ1回＋5回の中央値です。
結果は [生データ](native-evidence/simd_20261019/) に保存します。

INCONSISTENT・小モデルCase1（加速off）のSRM全探索では、STRICT 158.2265625秒 →
AUTO 154.2890625秒（約2.5%短縮）でした。17試行、Fs区間2.200～2.2125、中点2.20625、
成功増分7、分解6713、再利用4078、塑性点76が一致しました。
最終変位の最大絶対差は `3.0985e-12`、最大変位で正規化した差は `2.3364e-12`。
最終相対残差は `7.1557e-13` / `6.5596e-13` でした。
DLLの非対称分解呼出数は4960→4961、材料評価呼出数は56561→56583で、内部経路は完全一致しません。
材料点の例外的な分岐からVBAに戻る499要素は両者で同じです。
この条件は材料点・非線形処理の比率が大きく、帯LUのSIMD化による全体短縮は限定的でした。

初期版のINCONSISTENT・固定Fsの小回帰8条件（Case1～4、加速off/adaptive）はすべてPASS。
純VBAとの最大変位差は `2.7756e-17` でした。この記録はストライド更新の選択を調整する前の版です。

## 再構築・再実行

Visual Studio Build Tools（C++ x64、ml64）とPowerShell 7、ExcelのVBAプロジェクトへの
アクセス許可が必要です。出力先には存在しないディレクトリを指定してください。

```powershell
pwsh -NoProfile -File tests/native/build_simd.ps1 -OutputDirectory C:\temp\soilfem_simd -Precision AUTO
python tests/native/test_simd.py --dll native/bin/soilfem_native_x64.dll --baseline workbook/soilfem_native_x64.dll --out tests/results/simd_kernel.json
powershell -NoProfile -ExecutionPolicy Bypass -File tests/native/run_native_compare.ps1 -Book C:\temp\soilfem_simd\2DSoilFEM_simd_20261019.xlsm -Suite smoke -Flow INCONSISTENT -Precision AUTO -Out C:\temp\smoke.json
python tests/native/check_precision_results.py C:\temp\smoke.json
```

STRICT/AUTOをそれぞれ `-Modes AUTO` で実行した結果は、
`python tests/native/check_precision_results.py strict.json auto.json` で比較できます。
`-Suite smokesrm -CaseName Case1 -Variants off` で小モデル1条件の全探索を指定できます。
大きいINCONSISTENTの文献斜面全8条件、全施工履歴・JOINTとの組合せは今回の検証範囲外です。
