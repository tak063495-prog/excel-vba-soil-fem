# ネイティブDLLによる高速化（20261016_NATIVE_01 / 20261017_NATIVE_02）

[READMEへ](../README.md) · [高速化履歴](ACCELERATION_HISTORY.md) · [プログラムの概要](TECHNICAL_OVERVIEW.md)

## 1. 目的とボトルネック

既存ログ（ADAPT_07本番、3770.916 s）の内訳は、帯LU分解 2560 s（68%）、材料点評価 727 s（19%）、線形求解 188 s（5%）、接線 202 s（5%）、組立 21 s でした。時間の大部分は帯LU分解とその前進・後退代入で、どちらも O(n·bw²)/O(n·bw) の単純な二重・三重ループです。VBAの1演算あたりの費用が大きいので、ここをDLLへ移すのが最も効果的です。

## 2. DLL APIの設計（粗粒度）

VBA↔DLL呼出しのオーバーヘッドを避けるため、**行列全体を1回の呼出しで処理**します。要素・行・演算単位の呼出しはしません。配列はすべてVBAが所有し、DLLはメモリ確保もグローバル状態も持ちません（VarPtrで先頭アドレスを渡す）。

| エクスポート | 1回の呼出しで行う処理 | 置き換えるVBA |
|---|---|---|
| `sfn_abi` | ABI確認（0x00010001） | — |
| `sfn_band_factor` | 2次元帯配列の1次元パック＋帯LU分解全体＋ピボット検査 | `P6PackBand`＋`P6BandFactor1D` |
| `sfn_band_solve` | 前進消去＋後退代入（右辺全体） | `P6BandSolve1D` |
| `sfn_nsband_factor` | 非対称帯LU（部分ピボット、行交換、乗数保存）全体 | `P6SolveNonsymmetricBand` の消去ループ |
| `sfn_nsband_forward` | 保存した交換・乗数の前進再生 | `P6BandReplayForward`、反復改良の前進 |
| `sfn_nsband_back` | U部分の後退代入（非有限・小ピボット検査付き） | 同関数の後退代入、反復改良の後退 |

- 対称帯（DAVIS / BAND_LU NEW）と非対称帯（INCONSISTENT）の両経路が対象です。OLD 2Dカーネル、LDLT影比較、正則化回復（`P6SolveResidualRecovery`）はVBAのままです。
- **結果はVBAとビット単位で一致**します。VBA(x64)はDouble式をx87スタックで評価し代入時だけ丸めるため、更新式はすべて `sfn_fp.asm` のx87カーネル（呼出し側の制御語をそのまま使用）で計算し、ループ順もVBAと同じにしています。CのSSE2で書いた初版は丸めが1e-16程度ずれたため採用していません。
- 呼出しは分解1回につき1回、求解1回につき1回です。ピボット統計（`P6RecordPivot`）や失敗時の診断・`SetAnalysisFailure` はVBA側に残し、失敗時の動作・メッセージは従来と同じです。

## 3. 自動フォールバック

`FEMNative.bas` がDLLの読込みと呼出しを管理します。

1. ブックと同じフォルダの `soilfem_native_x64.dll` だけを `LoadLibraryExW`（DLL_LOAD_DIR|DEFAULT_DIRS）で読込みます。PATH検索はしません。別パスの同名DLLが読込み済みなら使いません。
2. 64bit Excel以外、DLLなし、エクスポート欠落、ABI不一致、配列の範囲不正、DLLの引数エラー、VBAの実行時エラーのいずれでも、その呼出しは**入力を変更しないまま従来のVBAカーネルで計算**し、以後その解析ではDLLを使いません。
3. 設定キー `NATIVE_DLL`（設定シートに無ければ `AUTO`）。`OFF` / `0` / `VBA` で純VBAに固定できます。試験用に `FEMNativeSetMode "OFF"|"AUTO"` もあります。
4. ログ：`perf_summary.txt` に `NativeDLL=NATIVE(...;factor=;solve=;nonsym=;fallback=)` または `VBA(理由)`、`run.log` に `SOLVER,NATIVE` 行を出力します。版は `FEM_BUILD_STAMP=20261016_NATIVE_01`。

## 4. 検証（Triela、Excel 64bit、1本ずつ実行）

同じブック・同じ入力で `NATIVE_DLL` だけを切り替え、`TDisp`（全自由度変位）と状態・収束値・増分数・分解数・再利用数・塑性点数を比較しました。

| ケース | 経路 | 判定 | 変位の一致 | 状態・カウント |
|---|---|---|---|---|
| smoke Case1〜4 × 加速OFF/適応ON（8件） | DAVIS（対称帯） | 全件PASS | 170/170成分 ビット一致、最大差0 | 全項目一致 |
| smoke Case1〜4 × 加速OFF/適応ON（8件） | INCONSISTENT（非対称帯） | 全件PASS | 170/170 ビット一致 | 全項目一致 |
| 同梱モデル 1976節点・701要素 固定Fs=2.00 | 対称帯 | PASS | 2010/2010 ビット一致 | 一致 |
| 同梱モデル 固定Fs=2.25 | 対称帯 | 両方 NONCONVERGED（純VBAでも同じ） | 2010/2010 ビット一致 | 一致 |
| 同梱モデル SRM全探索 | 対称帯 | PASS | 2010/2010 ビット一致 | 増分83・分解1019・再利用893・塑性697 一致 |
| DLLなしフォルダで smoke 8件 | DAVIS | 全件PASS、`VBA(DLL_NOT_FOUND)` | — | 純VBAと同じ収束値 |

固定Fs=2.25のNONCONVERGEDはDLLの有無に関係なく同じ増分で同じ値になり、Fs区間の失敗側として扱われる結果です。DLL導入で途中停止が増えた例はありません。

## 5. 計測時間

| ケース | 純VBA | DLL | 全体 | 分解 VBA→DLL | 求解 VBA→DLL |
|---|---:|---:|---:|---:|---:|
| 同梱モデル SRM全探索 | 574.8 s | 189.3 s | **3.04倍** | 378.3 s → 11.2 s（33.8倍） | 23.2 s → 2.2 s（10.4倍） |
| 同梱モデル 固定Fs=2.25 | 122.6 s | 44.6 s | 2.75倍 | 77.7 s → 2.3 s（33.4倍） | 4.55 s → 0.43 s |
| 同梱モデル 固定Fs=2.00 | 14.3 s | 6.0 s | 2.40倍 | 8.8 s → 0.30 s（29倍） | 0.52 s → 0.05 s |
| smoke（85節点） | 1.6〜2.3 s | 1.6〜2.1 s | ほぼ同じ | 小さすぎて差なし | — |

各1回の測定です。Timer分解能は約16 msです。DLL使用後に残る時間は、材料点評価（SRM全探索で約120 s）と接線計算（約34 s）が大部分です。

## 6. 制限と次の候補

- 材料点更新（Mohr–Coulomb戻り写像、全ガウス点）と要素接線・組立はVBAのままです。全体時間の残りの約8割がここなので、さらに速くするには要素・材料状態の配列をまとめて渡す `sfn_eval_all` 型のAPIが次の段階になります。分岐とログが多いため、ビット一致の検証と合わせて別版で行う予定です。
- 32bit Excelでは常に純VBAで動きます。
- DLLは Visual Studio 18 BuildTools（`native/build.cmd`、ml64＋cl /O2 /fp:strict /MT）でビルドしました。再ビルド後は `tests/native` の比較を再実行してください。

## 7. 再現方法

```powershell
# 作業コピーに変更モジュールを反映（元ブックは上書きしない）
powershell -File tests/native/make_native_workbook.ps1 -Src <元.xlsm> -Dst <name_native_yyyyMMdd_HHmm.xlsm> -ModDir src/vba
# DLLを同じフォルダに置き、VBA/DLLを交互に実行
powershell -File tests/native/run_native_compare.ps1 -Book <ブック> -Suite smoke -Flow DAVIS -Out smoke.json
powershell -File tests/native/run_native_compare.ps1 -Book <ブック> -Suite builtin -FixedFs 0 -Modes AUTO,OFF -Out srm.json
python tests/native/compare_results.py srm.json
```

証拠の生データは [docs/native-evidence](native-evidence) にあります。

## 8. 2026-10-17 追加検証と修正（20261017_NATIVE_02）

### 8.1 小規模モデルのSRM全探索（DLL ON/OFF）

| 流れ則 | ケース | 結果 |
|---|---|---|
| DAVIS | Case1〜4 × 加速 off/adaptive（8組） | 全組で変位170/170がビット一致、FOS・反復数・残差一致 |
| INCONSISTENT | Case1・加速off | 変位170/170がビット一致、FOS=2.206（PASS 2.200 / FAIL 2.213）、反復数一致。OFF 1773.0 s / DLL 1724.3 s |

### 8.2 INCONSISTENTのSRMが極端に遅くなる問題の修正

**症状**：INCONSISTENTで、破壊に近いFs（2.300、2.250、2.225）の試行1回に約14分かかり、全探索がいつまでも終わらないように見えました。テスト用スクリプトの問題ではなく、DLLを使わない純VBAでも同じでした。

**原因**：収束しない増分のたびに、Newton反復100回のあと弾性補正反復を上限の1000回まで回していました。破壊荷重付近では、弾性補正の残差は増分幅に比例したところで止まり、それ以上下がりません。そのため1000回回っても毎回カットバックになっていました。受理された弾性補正は最大でも93回で収束しており、残り約900回は無駄な計算でした。また既存の停滞判定は弾性補正モードを対象外にしていました。

**修正**（`FEMEngine.bas` `P3RunLoadStages`）：弾性補正モード（INCONSISTENTのみ）で、50反復ごとに残差の幾何収束率を測ります。工学的許容残差（1E-5）に届くまでの予測反復数が残りの上限回数を超えたら、その時点でカットバックします（イベント `ELASTIC_SLOW_CONTRACTION`）。判定は反復100回目以降に限るため、100回以内で収束する試行には影響しません。DAVISは判定に入らないため、DAVISの結果は変わりません。修正後のブックでDAVIS 8組を再実行し、修正前と変位・FOSが完全に一致することを確認しました（`smokesrm_davis_fixed_20261017.json`）。

**効果**（Case1・加速off・純VBA）：

| Fs試行 | 修正前 | 修正後 | 判定・残差 |
|---|---:|---:|---|
| 2.300 | 832 s | 511 s | 同じFAIL（1.017E-05） |
| 2.250 | 672 s | 211 s | 同じFAIL（1.019E-05） |
| 2.225 | 820 s | 453 s | 同じFAIL（1.013E-05） |
| 2.213までの累計 | 2328 s | 1179 s | −49% |
| 全探索 | 2898 s以上経っても終わらず中止 | 1773 s で完了 | FOS=2.206 |

残りの時間は、主に破壊荷重付近で収束しないNewton反復（1増分あたり100回、それぞれに直線探索）です。この部分の短縮は今後の課題です。

## 9. Phase B/C: 材料点・接線・組立・GMRES/ILU・RCM などのDLL化（20261017_NATIVE_03/04）

### 9.1 モード（設定キー `NATIVE_DLL`）

| 値 | 内容 |
|---|---|
| `OFF` | 純VBA（従来どおり） |
| `BAND` | 帯LU分解・求解・非対称帯LUのみDLL（PR #4 の版） |
| `MAT` | BAND ＋ 材料点更新（Mohr–Coulomb応力戻し、内力、hourglass内力）と要素接線（Phase B） |
| `AUTO`（既定） | MAT ＋ 以下のPhase Cカーネル |

どのモードでも、DLLが見つからない／APIが合わない／DLLが要素や行列を受け付けないときは、その部分だけ従来のVBAコードで計算します（ログ・`SetAnalysisFailure` の挙動は従来どおり）。

### 9.2 Phase Cで `soilfem_mat_x64.dll`（API 4）へ移した計算

| 関数 | 対象のVBA処理 | 備考 |
|---|---|---|
| `sfm_band_scatter` | `SetTotalMat` の要素剛性→帯行列の足し込み | 散布表の順で加算（加算順も同一） |
| `sfm_csr_scatter` | `P6AssembleCSRValuesDetailed`（非対称用CSR組立） | |
| `sfm_csr_resnorm` | `P6CSRMatVecResidualNorm`（非対称帯LUの残差・反復改良） | u-p混合時はVBA |
| `sfm_ilu0` | `P6BuildCSRILU`（ILU(0)） | |
| `sfm_gmres` | `P6SolveCSRGMRES`（右前処理GMRES(m)＋ILU0、2回MGS、Givens） | 再始動ごとの `GMRES_RESTART` 等のイベントはDLLが記録し、VBAが同じ順・同じ書式で出力 |
| `sfm_regularized` | `P6SolveResidualRecovery`（列スケーリングLM正規方程式・帯Cholesky・検証） | `REGULARIZED_SOLVE` イベントはVBA |
| `sfm_residual` | `P3AccumulateResidualMetrics`（残差ノルム・反力・エネルギー誤差の集計） | u-p混合時はVBA |
| `sfm_rcm` | `P6BuildRCMOrdering` ＋ `P6ComputeRCMCandidateBandwidth` | 同じ順序規則（次数→節点番号）で同一の置換を生成。`P6RcmSelfCheck` で全ケース一致を確認 |
| `sfm_hg_apply` | `P6AddQ8HourglassStabilization` のキャッシュ済み経路 ＋ `P3KmatIsUnsymmetric` | モード再計算が必要な要素はVBA |

VBA側のI/O改善（AUTOのみ）：`solver_events.csv` を1行ごとに開閉せず、メモリに貯めて512行ごと・他のCSV書込み前・解析終了時にまとめて書きます。出力ファイルはバイト単位で同一でした（INCONSISTENT Case1で SHA-256 一致、13.7 MB）。INCONSISTENTではイベントが約8万行出るため、これが残りの最大の待ち時間でした。

全カーネルはPhase Bと同じx87方式（`-mfpmath=387`、VBAと同じ丸め位置、`Sqr`/`^` はVBAと同じMSVCR100の関数）で、純VBAとビット単位で一致します。

### 9.3 Newton反復を1回の呼出しにしなかった理由

1反復は「材料点更新＋内力（1呼出し）」「残差集計（1呼出し）」「接線＋hourglass＋組立（2〜3呼出し）」「分解・求解（1〜2呼出し）」の粗粒度呼出しになり、要素・行単位の呼出しはありません。反復1回分を完全に1呼出しにするには、接線再構築の判定・線探索・加速（Anderson等）・停滞判定・弾性補正ガードなど約3000行の分岐をDLLへ移す必要があり、ログと失敗時処理を従来どおり保つ要件と両立しないため見送りました。計測では呼出し回数由来のオーバーヘッドは誤差程度で、残りは分解（DLL内のx87演算）とVBAの制御ロジックです。

### 9.4 RCM
RCMは既存の `RCM_POLICY`（AUTO/ON/OFF）がそのまま切替えスイッチです。DLL版は順序をVBA版と同一に生成するため、RCMのON/OFFで結果が変わることはあっても、DLLとVBAの間で結果は変わりません（全ケースで置換・帯幅一致、変位もビット一致）。

### 9.5 検証結果（Triela、Excel 64bit、1本ずつ実行）

正しさ（OFFとの比較。変位はビット比較、増分数・分解回数・再利用回数・塑性点数・FOS文字列も比較）：

| ケース | 比較 | 結果 |
|---|---|---|
| 小規模DAVIS SRM 8ケース | OFF / BAND / MAT / AUTO | 全ケース 170/170 ビット一致、FOS一致 |
| 小規模INCONSISTENT Case1（加速なし）SRM | OFF（Phase B時の結果） / AUTO | 170/170 ビット一致、カウント一致、FOS 2.206 一致 |
| 同梱モデル 固定Fs 2.00 | OFF / BAND / MAT / AUTO | 2010/2010 ビット一致 |
| 同梱モデル 固定Fs 2.25 | 同上 | 2010/2010 ビット一致（全モードで同じ増分で未収束＝破壊側判定） |
| 同梱モデル SRM全探索 | 同上 | 2010/2010 ビット一致、FOS 2.244 一致 |

計算時間（秒）：

| ケース | OFF | BAND | MAT | AUTO | OFF/AUTO |
|---|---:|---:|---:|---:|---:|
| 同梱 固定Fs 2.00 | 14.8 | 5.3 | 2.9 | 2.7 | 5.5倍 |
| 同梱 固定Fs 2.25 | 122.7 | 42.5 | 11.7 | 9.5 | 12.9倍 |
| 同梱 SRM全探索 | 574.4 | 184.6 | 49.2 | 36.4 | 15.8倍 |
| 小規模INCONSISTENT Case1 SRM | 1753.8 | — | — | 140.8 | 12.5倍 |
| 小規模DAVIS Case1 加速なし | 7.07 | 5.34 | 2.83 | 2.42 | 2.9倍 |

INCONSISTENTの内訳：GMRES/ILUのDLL化で 1753.8 → 691.8 s、イベント出力のまとめ書きで → 142.7 s、hourglassのDLL化で → 140.8 s。
同梱SRM（AUTO）の残り内訳：分解 11.2 s、材料点 4.9 s、接線＋組立 6.5 s、求解 2.1 s、その他（増分制御・ログ）約12 s。

証跡：`docs/native-evidence/phaseC_*.json`。
