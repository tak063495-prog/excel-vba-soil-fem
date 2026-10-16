# ネイティブDLLによる高速化（20261016_NATIVE_01）

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
