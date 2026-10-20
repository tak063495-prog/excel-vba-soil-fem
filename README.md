# VBA版 2D Soil FEM

## 入出力を整理した最新ブック（2026-10-20）

[一括ダウンロード（XLSM＋DLL 2個、ZIP）](workbook/SoilFEM_UI_20261020.zip)。
「操作ガイド」に対応機能・入力順・結果の読み方をまとめ、83設定に入力値と推奨・目安を並べました。
通常使うDLL・I/O設定9項目はAUTO/1を入力済みです。材料・荷重・支持・Fs条件を一括変更しません。
入力813項目、保存・互換性407項目、VBAコンパイル・再読込を確認しています。
Double・収束条件を維持し、SRM全探索のFs区間・全170変位・残差・診断CSV全体が元版と完全一致しました。
[操作・対応機能・検証記録](docs/WORKBOOK_UI_20261020.md) / [現在の83設定と入力済み値](docs/UI_SETTINGS_20261020.md)。

[詳細操作説明書（PDF・47ページ）](docs/manuals/2DSoilFEM_Detailed_Manual_20261020.pdf)：操作手順、実計算の変形・変位・Fs探索・残差など図19点、全83設定、36項目の用語集を収録しています。

## INCONSISTENT 要素行列積・ILU構築の高速化版（2026-10-20）

[一括ダウンロード（XLSM＋DLL 2個、ZIP）](workbook/SoilFEM_INCONSISTENT_ELEMENT_20261020.zip)。
固定3項行列積のループ削減、ILU構築の列検索削減、浮動小数点設定を維持したO3最適化を検証しました。
全探索5組の中央値40.934秒→40.453秒（追加1.18%短縮）。
Double・収束条件を維持し、Fs区間・全170変位・最終残差・GMRES途中値を含むCSV全体が完全一致しました。
VBA/材料DLLコンパイル、モデル17組、正則化回復112照合、材料点20万件も確認。
[内訳・全結果・不採用試作・再現手順](docs/INCONSISTENT_ELEMENT_PERF.md) を参照してください。

## INCONSISTENT ILU・材料評価内部の高速化版（2026-10-20）

[一括ダウンロード（XLSM＋DLL 2個、ZIP）](workbook/SoilFEM_INCONSISTENT_ILUMAT_20261020.zip)。
ILU対角逆数の再利用と、材料評価内の同じ試行応力・降伏判定・三角関数・二乗値の再計算削減を採用しました。
同条件の全探索3組で中央値43.552秒→41.560秒（追加4.57%短縮）。Double・収束条件を維持し、Fs区間・全170変位・最終残差が完全一致しました。
材料点20万件の全出力一致、回帰17組、VBA/DLLコンパイルと旧DLLへの切替を確認しています。
途中のGMRES診断値には差があります。[内訳・全結果・不採用案・再現手順](docs/INCONSISTENT_ILUMAT_PERF.md) を参照してください。

## INCONSISTENT GMRES内部高速化版（2026-10-20）

[一括ダウンロード（XLSM＋DLL 2個、ZIP）](workbook/SoilFEM_INCONSISTENT_GMRES_20261020.zip)。
GMRESの内訳を計測し、ILU前処理とCSR行列積を倍精度SIMD化しました。
小モデルSRM全探索3組で中央値48.452秒→44.359秒（8.45%短縮）。今回は10%に届いていません。
Double・収束条件を維持し、Fs区間・全170変位・最終残差が完全一致。途中のGMRES診断値には差があります。
VBAとDLLのコンパイル、独立GMRES87ケース、回帰17組、旧DLLへの切替を確認しています。
[内訳・全結果・制限・再現手順](docs/INCONSISTENT_GMRES_PERF.md) を参照してください。

## INCONSISTENT 行列組立高速化版（2026-10-20）

[一括ダウンロード（XLSM＋DLL 2個、ZIP）](workbook/SoilFEM_INCONSISTENT_ASSEMBLY_20261020.zip)。
行列組立を分解計測し、最大のCSR散布表作成をDLLへ移しました。
同条件の小モデルSRM全探索3組で中央値53.72秒→45.30秒（15.67%短縮）。
Double・収束条件を維持し、Fs区間・全170変位・最終残差・イベントCSV全体が完全一致しました。
VBAとDLLのコンパイルを確認しています。[内訳・比較結果・再現手順](docs/INCONSISTENT_ASSEMBLY_PERF.md) を参照してください。

## INCONSISTENT ログ整形・バッファ高速化版（2026-10-20）

[一括ダウンロード（XLSM＋DLL 2個、ZIP）](workbook/SoilFEM_INCONSISTENT_EVENT_20261020.zip)。
ログの共通列とGMRESイベントの準備処理をまとめ、バッファを予約領域への追記へ変更しました。
同条件の小モデルSRM全探索3組で中央値53.93秒→51.55秒（4.41%短縮）。
Double・収束条件を維持し、Fs区間・全170変位・最終残差・イベントCSV全体が完全一致しました。
DLL2個は前版と同一です。[内訳・全結果・再現手順](docs/INCONSISTENT_EVENT_PERF.md) を参照してください。

## INCONSISTENT I/O高速化版（2026-10-20）

[一括ダウンロード（XLSM＋DLL 2個、ZIP）](workbook/SoilFEM_INCONSISTENT_IO_20261020.zip)。
前版のDLLをそのまま使い、ログ出力先の再利用とフォルダ確認の集約で、
小モデルSRM全探索の中央値を75.95秒→60.93秒（19.78%短縮）にしました。
Double・収束条件を維持し、Fs区間・全170変位・最終残差・イベントCSV全体が完全一致しました。
ロック解除後のログ再試行、保存先変更など24項目も確認しています。
[処理時間の内訳・比較結果・ログ障害時の制限と再現手順](docs/INCONSISTENT_IO_PERF.md) を参照してください。

## INCONSISTENT高速化版（2026-10-20）

[一括ダウンロード（XLSM＋DLL 2個、ZIP）](workbook/SoilFEM_INCONSISTENT_20261020.zip)。
[実行ブック](workbook/inco_native_20261020/2DSoilFEM_inco_native_20261020.xlsm) と
[計算DLL](workbook/inco_native_20261020/soilfem_native_x64.dll)、
[材料DLL](workbook/inco_native_20261020/soilfem_mat_x64.dll) を同じフォルダに置いてください。
処理時間を高精度タイマーで分解し、最大のCSR/GMRES経路と材料状態の復元を高速化しました。
INCONSISTENTの小モデルSRM全探索は3回の中央値で155.81秒→70.55秒（2.21倍）、Fs区間・全170変位・最終残差が一致しました。
Double・収束条件・材料則・SRM探索条件を維持し、AUTOでCPUに応じてAVX2/FMAまたはSSE2を選びます。
従来演算はSTRICTで選択できます。[速度・Fs区間・全変位・残差の比較と再現手順](docs/INCONSISTENT_NATIVE_PERF.md) を参照してください。

## 倍精度SIMD試験版（2026-10-19）

[SIMD版ブック](workbook/simd_20261019/2DSoilFEM_simd_20261019.xlsm) と同じフォルダに
[帯行列DLL](workbook/simd_20261019/soilfem_native_x64.dll) と
[材料DLL](workbook/simd_20261019/soilfem_mat_x64.dll) を置いてください。
「高速化設定」の演算方式でSTRICT / SSE2 / AUTOを選べます。AUTOは倍精度を維持したまま
対応CPUでAVX2・FMAを使います。従来DLLからの変更、精度比較、測定範囲は
[SIMD DLLの説明](docs/SIMD_NATIVE.md) を参照してください。

## 旧標準版：候補48＋DLL（2026-10-18）

[標準ブックをダウンロード](workbook/2DSoilFEM_standard48_native_20261018.xlsm)。同じフォルダに [soilfem_native_x64.dll](workbook/soilfem_native_x64.dll) と [soilfem_mat_x64.dll](workbook/soilfem_mat_x64.dll) を置いてください。ビルドは `20261018_STANDARD48_NATIVE_05`、保存済み設定は `FLOW_POLICY=INCONSISTENT`、`NATIVE_DLL=AUTO` です。DLLを利用できない環境ではVBAへ自動で戻ります。

候補48の正則化作業配列の格納改善を、最新DLL版の孤立DOF・ユーザー中断修正と統合しました。DLLのAPI・バイナリ、材料則、演算順、収束判定とSRM探索条件は変更していません。DLLの正則化経路を優先し、VBAへ戻る経路にも48の改善を適用します。[変更内容・回帰結果・測定範囲](docs/STANDARD48.md) を参照してください。以下の旧ブックと測定結果は履歴として保持しています。

Excel上で2次元地盤FEM、強度低減法（SRM）、ジョイント、施工ステージ、メッシュ生成を実行するブックです。2026-10-08版では、操作シートを用途別に整理し、小モデルで見つかった結果処理・変位載荷の不具合を修正しています。

2026-10-10のINCONSISTENT修正版では、悪条件の非対称帯LUに反復改良を追加し、SRMの収束を材料応力から求める力で判定します。失敗した増分は確定状態へ戻して再試行し、数値求解の失敗からFsの上限を作らない判定を維持します。[改修・Fs検証の履歴](docs/INCONSISTENT_SRM_IMPROVEMENT.md) に診断、試行履歴、確認範囲を記録しています。旧配布ブックのFLOW_POLICYはDAVIS、新標準ブックはINCONSISTENTです。

限界安全率はFs区間、材料力の釣合い、降伏条件、メッシュ依存、変形を併せて評価します。今回の対象は小変形モデルのINCONSISTENT・ψ=0です。大変形の実装はFsの安定性確認後に進めます。過去のINCO_14～16とDAVISの比較は [旧版の限界安全率検証](docs/PRACTICAL_SRM_VALIDATION.md) に保持しています。特定のベンチマークの合格だけで、一般の斜面の最終設計をこのブック単独で判断できると認定したものではありません。

最新検証では6物理条件・8メッシュすべてでFs区間が成立し、保存下端の材料力の釣合いと3主応力による降伏条件を独立に確認しました。区間中点と文献代表値との差は−3.13～＋1.14%、主比較の粗密メッシュ差はφ=20°で0.92%、φ=0°で0.85%です。細メッシュの全探索は並列実行時の実測で約10.12時間を要しました。大きな局所ひずみが出た条件もあり、Fsの比較結果と変形の物理評価は分けています。[全結果・生ログ・再実行条件](docs/INCONSISTENT_SRM_IMPROVEMENT.md) を参照してください。

## ネイティブDLL版（2026-10-17、20261017_NATIVE_02）

**DLL高速化 第2弾（Phase B/C、レビュー中）**：[実行用ブック](workbook/2DSoilFEM_nativekern_20261017_0846.xlsm) と同じフォルダに [soilfem_native_x64.dll](workbook/soilfem_native_x64.dll) と [soilfem_mat_x64.dll](workbook/soilfem_mat_x64.dll) を置くと、材料点更新・要素接線・hourglass・全体組立・残差集計・ILU(0)/GMRES・正則化補正・RCMもDLLで計算します（`NATIVE_DLL`=OFF/BAND/MAT/AUTO）。純VBAとビット一致のまま、同梱モデルのSRM全探索は 574.4 s → 36.4 s（15.8倍）、INCONSISTENTの小規模SRMは 1753.8 s → 140.8 s（12.5倍）。詳細は [ネイティブDLLによる高速化](docs/NATIVE_DLL.md) の9章。

[実行用ブック（DLL対応）](workbook/2DSoilFEM_native_20261017_0505.xlsm) と同じフォルダに [soilfem_native_x64.dll](workbook/soilfem_native_x64.dll) を置くと、帯LU分解と求解をDLLで計算します。行列全体を1回の呼出しで渡すため、呼出しのオーバーヘッドはほとんどありません。結果は純VBAとビット単位で一致します。同梱モデルのSRM全探索は 574.8 s → 189.3 s（3.04倍）、分解だけでは33.8倍速くなりました。DLLがない場合や読込みに失敗した場合は、自動で従来のVBAで計算します。INCONSISTENTのSRMで破壊付近の試行が極端に遅くなる問題も修正しました（弾性補正反復の収束速度判定）。設計・検証・制限は [ネイティブDLLによる高速化](docs/NATIVE_DLL.md) を参照してください。

## 説明書

初めて使う方は **[操作説明書](docs/USER_GUIDE.md)** から始めてください。ダウンロード・マクロ起動、シートの役割、メッシュ生成、載荷・施工・接合の例、SRM、結果保存まで記載しています。

|文書|内容|
|---|---|
|[操作説明書](docs/USER_GUIDE.md)|起動から解析までの手順、形状入力、BIRTH/DEATH、JOINT、SRMの操作例|
|[入力シート仕様](docs/MODEL_INPUTS.md)|材料・節点・要素・接合・載荷・ステージの各列、節点順序、MATSETの入力|
|[設定一覧とKEY対応](docs/SETTINGS_REFERENCE.md)|全74設定の入力セル・KEY・配布保存値・選択肢、旧C37〜C46との対応|
|[結果・ログ・再開と対処法](docs/RESULTS_AND_TROUBLESHOOTING.md)|結果の列と図の1〜68番号、FOS区間、限界状態、CSV、再開、問題の確認方法|
|[プログラムの概要と保守](docs/TECHNICAL_OVERVIEW.md)|計算の流れ、VBAモジュールの担当、高速化候補、開発・テスト方法|
|[ネイティブDLL](docs/NATIVE_DLL.md)|DLLのAPI設計、自動フォールバック、VBAとの一致確認、計測時間|
|[高速化履歴](docs/ACCELERATION_HISTORY.md)|V1〜V5、組合せ、ADAPT_01〜07の測定結果と採用・見送り理由|
|[動作確認報告](docs/VALIDATION.md)|再現した不具合、修正、確認結果、未確認の範囲|
|[文献・理論解による検証](docs/LITERATURE_VALIDATION.md)|5載荷条件、Griffiths・PruškaのSRM、メッシュ・流れ則の比較、再実行用XLSMと生ログ|
|[限界安全率と実用性の検証](docs/PRACTICAL_SRM_VALIDATION.md)|複数斜面のFs探索、独立Bishop、メッシュと安定化感度、実務で使える範囲と残る課題|
|[INCONSISTENTのSRM改良](docs/INCONSISTENT_SRM_IMPROVEMENT.md)|失敗行列の診断、材料力の釣合い、確定状態からの再試行、最新のFs比較と生ログ|

操作説明は配布版 `20261010_INCO_24` のブックとVBAを照合しています。入力値の例は操作確認用です。材料・荷重・メッシュの妥当性は対象モデルに合わせて設定してください。

## Issue #1への対応（2026-10-09）

本版は間隙水圧を考慮しない全応力解析です。圧密の指摘2・4・5は休眠コードについての問題で、配布版の計算経路には該当しません。CSR補助混合UPも本番では起動せず、追加A・Bは将来その経路を有効化する場合の課題です。

自重→外力を同一呼び出しで処理する互換経路の荷重保持を修正しました。JOINTの未使用ψ入力は0に限定し、履歴なし応力上限モデルの制限をシートと説明書に明記しています。圧密の隠し設定を有効にすると解析を拒否します。各項目の判定とExcelでの再現結果は [Issue #1レビュー記録](docs/ISSUE_1_REVIEW.md) にあります。

## 使用方法

1. [最新ZIP](workbook/SoilFEM_UI_20261020.zip)を展開し、XLSMとDLL 2個を同じローカルフォルダに保存し、デスクトップ版Excelで開きます。信頼するファイルとしてマクロを有効にしてください。
2. 「操作パネル」から、形状・メッシュ、材料、載荷、接合、施工ステージを設定します。「解析設定」「高速化設定」「メッシュ設定」「出力・表示」「詳細設定」で用途別に設定できます。
3. 初期図・全体図で形状と要素の有効状態を確認し、解析を実行します。結果図、結果節点、結果要素、ステージ結果、接合結果を確認します。

ブックには既存のモデルと保存結果が含まれています。今回の選別方式で再解析してから結果図を使用してください。元の「設定」シートのKEYと新しい入力シートは連動します。

## 今回修正した動作

- JOINTを連続体の面積・応力回復に含めたことによる解析終了時のエラーを修正しました。接合の3積分点の法線応力・せん断応力・開口・すべり・接触状態は「接合結果」へ保存します。
- DEATHで無効になった要素を節点応力の平均から除外しました。
- JOINTの結果保存で開口・すべりを連続体ひずみで上書きしないよう修正しました。結果図には6節点の接合を紫線で表示します。
- 自重・外力がゼロの変位載荷／除荷で、丸め誤差により非収束になる問題を修正しました。変位載荷点の反力を残差の基準値に使います。除荷時は前ステージの確定反力を基準に保持します。収束許容値は変更していません。
- RESET_STRESSで基準座標を変形後へ更新した場合に、保存結果の座標を使って結果図を表示します。入力変更後の再解析チェックは保持しています。

## 高速化の版と履歴

最新の計算部は `20261020_INCO_ELEMENT_01`、操作シートは `20261020_UI_01`、実務修正は `20261009_PRACTICAL_04` です。V1〜V5の単独・組合せとADAPT_01〜07の結果、見送り理由、出典は [高速化履歴](docs/ACCELERATION_HISTORY.md) と [比較CSV](docs/benchmark-summary.csv) にまとめています。

ADAPT_07は3770.916秒で、実測最速のADAPT_03（2929.445秒）より28.72%遅くなりました。今回の版ではADAPT_03の選別ロジックと採算推定を復元し、V1/V2a/V4・適応選別をON、増分幅の追加回復とfresh LU試験をOFFにしています。実務上の修正と診断ログは保持しています。[今回の評価と復元内容](docs/evidence/adaptive_policy7_evaluation_and_restore_20261008.txt) に内訳を記載しています。

過去のADAPT_03は単発測定です。今回の実務修正・診断を含むブックのフル速度は再測定していません。新しいログのビルド番号と `AdaptivePolicy=ADAPT_03_RESTORED`、`StepRecoveryEnabled=False` で使用した方式を確認できます。

## 確認内容

文献検証では、弾性の5載荷条件は全節点変位・積分点応力・支持反力が理論式と一致しました。SRMはGriffithsとPruškaの乾燥斜面を比較しています。旧R1のINCONSISTENT・ψ=0には早い非収束がありました。修正版の確認範囲はINCONSISTENT修正報告で分けて記録しています。DAVISは降伏強度と流れ則を変えます。表示Fsが近いことだけで再現一致とは扱いません。[条件・比較・限界と検証ケース](docs/LITERATURE_VALIDATION.md) を参照してください。

[動作確認報告](docs/VALIDATION.md) に、24条件の回帰解析、施工・載荷テスト、接合結果の117確認、設定・UI確認の結果を記載しています。フルモデルの所要時間測定は含みません。

Windows、デスクトップ版Excel、PowerShellで次を実行すると、同梱ブックのコピーに対してテストできます。テストではExcelの「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」が必要です。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/run.ps1
```

テストは専用のExcelインスタンスを作り、`tests/tmp` のコピーに一時的なVBAテストコードを追加します。結果は `tests/results` に保存します。配布ブックや、別に開いている解析ブックは変更しません。

## VBAソース

`src/vba` にブックから取り出したソースと `modules.json` を収録しています。日本語を保持するためUTF-8で保存しています。`.bas` / `.cls` はCodeModuleの本文で、VBEのエクスポート形式のAttribute行やUserForm等を含みません。VBEの「ファイルのインポート」でブック全体を再構成する形式ではありません。実行用には同梱のXLSMを使用してください。

計算結果の採用には、メッシュ・境界条件・材料の妥当性と、変位・応力・塑性状態の確認が必要です。数値上の収束と力学的な限界状態は別の診断です。物理圧密は未対応です。CSR補助混合UPやすべての施工履歴の組合せを今回の確認だけで検証したものではありません。

## ライセンス

このリポジトリのライセンスは [Apache License 2.0](LICENSE) です。
