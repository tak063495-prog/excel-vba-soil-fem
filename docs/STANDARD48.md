# 標準48とDLL版の統合（2026-10-18）

[標準ブック](../workbook/2DSoilFEM_standard48_native_20261018.xlsm) · [README](../README.md) · [DLLの仕様と旧検証](NATIVE_DLL.md)

## 変更と配布設定

利用者の指定により候補48を標準にしました。GitHubの `093ff58`（孤立DOF・Err18修正版）を基に、正則化残差回復のVBA作業配列を行方向に連続した1次元配列へ変更しています。ビルド番号は `20261018_STANDARD48_NATIVE_05` です。

列スケーリング、正規方程式、Choleskyの演算順、独立残差検査、収束条件、SRM探索、材料則を変更していません。DLL版に既存の反復制御・孤立DOF修正・中断処理も保持しています。`FEMSolver` の正則化関数群と `FEMCore` のビルド番号以外の計算モジュールは変更していません。

`AUTO` は既存のDLL正則化カーネルを引き続き優先します。`OFF` / `BAND` / `MAT`、またはDLLを利用できない場合のVBA経路で48の配列改善が働きます。DLLのABI・Cソース・バイナリを変更していないため、既存DLLをそのまま使えます。

Long添字の容量を超える寸法や配列割当のError 7では、従来の2次元VBA処理へ戻ります。部分割当を解放し、分解回数の二重計上を避けます。ユーザー中断のError 18は回復処理へ変換せず、既存の中断処理へ伝えます。

保存済みの設定は `FLOW_POLICY=INCONSISTENT`、`NATIVE_DLL=AUTO` です。26モジュールを更新し、VBAコンパイル、保存・再読込後のソース照合（VBEの識別子表記・行末空白を正規化）、KEY値の保存を確認しています。旧配布ブックは上書きせず履歴として残しています。

## 確認した測定範囲

統合前の候補48単独について、Excel `16.0.20430.20182` のH7・φ20 LIMIT 1ケースを実測しました。全体時間は **4899.8326121秒（81分40秒）**。seed、SRM完全確認、最終抽出、保存、独立力・全3D Mohr–Coulomb検査を含みます。入力準備とExcel起動は除きます。

Fs区間は1.6625～1.675、中点1.66875、幅0.0125。独立材料力の相対残差は `3.770788769155266e-11`、全3D MCの正規化最大降伏余裕は `4.738008280352822e-15`（960積分点）でした。実行前後のExcel版・バイナリ・OS起動・電源設定は一致しています。

32の時間は異なるExcel版で測られているため、正式な速度比として扱いません。この81分40秒は統合後DLL版の全体時間でもありません。今回の標準化は利用者の採用判断によるもので、全8条件・逆順の速度採用ゲートが合格したという意味ではありません。小変形の数値検査は有限変形の妥当性を保証しません。

## 統合後の小回帰

- 正則化28条件（非対称・特異・スケール差・拘束・帯幅・不正damping・メモリ拒否）を、OFF / BAND / MAT / AUTOで従来2次元処理と比較。
- DLLあり・なしの各112比較、合計224比較で、成否・全解ベクトルのDoubleビット・独立期待解・主診断・分解回数を確認。DLLありのAUTOでは正則化DLLの実呼出数も確認。
- Long容量の8境界、配列割当Error 7からの復帰、中断Error 18の伝播をDLLあり・なしの両方で確認。
- 既存の孤立DOF・中断回帰と、小モデルのVBA/DLL比較も実行。実行記録は `docs/native-evidence/standard48_*` に収録。

再実行例（出力先には毎回新しい名前を指定、PowerShell 7を使用）：

```powershell
pwsh -NoProfile -File tests/native/check_packed48.ps1 -SourceWorkbook workbook/2DSoilFEM_standard48_native_20261018.xlsm -OutputRoot tests/results/standard48_new
pwsh -NoProfile -File tests/native/check_packed48.ps1 -SourceWorkbook workbook/2DSoilFEM_standard48_native_20261018.xlsm -OutputRoot tests/results/standard48_new_no_dll -WithoutDll
```

テストはブックのコピーだけを使います。別のExcelが起動していれば開始せず、終了時もテスト自身のExcelだけを閉じます。PIDの強制終了は行いません。
