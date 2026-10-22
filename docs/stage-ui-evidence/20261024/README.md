# 1024版の合格記録

[集計](stage_ui24_check15.json)／[全ログ・数値・画面操作記録](verified_records.zip)／[計算済みブック](calculated_workbooks.zip)／[配布物ハッシュ](artifact_hashes.json)

`verified_records.zip` は最終配布ブックに対する合格記録です。試行途中の失敗ブックや旧候補は含めません。

| 保存先 | 内容 |
|---|---|
| package15.json / build15.log | VBEコンパイル、35モジュールの厳密なソース一致、保存・再読込、既存入力表の保持 |
| ui15/checks.json | 実イベントを使った入力・工程編集・設定切替・コピー・保存再読込の102確認 |
| regression15 | 基礎水圧・浸透・浮力・施工・JOINT・SRM・異常入力、出力表、残差・反力、実行ログ |
| front15 | 新画面から物性・水位を編集して計算、保存状態から再開 |
| disabled15 / disabled_ref15 | 水理工程を無効化して設定を保持する場合と、その計画を除いた基準計算 |
| screen_dist15 | 7画面・メッシュ生成の実Excel操作記録（FEMStageUi／FEMStageEditorは配布版と同一） |
| screen_result15 | 最終配布ブックから計算した水位1m斜面SRMの専用コピーによる図・水圧・保存再表示・入力変更検知 |
| reference | 比較に用いた従来版の保存数値 |

JSONの元記録には作者PCの実行パスが残っています。入力や検証資料へアクセスするためのリンクではありません。検証用の公開入力はリポジトリの `tests/fixtures`、ドライバーは `tests/run_water_display_regression.ps1`、`tests/run_groundwater_fixture.ps1`、`tests/run_stage_ui24.ps1`、`tests/inspect_stage_ui24.ps1` にあります。

## 再実行

Excelをすべて閉じた環境で、配布ZIPを通常フォルダーへ展開します。テストは専用コピーだけを編集し、他のExcelが起動中なら開始しません。既存の出力先も上書きしません。

```powershell
pwsh -NoProfile -File tests/run_stage_ui24.ps1 -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -Out tests/tmp/stage_recheck/ui15
pwsh -NoProfile -File tests/run_water_display_regression.ps1 -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -Out tests/tmp/stage_recheck/regression15
pwsh -NoProfile -File tests/run_groundwater_fixture.ps1 -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -Fixture tests/fixtures/water_stage_one_way.json -Out tests/tmp/stage_recheck/front15 -FrontUi -Resume
pwsh -NoProfile -File tests/run_groundwater_fixture.ps1 -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -Fixture tests/fixtures/water_stage_disabled_plan.json -Out tests/tmp/stage_recheck/disabled15 -Resume
pwsh -NoProfile -File tests/run_groundwater_fixture.ps1 -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -Fixture tests/fixtures/water_stage_disabled_reference.json -Out tests/tmp/stage_recheck/disabled_ref15 -Resume
```

各コマンドのExcelが自然終了してから次へ進みます。時間でSRMを打ち切りません。

保存記録の独立照合はExcelを起動しません。ZIPを `tests/tmp/stage_retained` へ展開した場合：

```powershell
python -X utf8 tests/check_stage_ui24.py --root tests/tmp/stage_retained --generation 15 --book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm --baseline tests/tmp/stage_retained/reference/acceptance20 --davis-baseline tests/tmp/stage_retained/reference/davis/result.json
```

`calculated_workbooks.zip` は施工水位変更と、同一メッシュの斜面SRM（INCONSISTENT／DAVIS）の3ブックを含みます。DLLは共通の1個です。開いて最終保存状態の図を確認できます。施工水位変更は保存再開の検証後のブックです。最初から再計算するときは「解析・保存 → 保存・再開」で再開元を空欄に戻し、新規の出力先を使います。入力変更・メッシュ再生成をすると、この保存結果と一致しなくなります。

画面画像は実Excelから取得したものです。`mesh_generated.jpg` は既存形状から再生成した1,261節点・392要素で、配布時に保持する1,005節点・306要素の入力メッシュとは区別します。`result*.jpg` と `water*.jpg` の結果画面は、85節点・20要素の水位1m斜面SRMです。
