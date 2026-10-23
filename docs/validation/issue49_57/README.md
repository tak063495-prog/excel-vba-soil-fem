# Issue #49〜#57の計算記録

[修正内容と実測値](../../ISSUE49_57_VALIDATION.md)／[検証集計・SHA256](validation.json)

- [実計算の生記録](records.zip)：21入力、途中LINEAR再開、17連成入力、従来Issue回帰、17混在計算、材料点50条件、荷重・反力表示、別Excelからの再開3条件、入力画面の操作。
- [コンパイルした52モジュール](source_snapshot.zip)：実行時のソースとファイル別SHA256。
- [説明書の確認画像](manual_review.zip)：全234ページの画像、一覧12枚、変更9ページの拡大画像。[ページ・図・参照の照合](manual.json)。
- [最後の荷重引継ぎ修正前の不一致](partial_before_completion_fix.zip)：途中LINEARから継続すると最終反力が57.62 kNへ戻った記録。合格例には数えない。

`records.zip`内の主要フォルダーは次のとおり。

|フォルダー|内容|
|---|---|
|targeted6|Issue入力21個の解析・拒否・再開記録。water_linearのfresh_results.xlsmは専用検証コピー。DLLは配布ZIPの同一DLLを使う。|
|partial3|LINEARのC000003から継続したケース。次のWAIT・除去後の全自由度荷重も検査。|
|reset_final3|2 T6の非アフィンRESET拒否時の、正しい座標取得方法による最終記録。|
|mixed_reset3|Q8/T6混在のRESET拒否と保持状態。|
|full3／legacy3／mixed3|既存の連成、過去Issue、混在要素の回帰。|
|material4|実ExcelのP2材料点出力と独立検算50条件。|
|frame51_2|時刻を後→前→後と選び、表示した全節点力・反力を独立積分と照合。|
|cold3|別のExcelプロセスで、表示等を先に読み込まず直接再開した3条件。|
|ui6|明示した入力を使った、材料・時間・工程編集9項目の合格記録。|

`targeted6/reset_t6_nonaffine`の解析拒否自体は正しいが、初回の座標収集でPowerShell配列が展開されていた。状態保持の照合には`reset_final3`を使う。UI補助試験も、配布初期値に工程数・時間設定を仮定していたスクリプトを直し、明示した検証入力で再実行した。これらのドライバー修正に伴う数値計算コードの変更はない。

記録を任意の`<展開先>`へ展開して、リポジトリのルートから再検算できる。

```powershell
python -X utf8 tests/check_issue49_57.py --root <展開先>/targeted6 --partial <展開先>/partial3 --reset-run <展開先>/reset_final3 --book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm --out <新規集計.json>
python -X utf8 tests/check_issue54_55_material.py --actual <展開先>/material4/material_points.json --out <材料検算.json>
python -X utf8 tests/check_issue51_frame_forces.py --fixture tests/fixtures/issue49_57/water_linear.json --actual <展開先>/frame51_2/frame_forces.json --records-root <展開先> --out <表示検算.json>
python -X utf8 tests/check_issue49_57_extra.py --root <展開先> --out <追加検算.json>
```

生記録の実行時パスは来歴として保持し、表示検算の`--records-root`で展開先へ読み替える。PDFの資料リンクはmainを参照する。
