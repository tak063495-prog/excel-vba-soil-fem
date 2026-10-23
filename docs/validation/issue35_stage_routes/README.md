# Issue35：工程別出力を取り出す実Excel確認

[独立照合](checks.json)／[入力と生結果](records.zip)。配布XLSM SHA256 `508c99ef084129a0b70e01c28dd1d1e75de23a0ef78e212919f044998e0223bd` の専用コピーを使用。VBA・DLLは変更せず、実Excelでコンパイルと全ソース不変を確認。

「解析・保存」C4＝保存・再開→D8＝ALLの変更イベントから設定表への反映、解析、結果C4＝節点結果／要素結果→元の結果表を開く、元表A/BのAutoFilter抽出まで実行。道路11工程の節点3608行・要素879行、トンネル通常9工程の節点4473行・要素1120行を保存。道路1/4/6/11の節点19と、トンネル1〜9の地表351を13条件で抽出した。元表の番号列が小数書式のため、フィルター候補の表示値で選び、必要ならA:B列幅を広げる。

この記録は最新版での操作・保存経路の確認。PDF図116-A・10-1は出典を保持した旧合格記録からの再集計で、今回の新規結果へすり替えていない。旧結果との全節点差をchecks.jsonに記録した。道路のY変位は一致、トンネルの最大Y変位差は0.101353 mm、地表追加沈下は旧5.689765 mm／今回5.694450 mm。異なる版の同値性を主張しない。SRM工程10はこの操作試験では実行せず、通常施工9工程だけを取り出した。

再現（リポジトリルート、専用Excelのみ）：

```powershell
python -X utf8 tests/check_manual_stage_routes.py --prepare
pwsh -NoProfile -File tests/run_manual_stage_routes.ps1 -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -InputJson tests/tmp/issue35_routes/road_input.json -Out tests/tmp/issue35_routes/road_new
pwsh -NoProfile -File tests/run_manual_stage_routes.ps1 -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -InputJson tests/tmp/issue35_routes/tunnel_input.json -Out tests/tmp/issue35_routes/tunnel_new
python -X utf8 tests/check_manual_stage_routes.py --road road_new --tunnel tunnel_new
```

元の入力・メッシュ・材料・工程は公開済みgzipから準備する。出力先の既存記録を上書きしない。実行中の別Excelがあれば起動を拒否する。倍精度・材料定数・収束条件・SRM条件は変更していない。
