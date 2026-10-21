# 応用実習：道路の再盛土高さを1m増やす

[計算済みブックZIP](road_taller_workbook.zip)を新規フォルダーへ展開し、同名XLSMとDLL1個を同じ場所へ置きます。UI_05の専用計算コピーで、汎用原版とは別の入力です。352節点/103 Q8、11工程PASS、Fs=1、工程1/11支持反力1440/1899kN。まず保存メッシュを確認し、再生成前に別コピーを残します。旧道路例と取り違えないでください。

点17/18のYを4→5mへ変更した再生成実習です。材料8γ=19を維持。上部面積6→9m²、追加自重57kNを独立に確認しました。角度品質のWARN4要素は残り、精度収束済みや実務設計値とはしません。説明書p.57と[対応表](../../../../reviews/2DSoilFEM_manual_applied_response_20261021.md)を参照。

## ファイル

- `road_widening_taller_shape.json`：生成用の変更後入力。同じ内容を`tests/manual_cases/`へ公開。JSONを読む一般操作ボタンがあるという意味ではなく、手入力の対応表と再現ドライバー用です。
- `result.json.gz`：全入力、全工程の全節点・GP応力/ひずみ・保存増分、診断・設定・実測秒・元XLSMのSHA。アプリコードの不変性を記録。
- `independent_checks.json.gz` / `balance_checks.json.gz`：Q8幾何・倍率1の反転・支持変位・主ひずみ・領域面積・全工程の自重/反力/残差釣合い。
- `exercise_summary.json`：元のUI_05道路記録との比較、変更した設定メタデータ、WARNと検証の限界。
- `geometry_before_after.png`：実生成されたQ8メッシュの比較（辺は中間節点を含む形状関数で描画）。赤印は支持節点。
- `logs/`：CSV・ログをgzip保存。時間・最終残差の原記録。
- `manual_section_source.py`：今回の追記ページと編集処理。旧版の組版へ加える断片で、単独の全冊ビルダーではありません。
- `baseline_release.json`：実計算に使ったUI_05の配布記録。既存汎用XLSMとDLLを変更していません。

## 再現

公開済みgzipからの独立照合（Excelは起動しません）：

```powershell
python tests/check_manual_applied.py docs/manuals/construction-evidence/20261021/manual-applied --check-only
```

実解析は元のXLSMをコピーし、ほかのExcelがあれば開始しない既存ドライバーを使います。Windows ExcelとVBAプロジェクト参照許可が必要です。出力先は未作成フォルダーを指定します。

```powershell
pwsh -NoProfile -File tests/run_manual_case.ps1 -Fixture tests/manual_cases/road_widening_taller_shape.json -OutDirectory tests/tmp/road_taller_repro -Flow INCONSISTENT
python tests/check_manual_case.py tests/tmp/road_taller_repro/result.json
python tests/check_manual_applied.py tests/tmp/road_taller_repro
```

ブックのコード、Double、収束条件、SRM探索条件は変えず、INCONSISTENTと記録用出力設定だけを明示します。この例は施工11工程で、SRMは追加しません。形状変更以外の条件を同時に変えず、再生成後の材料・支持・工程・IDを照合します。材料8γ=20の旧演習は別条件なので混ぜません。
