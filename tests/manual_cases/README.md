# 説明書の実計算入力

配布ブック `workbook/ui_20261020/2DSoilFEM_ui_20261020.xlsm` と同じフォルダーの単一DLLを、空の出力先へコピーして実行します。元のブック・VBA・DLLは上書きしません。

| 入力 | 内容 |
| --- | --- |
| `slope_wall_shape.json` | 斜面・L型擁壁・背面盛土の形状入力。形状メッシャーによる1857節点・580要素の作成を含む |
| `slope_wall_mesh.json` | 同じ1857節点・580要素の保存メッシュ。アンカー相当荷重の例にも使用 |
| `slope_wall_fine_mesh.json` | 同じ物理形状の3383節点・1074要素。DAVIS / INCONSISTENTの同条件比較用 |
| `tunnel_mesh.json` | 内空天端まで20m、内空径8m、掘削外径8.6m、支保厚0.3m。497節点・160要素の共有Q8メッシュと9施工工程 |
| `road_widening_taller_shape.json` | 道路例の点17/18のYを4→5mへ変更。352節点・103 Q8、11工程。材料8γ=19を維持し、最終支持反力1899kNを独立自重と照合。角度WARN4要素は残存 |

JSONの材料・節点・要素・ステージ・載荷は、各シートのA列からの値です。形状入力は点をB:D、外側境界をG:Jへ配置します。材料はA:M、節点はA:I、要素はA:J、ステージはA:F、載荷はA:Jです。

保存メッシュを入力するケースでは、元ブックの無関係な形状入力を消し、節点・要素表を使います。「形状メッシュを作成」を押して再生成する手順ではありません。既存メッシュの確認から解析へ進んでください。壁の形状を変更して生成し直す場合は形状入力のケースを使い、材料・支持・荷重の対象も再確認します。

リポジトリのルートからPowerShellで実行します。デスクトップ版ExcelとVBAプロジェクトの参照許可が必要です。ほかのExcelが起動中なら処理を開始しません。出力先は存在しないフォルダーを指定します。

```powershell
pwsh -NoProfile -File tests/run_manual_case.ps1 -Fixture tests/manual_cases/slope_wall_shape.json -OutDirectory tests/tmp/manual_wall -Srm -Flow DAVIS
pwsh -NoProfile -File tests/run_manual_case.ps1 -Fixture tests/manual_cases/slope_wall_fine_mesh.json -OutDirectory tests/tmp/manual_inco -Srm -Flow INCONSISTENT
pwsh -NoProfile -File tests/run_manual_case.ps1 -Fixture tests/manual_cases/slope_wall_fine_mesh.json -OutDirectory tests/tmp/manual_davis -Srm -Flow DAVIS
pwsh -NoProfile -File tests/run_manual_case.ps1 -Fixture tests/manual_cases/slope_wall_mesh.json -OutDirectory tests/tmp/manual_anchor -Anchor -Srm -Flow INCONSISTENT
pwsh -NoProfile -File tests/run_manual_case.ps1 -Fixture tests/manual_cases/tunnel_mesh.json -OutDirectory tests/tmp/manual_tunnel -Srm -Flow INCONSISTENT
```

各コマンドは前のExcelが終了してから順に実行します。時間による打切りはありません。倍精度・収束条件・SRM探索設定は配布値を保持し、上記の流れ則と出力設定だけを明示します。`-StageLimit 4` のように指定すると、その工程までを有効にして最初から再計算できます。SRMを追加しない場合は `-Srm` を省きます。

`-Anchor` は専用アンカー要素ではありません。壁側3節点までと地盤の定着側節点群に、重心間の方向に沿う等大反対向きの合力50kN（100kN/本・間隔2m・奥行1m）を分配し、LOAD工程を追加します。軸剛性・引抜き・自由長等は表現しません。

出力は `result.json`（入力表・全工程結果・診断・設定・実測秒）、`case_out/`（CSVログ）、入力状態の `case.xlsm`、計算後の `2DSoilFEM_ui_20261020.xlsm`、表示比較の `A.xlsm` / `B.xlsm` です。`A.xlsm` / `B.xlsm` は図の保存用で、通常操作には元の名称で保存した計算後ブックを使います。

実行前のVBAコンパイル、実行後の全非空コードモジュール照合、元ブックのSHA-256照合を行います。自然終了の失敗・未確定も結果として保存します。`PASS`とFs区間の意味は別であり、LOWER_BOUNDやUNDETERMINEDを物理破壊安全率へ読み替えません。

完了したケースはPythonとNumPyで独立検算できます。

```powershell
python tests/check_manual_case.py tests/tmp/manual_wall/result.json
```

Q8形状関数の4積分点と11×11点で基準・倍率1変形後のJacobian、固定節点変位、増分の主値、自重・反力、構造材の非塑性化、アンカー相当力の合力・モーメントを確認します。`independent_checks.json` と `balance_checks.json` を同じ出力先へ保存します。計算自体の失敗・未確定は、先に診断と停止理由を読み取ってください。

独立釣合い検査では、支持点の反力だけでなく、結果節点の反力列に保存された自由節点残差の合計も含めて自重と照合します。支持反力と自重の差は別途記録します。

細メッシュINCONSISTENTの最終数値収束状態には、倍率1の変形後形状で一部要素の反転があり、通常の厳格な形状検査は失敗します。これは表紙・A/B見本には使用していません。反転箇所も含む診断を保存する場合のみ、次を指定します。反転がある場合の終了コードは2で、合格にはなりません。

```powershell
python tests/check_manual_case.py tests/tmp/manual_inco/result.json --report-invalid-geometry
```

公開した4工程の擁壁ペアについては、同条件・流れ則のみの差と全変位を次のコマンドで再照合できます。

```powershell
python tests/compare_manual_flows.py --inco tests/tmp/manual_inco/result.json --davis tests/tmp/manual_davis/result.json --out tests/tmp/manual_flow_comparison
```

変形後形状の反転は比較JSONの警告として保存します。他の独立検査の失敗は処理を停止します。


## 道路拡幅と水平・斜面の支持力例（追加改訂）

`road_widening_shape.json` は道路拡幅の11工程です。現行配布原版のコピーで形状メッシュを作成し、初回盛土→盛土の一部と基盤の掘削→MATSET→別材料・別形状の再盛土を再計算しました。

```powershell
pwsh -NoProfile -File tests/run_manual_case.ps1 -Fixture tests/manual_cases/road_widening_shape.json -OutDirectory tests/tmp/manual_road_new -Flow INCONSISTENT
python tests/check_manual_case.py tests/tmp/manual_road_new/result.json
```

`generate_bearing_cases.py` は水平地盤・斜面の粗密共有Q8メッシュを生成します。E=30000kPa、ν=.30、厚さ1m、γ=18kN/m³、c=100kPa、φ=ψ=0°の例です。底面XY固定・側面X固定、幅2mの載荷帯へ自重後の追加沈下を与えます。剛体基礎・片側接触・離間はモデル化しません。

```powershell
python tests/manual_cases/generate_bearing_cases.py
python tests/run_bearing_cases.py --out tests/tmp/manual_bearing_new
```

ドライバーは粗メッシュ1/5/10/20/50mm、細メッシュ10/20/50mmを、水平・斜面の合計16ケースで順次実行します。各ケースを初期状態から再計算し、倍精度・収束条件を維持します。ほかのExcelがある場合は起動せず、所有Excelの終了を待って次へ進みます。中途半端な出力先へ自動で上書きしません。再開時には同じ実行計画と入力ハッシュを要求します。

テンプレートはLOAD工程をE=0にしてあります。自重後の各対象節点のuy0を先に取得し、派生fixtureでDISP目標をuy0−追加沈下へ置換してE=1にします。DISPへ追加沈下だけを入れると、地盤を引き上げる入力になり得ます。通常実行ではfixtureの明示的な無効工程を保持し、`-StageLimit`を明示した場合だけ指定工程までを有効化します。

```powershell
python tests/check_bearing_case.py tests/tmp/manual_bearing_new/bearing_flat_s010000/result.json
```

独立検算では、元の支持変位、全Q8の4点/11×11点Jacobian、各載荷節点の追加沈下、反力Y（結果節点J列）、自重・自由節点残差・支持反力、最終相対残差を照合します。H列の外力Yを基礎反力として集計しません。q=−ΣR_y/(B×厚さ)は指定沈下時の平均圧力であり、極限・許容支持力の自動判定や現場載荷試験値ではありません。

公開記録を別の場所へ展開した場合は、`--fixture`と`--baseline`で派生入力JSONと対応する自重結果JSONの位置を指定できます。圧縮ファイルは元のバイト列へ展開してください。baselineのSHA-256も照合します。
