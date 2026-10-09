# 文献ケースの再実行

配布版 `20261009_BEST_03_R1` に対する、2026-10-09の検証用入力です。結果の評価・出典・論文と異なる条件は [文献検証報告](../../docs/LITERATURE_VALIDATION.md) を参照してください。SRMの数値PASSは論文との一致や力学的安定の認定ではありません。

## 検証ブックを使う

[ケース一式](../../workbook/validation/literature_cases_20261009.zip) を展開し、`cases` 内の対象XLSMをデスクトップ版Excelで開きます。「図」にはそのケースの初期メッシュを保存しています。材料、節点、要素、載荷、ステージは入力済みです。解析実行で再計算できます。ログはブック横の `ケース名_out` に保存されます。

メッシュは文献形状を独立に座標化したものです。**「メッシュ生成」は押さず、入力済みの節点・要素を使ってください。** 形状・メッシュ設定には元の配布モデルの生成用設定が残り、再生成するとこの検証メッシュにはなりません。流れ則を変えると材料条件も変わります。

## 自動で作成・実行する

Windows、デスクトップ版Excel、PowerShell、Python 3が必要です。VBAテスト関数を一時追加するため、Excelの「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしてください。一時関数は解析後に削除します。ブックを外から編集する独立したExcelインスタンスを使います。

リポジトリのルートで実行します。Python部分の入力生成・照合は標準ライブラリだけを使います。グラフ作成は任意で、ReportLabが必要です。

```powershell
python tests/literature/build_fixtures.py

# 弾性の5載荷条件（短時間）
powershell -NoProfile -ExecutionPolicy Bypass -File tests/literature/run_cases.ps1 -Mode Elastic -FlowPolicy INCONSISTENT

# SRM：2種類のGriffithsメッシュ、3種類のPruška斜面
# 破壊付近は微小増分が多く、長時間かかります。
powershell -NoProfile -ExecutionPolicy Bypass -File tests/literature/run_cases.ps1 -Mode SRM -FlowPolicy DAVIS

# 論文のpsi=0を直接用いる対称接線近似との比較
powershell -NoProfile -ExecutionPolicy Bypass -File tests/literature/run_cases.ps1 -Mode SRM -Names griffiths_1999_coarse -FlowPolicy INCONSISTENT

python tests/literature/analyze_results.py --output tests/tmp/literature_validation
python tests/literature/verify_inputs.py --output tests/tmp/literature_validation

powershell -NoProfile -ExecutionPolicy Bypass -File tests/literature/finalize_cases.ps1 -OutputRoot tests/tmp/literature_validation
powershell -NoProfile -ExecutionPolicy Bypass -File tests/literature/verify_case_books.ps1 -CasesDirectory tests/tmp/literature_validation/cases -ResultsDirectory tests/tmp/literature_validation/results -BaselineWorkbook workbook/2DSoilFEM_20261008_practical.xlsm -OutputDirectory tests/tmp/literature_validation/verification

# 任意：自作の形状図と変位比較図
python tests/literature/plot_validation.py --output tests/tmp/literature_validation
```

`-Names` で1ケースに限定できます。`-SourceWorkbook` と `-OutputRoot` で元ブック・保存先を変更できます。同じ保存先で同じケースを同時実行しないでください。同名の検証用ブック・ログは再実行時に更新されます。

早い非収束の切り分けに使った高速化OFF・固定Fs=1の追加条件は、次で再実行できます。これは安全率の全探索とは別の診断です。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/literature/run_cases.ps1 -Mode SRM -Names griffiths_1999_coarse -FlowPolicy INCONSISTENT -FixedFs 1 -DisableAcceleration -CaseSuffix _baseline_fs1 -OutputRoot tests/tmp/literature_diagnostic
```

## 保存される内容

- `cases/*.xlsm`：元の計算VBAを保持した検証ブック。
- `cases/*_out/`：Excel自身が出したログ。`run.log` はCP932、主なCSVはASCII数値です。
- `results/*.json`：取得した変位・積分点応力、正確なFs区間、状態、条件。`*_state.json` は進行状態なので評価から除外します。
- `comparison.csv/json`：理論式との誤差、論文のFsとの比較。
- `srm_trials.csv`：試行ごとの数値・力学状態と変位。ログのFsは小数3桁なので、正確な端点は結果JSONを使います。
- `verification`：全入力の照合、VBA全文の一致、一時コードの除去、読取確認でファイルが変わらないことの記録。
- `visual_checks`：保存結果の初期図・結果図を開けることの確認と代表PDF。

`fixtures.json` は節点、要素、境界、材料、ステージ、Q8一致節点力、理論解を含みます。Q8の中間節点を共有し、全積分点の正のJacobianと全体面積を確認してから作成しています。載荷のないブックではVBAが値0のDISP例を1行自動補完します。照合はその行の全値とLOADステージがない条件を確認し、実際の追加荷重とは扱いません。

文献PDFは著作権のため同梱していません。出典へのリンクと、比較に用いた条件を `srm_sources.json` に保存しています。

## 材料点の追加診断

7つの拘束圧縮・純せん断条件を独立の理論式と比較します。この処理は保存しない作業コピーを使い、基準ブックを変更しません。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/literature/check_material_points.ps1 -SourceWorkbook workbook/2DSoilFEM_20261008_practical.xlsm -OutputRoot tests/tmp/literature_material
python tests/literature/analyze_material_points.py --output tests/tmp/literature_material
```

配布R1では、独立の7照合はPASSですが、内蔵自己テストの2つの引張入力の期待結果に不整合があり、解析スクリプトは終了値1でこれを報告します。`material_point_checks.json` の `independent_ok` と内蔵試験の結果を分けて確認してください。材料モデル全般の合格を表示する仕組みではありません。

同じ記録に、材料サブステップを使わない接線の中心差分照合を含めています。塑性点からの微小再載荷と2つの差分幅を使い、前進・後退差分と分岐メッセージも保存します。今回、一部の状態で19～22%の差が残りました。稜線では接線の一意性にも注意が必要で、単純な合否閾値で材料モデル全体を認定していません。

自己テスト入力だけの修正候補は、次のように別の保存先のコピーで確認できます。計算部の材料更新・ソルバは変更せず、このコピーの変更も保存しません。修正候補で旧引張例を置き換える場合は、旧例を期待する失敗の試験として残すことも必要です。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/literature/check_material_points.ps1 -SourceWorkbook workbook/2DSoilFEM_20261008_practical.xlsm -OutputRoot tests/tmp/literature_material_trial -RepairSelfTestInputs
python tests/literature/analyze_material_points.py --output tests/tmp/literature_material_trial
```

SRMの幅0.0125付近で丸めにより追加探索する診断は、実際の非公開関数 `NextFsByBracket` を保存しない作業コピーから呼び出します。FEMの載荷計算は行いません。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/literature/check_srm_roundoff.ps1 -SourceWorkbook workbook/2DSoilFEM_20261008_practical.xlsm -OutputRoot tests/tmp/literature_roundoff
```
