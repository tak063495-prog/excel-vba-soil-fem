# 水条件検証記録（acceptance20）

本記録の範囲は [GROUNDWATER.md](GROUNDWATER.md) に記載した水条件拡張です。比較元はmainの`a444247`（20261022 UI_06）です。配布ブック`DF0BC434…A3A62ED`そのものを対象に、全回帰・独立検算・Excelコンパイル・保存再読込を確認しました。水条件の初期値はNONE、従来の材料・節点・要素・工程入力を保持し、DLLは従来と同一の1個です。

## 数値回帰

acceptance20では乾燥系と水条件系のfixtureがPASSしました。乾燥比較は170自由度の全DOFを対象に、`INCONSISTENT` と `DAVIS` の両flow、およびfull SRMを含みます。10行の変位比較はすべて `max_du=0` です。残差は記録されたsolver精度内で、DAVISは `1.05e-15`〜`5.53e-8`、INCONSISTENTは `9.44e-16`〜`1.94e-9` でした。full SRMの記録区間はDAVIS `1.475–1.4875`、INCONSISTENT `1.550–1.5625` です。

独立計算ではdry、hydro、submerged、steady、totalの沈下が解析解と一致しました。最大の解析解誤差はsubmerged/steadyの `6.49e-15 m` です。boxでは水圧合力 `Fx≈0`、`Fy=90.0 kN`、モーメント `135.0 kN·m`、鉛直反力 `102.0 kN`、再開時の変位差 `4.91e-13 m`、残差 `2.48e-9` でした。cutoffでは圧力誤差 `6.26e-10 kPa`、流束誤差 `2.53e-15 m/s`、水頭誤差 `1.41e-10 m`、局所収支 `9.74e-15` です。separatedの独立cutoffではそれぞれ `5.24e-11 kPa`、`6.36e-18 m/s`、`2.48e-13 m`、`2.08e-15` でした。

`OFF`/`AUTO`の170自由度比較はNATIVE_DLLバックエンドの比較です。変位差は `1.11e-16`、残差は `4.98e-15` と `4.50e-15` でした。これは外水圧のAUTO切替比較ではありません。inside-waterは `Fy=85.0 kN`、`M=127.5 kN·m`、反力 `107.0 kN` でした。stage、initial-birth、all-birth、wet-SRMの連続計算と再開計算のmetricsも記録精度内で一致しました。計測時間（連続/再開）はstage `6.622/2.051 s`、initial-birth `4.715/1.591 s`、all-birth `4.715/1.569 s`、wet-SRM `137.598/2.024 s` です。これは一回の計測値であり性能保証ではありません。

証拠：[comparison.json](groundwater-evidence/20261022/comparison.json)、[independent.json](groundwater-evidence/20261022/independent.json)、[separated_independent.json](groundwater-evidence/20261022/separated_independent.json)、[figure.json](groundwater-evidence/20261022/figure.json)。画像：[input.png](groundwater-evidence/20261022/input.png)、[preview.png](groundwater-evidence/20261022/preview.png)、[result.png](groundwater-evidence/20261022/result.png)。

## 保存再開と判定値

チェックポイントの `gauss.csv` は応力と塑性状態を保存します。ひずみCSVは保存せず、再開時に保存された `TDisp` と再構築した `B` から `B*TDisp` で全ひずみを再計算します。したがって、ひずみの再開判定はCSVのひずみ列の一致ではなく、再構築後の値を比較します。

受入れ判定の許容差は、物理的な応力・ひずみが絶対 `1e-6`、`Stmat(9)` が2e9級CSV sentinelを含むため `1e-6 + 5e-15·maxAbs`、変位が `1e-9`、残差が `1e-7` です。修正後のbox再開では、全ひずみの最大差は`2.85e-12`、応力成分は約`1e-12`です。2e9級の降伏関数の約`4.53e-6`差は従来CSVの約15桁への整形で説明され、計算の収束条件は変更していません。

各result.jsonの実際の `book_sha256` をブック識別値として使用します。dry比較のメタデータに存在しないreference hashを推測して補いません。`figure.json` はExcel 16.0 build 20430、compile PASS、figure PASSを記録しています。

## 再現方法

リポジトリルートで `numpy` をインストールし、次を実行します。

```powershell
pwsh -NoProfile -File tests/run_groundwater_regression.ps1 `
  -Book workbook/ui_20261022/2DSoilFEM_ui_20261022.xlsm `
  -ReferenceBook tests/tmp/reference_main.xlsm `
  -Out tests/tmp/groundwater_repro
```

比較元は次のようにバイト列で抽出します。DLLは配布ブックと同じ場所から参照します。Excelはすべて閉じ、専用コピーで実行してください。

```powershell
python -c "import pathlib,subprocess; p=pathlib.Path('tests/tmp/reference_main.xlsm'); p.parent.mkdir(parents=True,exist_ok=True); p.write_bytes(subprocess.check_output(['git','show','a444247:workbook/ui_20261022/2DSoilFEM_ui_20261022.xlsm']))"
Copy-Item workbook/ui_20261022/soilfem_native_x64.dll tests/tmp/soilfem_native_x64.dll
```

上記の抽出後に回帰コマンドを実行します。Outには未作成のフォルダーを指定します。

## 制限

|項目|検証範囲と制限|
|---|---|
|浸透|飽和・定常・固定境界・共有節点の水頭連続。自由水面、不飽和、過渡解析は対象外。|
|力学|水圧を力学へ一方向に渡す。水理と力学の反復、接触による浮上解析は対象外。|
|材料|`TOTAL`/`EFFECTIVE`を入力で明示し、強度値を自動変換しない。|
|再開|応力・塑性状態はチェックポイント、ひずみは`B*TDisp`再構築。CSV精度と判定許容差が一致範囲を決める。|
|実Excel表示|水条件入力・境界表・水圧プレビュー・結果図・保存再表示・入力変更時の失効を専用コピーで確認。従来のメッシュと変形・応力図、7種類の増分図fixtureも確認。Excel16.0 build20430であり、Excel2016の製品版固有の再試験を意味しない。|


## 根拠ファイル

[全ログ・全変位・応力・水頭・流量・チェックポイント](groundwater-evidence/20261022/records.zip)、[ブック・DLL・入力・ソースの識別値](groundwater-evidence/20261022/provenance.json)、[配布ブックのコンパイル・再読込・入力保持](groundwater-evidence/20261022/package.json)、[実画面の確認記録](groundwater-evidence/20261022/ui_checks.json)。

|水なしの全SRM|main [s]|配布版 [s]|Fs区間・全170変位・残差|
|---|---:|---:|---|
|INCONSISTENT|20.628|21.812|完全一致、Fs 1.550–1.5625|
|DAVIS|5.343|5.243|完全一致、Fs 1.475–1.4875|

各1回の計測です。高速化の評価や同一処理時間の保証ではありません。水あり137.598 sは別条件の例なので、上表との差を水圧処理だけの時間と解釈できません。
