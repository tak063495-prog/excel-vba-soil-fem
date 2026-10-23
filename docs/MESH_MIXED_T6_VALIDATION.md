# Q8/T6混在対応の検証記録

計算識別`20261024_MIXED_01`。64bit Excel 16.0 build 20430で実行。Double、材料則、収束条件、SRM探索条件を維持し、DLLは従来と同じ1個です。配布ブックのSHA256は`34c9e4b8ed7ad1a7149139fd70be557ea58efea85b7391aae934a4663c21ac2d`です。

[入力・対応範囲](MESH_MIXED_T6.md)／[検証メタデータ](mixed-t6-evidence/20261024/validation.json)／[全実行ログ・結果表・再開記録](mixed-t6-evidence/20261024/accepted_records.zip)。ZIPのCHECKSUMS.jsonで各記録を照合できます。配布原版には検証用マクロを追加していません。

## コンパイル・入力・表示

- VBEの実コンパイル、保存前と通常保存後の再オープンで全対象モジュールをソース照合。[ビルド記録](mixed-t6-evidence/20261024/compile_package.json)。
- 元の節点・要素A:J・材料・ステージ・載荷・形状入力と要素K列のデータを保持。新しいK1見出しとメッシュ設定C30を追加し、初期OFFを実ブックで確認。
- [メッシュ操作 19 項目](mixed-t6-evidence/20261024/mesh_ui_checks.json)。ON/OFFの実イベント、正常Q8の保持、元の節点条件、保存再オープン、共有中間節点の不一致、入力行順、前モデルの寸法の持越し、地形格子／デローニの修正後T6置換、不正形状と入力保持を確認。
- 最大角約177°のQ8を170°基準で不合格とし、対角線2–4によるT6へ置換。正常Q8×1とT6×2、14節点。XMLを直接読んだ独立検算39項目で、面積1.5＋0.5＋0.5=2.5m²、γ=18kN/m³、厚さ1m、自重45kN、共有辺・中間節点・正のJacobian・未使用節点なしを照合。[独立結果](mixed-t6-evidence/20261024/mesh_independent.json)。実Excel反力45kN、自由DOF残差1.319e-15。
- [既存地形メッシュ 19 ケース](mixed-t6-evidence/20261024/terrain_checks.json)。T6初期OFFで段差・複数材料・円形近似トンネル・テーパー擁壁の生成、材料面積、接続、中間節点、想定エラーを独立照合。これらはメッシュ回帰であり、各形状の力学解析を行った件数ではありません。
- [7画面 102 項目](mixed-t6-evidence/20261024/stage_ui_checks.json)と[結果図 11 項目](mixed-t6-evidence/20261024/output_checks.json)。実ボタンのメッシュ確認と変形・応力図、増分A/B、せん断増分・塑性クロス、塑性状態／累積ひずみ、通常再オープン、水頭等値線・Darcy矢印を確認。画面も目視し、T6凡例の重なりを修正。

## 基準解との比較

2m×1mの平面ひずみパッチ。E=10000kPa、ν=0.3、X変位固定、底面Y固定。拘束弾性係数 M=E(1−ν)/[(1+ν)(1−2ν)]。自重γ=18の解はuy=−18(y−y²/2)/M、鉛直面荷重100kPaの解はuy=−100y/M。反力はそれぞれ36／200kNです。定常浸透は全外周にh=3−0.5x、K=1e-4m/sを与え、p=9.81(h−y)、qx=5e-5m/s、qy=0と照合しました。

| ケース | 時間(s) | 力学残差 | 基準解との差の最大値 |
|---|---:|---:|---|
| q8_gravity | 5.223 | 6.563e-16 | 9.010e-15 m |
| q8_traction | 5.208 | 1.313e-15 | 1.163e-13 m |
| q8_seepage | 5.467 | 3.939e-16 | h=0.000e+00 m / q=1.423e-19 m/s |
| mixed_gravity | 5.205 | 8.429e-16 | 1.026e-14 m |
| mixed_traction | 5.302 | 8.456e-16 | 1.083e-13 m |
| mixed_seepage | 5.446 | 3.566e-16 | h=0.000e+00 m / q=8.882e-20 m/s |
| t6_gravity | 5.267 | 1.056e-15 | 1.000e-14 m |
| t6_traction | 5.209 | 1.520e-15 | 4.337e-18 m |
| t6_seepage | 4.623 | 9.647e-16 | h=4.441e-16 m / q=1.561e-19 m/s |

[全数値](mixed-t6-evidence/20261024/numerical_checks.json)。小パッチは精度確認用で、時間は単回実測です。速度向上率を示す試験ではありません。別の独立計算5試験で、T6の剛性を標準三角形3点積分と照合し、面積積分・剛体3モード・剛性ランク9・線形ひずみ・一定水頭勾配を確認しました。

## 非線形・SRM・保存再開

- 17ケース＋3件の外部チェックポイント再開。JOINT水圧とBIRTH→DEATH→BIRTHの最終反力は88kN。再開後の全変位差はそれぞれ3.441e-22／4.500e-22m。
- 乾燥・非関連流れの塑性載荷で累積塑性乗数最大0.00181336678を確認し、DLL材料評価・接線組立を実際に使用。AUTO/OFFの全変位・応力差は0。時間は4.509／5.210秒。
- 水位3m斜面のINCONSISTENT SRM：元Q8と新版Q8は全変位・全17応力関連フィールド・Fs・最終残差が完全一致。Fs区間は[1.5500, 1.5625]、残差2.483e-14。時間は元版142.536秒、新版142.882秒。
- 同じ斜面の1個のQ8を2個のT6へ分けた混在例もFs区間[1.5500, 1.5625]。最終残差7.976e-10、時間208.342秒。再開後の全変位差最大4.972e-13m。未収束側は数値上の上側境界であり、物理的な破壊の確定値ではありません。
- メッシュを変えたQ8／混在間では共通85節点の変位ベクトル差最大0.00564926m。最大変位は1.43818943／1.43773826mです。この差は要素形式変更に伴うもので、元Q8の回帰一致とは分けて記録しています。

混在SRMの208.342秒は、再開を検証するためSTAGE保存を有効にした実測です。Q8との時間差にはメッシュ、非線形反復の経路、保存条件の違いが含まれます。

外部再開形式は増分図用の最終収束増分を復元しません。再開後の増分表示が停止することを確認し、そのログをZIPのlimitationsに残しました。増分図は最初から計算した保存結果でA/B・通常再オープンを検証しています。この表示用再計算（208.756秒）の全変位・全応力・Fs・残差も最初の混在SRMと一致しました。曲線辺T6、混在メッシュの適応再分割／SMOOTH、fixed-stress時間圧密は今回の対応範囲外です。

[増分Aの実画面](mixed-t6-evidence/20261024/mixed_srm_A.jpg)／[増分Bの実画面](mixed-t6-evidence/20261024/mixed_srm_B.jpg)／[定常浸透の実画面](mixed-t6-evidence/20261024/mixed_seepage.jpg)。

## 再現

Excelを閉じ、リポジトリ直下で実行します。Pythonのnumpy／pytestとPowerShell 7を使用します。各Outは新規フォルダーにします。元Q8の比較にはmainの更新前ブックを別フォルダーへ保存し、同じDLLを置いてReferenceBookへ指定できます。

```powershell
python -m pytest tests/test_mixed_t6_reference.py -q
pwsh -File tests/run_mixed_t6_mesh.ps1 -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -Out tests/tmp/t6_mesh
python tests/check_mixed_t6_mesh.py --book tests/tmp/t6_mesh/mixed_t6.xlsm --out tests/tmp/t6_mesh/independent.json
pwsh -File tests/run_mixed_t6_regression.ps1 -Book workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm -Out tests/tmp/t6_numeric
pwsh -File tests/run_mixed_t6_output.ps1 -SrmBook tests/tmp/t6_numeric/mixed_slope_srm/fresh_results.xlsm -SeepageBook tests/tmp/t6_numeric/mixed_seepage/mixed_seepage.xlsm -Out tests/tmp/t6_figures
```

複雑地形の回帰はtests/run_complex_terrain.ps1とtests/check_complex_terrain.py、7画面回帰はtests/run_stage_ui24.ps1で再現できます。入力はtests/fixtures/mixed_t6のJSON、形状メッシュの操作試験はtests/mixed_t6_mesh_driver.basです。
