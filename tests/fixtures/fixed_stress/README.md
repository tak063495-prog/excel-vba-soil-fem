# Fixed-stress工程入力

`tests/run_fixed_stress_regression.ps1 -Book <1024 XLSM> -Out <新規フォルダー>`で順次実行する。Windows 64bit Excel、VBAプロジェクトへのアクセス許可、およびブックと同じフォルダーの配布DLLを用いる。起動中のExcelがある場合は開始しない。既存の検証記録は上書きしない。

|入力|確認する機能|
|---|---|
|column_smoke|拘束柱の瞬時載荷・時間圧密・保存再開|
|column_biot_anisotropic|α=0.65、非零貯留、異方透水の独立連立解との比較|
|column_nonlinear_srm / nonlinear_control|塑性を伴う圧密、途中SRM有無による施工本線の一致|
|fill_30day|盛土直後SRM、30日圧密、再照査、次の施工|
|water_external_linear|変動する境界水頭と外水圧荷重の一致|
|column_refined|拘束柱の空間・時間細分化（個別実行）|
|watercurve|水位の時間折点・保持・保存間隔|
|closed|閉鎖境界の非排水応答|
|t6|直線T6/P1の時間連成|
|mixed_q8_t6|Q8 7要素＋T6 2要素の混在圧密・保存再開（個別実行）|
|joint_transient|既存JOINTの固定水理接続と時間変動する圧力荷重|
|wet_birth_ramp|湿潤掘削・再盛土・施工時間内の自重載荷・水量持込持出|
|named_unload|名前付き載荷セットと0への除荷|
|saturated_slope_srm|飽和斜面の親水圧保持SRM全探索|
|saturation_rejection|負の受理候補水圧を拒否し、最後の受理状態を維持|
|srm_time_rejection|SRM工程へ時間を指定した入力を拒否|
|pressure_gauge|貯留0・全体積拘束・閉鎖の水圧不定性を拒否|
|rollback_instant|連成反復不足時の未受理状態の巻戻し|

Mandel入力は`tests/generate_fixed_stress_mandel.py --nx 16 --dtau .002 --steps 100 --out <JSON>`で生成する。細分化例はnx=32、dtau=.001、さらにdtau=.00025。解析解比較は`check_fixed_stress_mandel.py`、共通時刻の細分化判定は`compare_fixed_stress_mandel.py`と`mandel_acceptance_policy.json`を使う。上面変位は解析解から与える境界条件であり、自由応答としての検算は全節点uxと内部uy、水圧について行う。

受入で使う`mandel_hybrid_refine.json`は、τ=0.025までΔτ=0.00025、その後τ=0.1までΔτ=0.001とする入力である。17入力の標準回帰とは別に、次のように実行する。

```powershell
pwsh -NoProfile -File tests/run_groundwater_fixture.ps1 -Book <1024 XLSM> -Fixture tests/fixtures/fixed_stress/mandel_hybrid_refine.json -Out <新規フォルダー> -Resume
python tests/check_fixed_stress_mandel.py --fixture tests/fixtures/fixed_stress/mandel_hybrid_refine.json --run <同フォルダー> --report <診断JSON>
```

`mandel_coarse.json`と`mandel_fine.json`も同じ手順で個別に実行する。3つの診断JSONを比較器の`--coarse`、`--fine`、`--early`に与え、`--policy tests/fixtures/fixed_stress/mandel_acceptance_policy.json --out <判定JSON>`で共通時刻の合否を得る。入力の粗密による誤差と分離反復の収束は別々に確認する。

結果ブックと`*_out`の時刻別フォルダーを一緒に保管する。詳細は[操作説明](../../../docs/FIXED_STRESS_STAGES.md)と[受入検証](../../../docs/FIXED_STRESS_VALIDATION.md)。
