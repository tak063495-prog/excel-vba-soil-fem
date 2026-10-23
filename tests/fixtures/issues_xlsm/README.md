# XLSM issue再現入力

`tests/run_xlsm_issue_regression.ps1`で配布ブックの専用コピーへ入力するJSONです。数値は不具合の切り分け・独立検算用で、実務用の推奨材料値ではありません。

|入力|確認内容|
|---|---|
|issue45_start_only / hidden_knot / numeric_head|HYDROSTATICの面全区間の不一致をプレビューで拒否|
|issue45_equivalent|ID・折れ点配置が異なっても同じ水位分布なら受理。外水圧合力・モーメント・支持反力を検算|
|issue44_internal_head|共有する内部透水面のHEADを拒否|
|issue44_external_linear / external_layers|外周HEADの均一・2層地盤。全節点水頭、GP水圧・流束、面流量と水収支を解析解と比較|
|issue46_reordered|工程1→3(RESET_STRESS)→2(LOAD)|
|issue46_multiple|工程1→40(RESET)→7(LOAD)→12(RESET)→3(LOAD)|
|issue46_final_reset / no_reset / final_output|最終RESET、RESETなし、FINAL出力の表示・保存再読込・入力変更ガード|
|issue43_water_ramp_cache|総荷重を保持して水位をA=3 mからB=2 mへ変更。材料水位線は`*`で工程指定を継承|
|issue43_steady_stages|水位変更・DEATH・BIRTH後の定常浸透と固定Fs照査。水理接続を意図的に変えた専用保存復元試験にも使用|
|issue43_water_srm|水位W=2 mからB=1 mへ変更後、斜面SRMを自然終了まで探索し保存再開|
|issue26_hold|塑性自重→同じ自重→追加載荷→同じ載荷。保持工程の32 GPで塑性増分ゼロを確認|
|issue38_elastic_control|高粘着力の新規解析対照。同じブックの塑性解析→高粘着力再解析の全変位・応力・診断値と比較|

`issue43_water_srm`だけは`checkResumeAgainstSolverTolerance=true`を指定します。収束境界に近い採用状態の残差は、ブックの既存定数`P3_ENGINEERING_RESIDUAL`を読み取り、新規・再開の双方がその基準内で、残差差が1e-9以下であることを照合します。解析側の基準や他の入力の汎用照合基準は変更しません。

入力・実計算結果・画面・適用範囲は[検証記録](../../../docs/XLSM_ISSUE_VALIDATION.md)を参照してください。

JOINT水理接続の追加入力は `issue44_joint_soil`（地盤・構造体）、`prescribed`（両側地盤で分離した水理境界、解析解照合）、`continuous`（共有内部面のHEAD拒否）、`conflict`（PRESCRIBED固定水頭の衝突拒否）。`run_issue44_joint_regression.ps1` で従来の境界条件も合わせて確認する。
