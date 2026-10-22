# 受入ケース台帳

個別JSONの `status=PASS` と、最終source22で保存後再読込した結果署名・ノード/要素結果表が確認できたものだけを「合格」とします。最終ファイルの `comparison.json` のケース名、例題ZIP内のファイル名、個別検算記録を対応付けます。最終回帰・保存再読込の合格14ケースを収録しました。水条件NONEのdryは水圧署名の対象外です。

|ケース|検証内容|例題ZIP|個別記録|
|---|---|---|---|
|dry|乾燥JOINT、接触ばねのみ|`joint_water_dry.xlsm`|[dry.json](dry.json)|
|wet|土―構造の湿潤JOINT、地盤側の接合面水圧|`joint_water_wet.xlsm`|[wet.json](wet.json)|
|prescribed|PRESCRIBEDプロファイル|`joint_water_prescribed.xlsm`|[prescribed.json](prescribed.json)|
|continuous|SOIL―SOILの水頭連続|`joint_water_continuous.xlsm`|[continuous.json](continuous.json)|
|prescribed_steady|定常水頭付きPRESCRIBED|`joint_water_prescribed_steady.xlsm`|[prescribed_steady.json](prescribed_steady.json)|
|soil_steady|土側定常水頭|`joint_water_soil_steady.xlsm`|[soil_steady.json](soil_steady.json)|
|partial|乾湿境界を含む区間積分|`joint_water_partial.xlsm`|[partial.json](partial.json)|
|slip|接線滑りと非対称接線|`joint_water_slip.xlsm`|[slip.json](slip.json)|
|opening|開口時の接触力ゼロと水圧維持|`joint_water_opening.xlsm`|[opening.json](opening.json)|
|stages|OFF→BIRTH→DEATH→BIRTH工程|`joint_water_stages.xlsm`|[stages.json](stages.json)|
|srm|JOINTを含む固定Fs=1.2照査|`joint_water_srm.xlsm`|[srm.json](srm.json)|
|struct_pair|構造体―構造体の指定水圧|`joint_water_struct_pair.xlsm`|[struct_pair.json](struct_pair.json)|
|wet_davis|湿潤JOINTのDAVIS材料則|`joint_water_wet_davis.xlsm`|[wet_davis.json](wet_davis.json)|
|slip_davis|DAVISでの接線滑り|`joint_water_slip_davis.xlsm`|[slip_davis.json](slip_davis.json)|

## 拒否ケース

`missing_mode`、`inconsistent_pressure`、`conflicting_heads`、`wrong_soil_mode` は、入力を受け付けずエラーとなることを確認した負例です。これらは例題ZIPには含めていません。詳細は `comparison.json` の `rejected` と実行ログを参照してください。

## 保存物の使い分け

- **説明書掲載用**：例題ZIPの14ブック、`joint_input.jpg`、`joint_result.jpg`、`water_input.jpg`、`water_result.jpg`、`preview.jpg`、`mesh.jpg`、`reopen.jpg`、各個別JSON。
- **再現・監査用**：`records.zip`、`compile.log`、`package.json`、`provenance.json`、`source_audit.json`。
- **比較用**：`comparison.json` の `unchanged_groundwater`。旧WATER_01との差分が0であることを確認するための記録です。
- **除外**：診断開始前のクリア状態、保存前の試行、失敗・中断・旧版の `debug*.xlsm` や `ui18*`。これらは作図や数値表の根拠にしません。

最終保存再チェックは、受入ケースについて最終状態の保存後再オープン、結果署名、ノード・要素結果表、変位・物理場の差分を確認したものです。途中のステージ・増分履歴はログ側に保持し、再開後の結果表と混同しません。最終回帰で合格したケース数とSHA-256は、更新された `comparison.json`、`saved_results.json`、`provenance.json` を正とします。ZIP内の `README.txt` と本台帳を併読してください。
