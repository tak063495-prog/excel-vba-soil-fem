# INCONSISTENT の材料・求解・失敗判定テスト

Windows版Excel、VBAプロジェクトへのアクセス許可、NumPyを導入したPythonを使います。配布ブックの作業コピーだけへ検証用コードを挿入し、配布原本を保存しません。

```powershell
powershell -NoProfile -File tests/inconsistent/run.ps1 -PythonPath python
powershell -NoProfile -File tests/test_practical_workflows.ps1 -FlowPolicy INCONSISTENT
```

|検証|条件数|独立計算・確認対象|
|---|---:|---|
|spectral_points|75|6降伏面のactive-set計算、面・稜線、回転、面外順位、apexの内部・境界|
|material_substeps|50|4/8分割の独立応力更新の合成、全写像の差分、片側微分、接線精度が不足する場合の棄却|
|nonsymmetric_solver|26|真の残差、pivot・fill・拘束、直接求解/GMRES、特異・容量の拒否|
|regularized_solver|28|NumPyの密行列との比較、列スケール、主行列と主診断の保持|
|failure_classification|6|線形非収束・特異・容量を探索/固定Fsへ注入し、偽の上限とFSSを作らないこと|
|finalization|2|最終状態の復元・再解析失敗時に過去のFSSを無効化すること|
|watchdog|9|合成補正、棄却時の復元、致命的停止理由、Newton/LMの99・100回境界|
|fd_guard|2|中心/片側の方向が幅間で異なる接線を原子的に棄却すること|
|合計|198|材料と数値実装の確認。斜面の力学的破壊を認定する試験ではありません|

`spectral_reference.py` はVBAの主応力順位・分岐を使用せず、物理成分の6面を列挙します。`substep_reference.py` は独立の応力更新を繰り返して塑性ひずみと接線を照合します。失敗注入テストとwatchdog制御テストは、状態復元・予算・診断の確認のために合成系を使います。

各証跡には入力ブックのSHA256を記録します。引張側のψ=0復帰不能や分解能を確保できない差分を、期待した棄却として含めています。実行フォルダーのXLSMには検証用コードが含まれるため、解析配布版として使わないでください。

INCONSISTENTによる施工・載荷の10小モデルと30表示は別途確認します。標準の `tests/run.ps1` はDAVISの回帰です。いずれも任意の実務モデル、メッシュ依存性、限界近傍のSRMを保証しません。[修正報告と実測証跡](../../docs/INCONSISTENT_REPAIR.md) を参照してください。
