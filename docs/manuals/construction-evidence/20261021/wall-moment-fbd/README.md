# 擁壁の断面力：自由体の確認（2026-10-21 追加）

- `sections_from_result.py`：`--result` で指定した result.json.gz（自分で計算した新しい結果）の断面力と、切断面を除いた自由体の釣合いを出力します。読んだファイルのパス・SHA-256・工程の行を表示します。`../wall-moment-guide/rebuild_sections.py` は保存済み比較図（20261020 wall_davis08/12 固定）を作り直す専用で、変更していません。
- `ex09_coarse/`：`docs/manuals/examples/ex09_wall_result_20261021.json.gz`（粗い580要素、工程3）の出力。
- `davis12_fine/`：`construction-evidence/20261020/wall_davis12/result.json.gz`（細かい1074要素、工程3）の出力。
- fbd.json の `moment_imbalance_residual` が説明書の「モーメントの不釣合い残差」（細かいメッシュ底版 −18.955 kN·m/m）、`M_cut_integrated` が断面M、`M_cut_from_other_faces_and_weight` が切断面以外の3辺＋自重から求めたMです。
