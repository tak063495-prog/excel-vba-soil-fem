# 斜面SRM：既存INCO_24の8検証例

説明書の第3章（37–44ページ）へ、既存のGriffiths–Lane/Pruška再構成斜面の実計算を統合しました。これは2026年10月10日に完了したINCO_24の記録で、現行UI_04の新規解析を主張するものではありません。

- [固定入力](../../../../inconsistent-evidence24/fixtures.json)
- [Fs区間と数値診断](../../../../inconsistent-evidence24/practical_comparison.json)
- [独立材料内力・全3D降伏検算](../../../../inconsistent-evidence24/material_force_audit.json)
- [8ケースの実行ブック・元ログZIP](../../../../inconsistent-evidence24/srm_cases_INCO24.zip)
- [今回の8ケース集計](cases.csv) / [出典SHA・図の形状検査](summary.json)

リポジトリのルートで `python docs/manuals/construction-evidence/20261021/slope-srm-guide/rebuild_slope_srm.py` を実行すると、公開済みの固定入力/結果/検算JSONと元ZIPだけから図と集計を再作成します。必要ライブラリはNumPy、Matplotlib（日本語フォントYu Gothic）、既存のtests/check_manual_case.pyです。Excelは起動しません。

Pruška H7のφ10/20は節点・要素・施工・φ以外の材料が同一であることを検査しています。φ0の例はcと形状も異なり、cだけの感度実験とは扱いません。全てINCONSISTENT・ψ=0。φ=ψの関連流れと、DAVISの等価関連流れへの換算も説明しますが、新規のφ>0・ψ=φ対照計算はありません。

total_shear.pngはG φ20の成功下端2状態の**全変位から求めた面内せん断ひずみ**です。最終荷重増分図ではありません。倍率1で11×11点の変形後Q8 Jacobianが正で、変位ノルムも既存検算と一致することを確認しました。その他の大きな変形を持つ状態は物理的な移動量の見本にしていません。数値Fs区間は厳密な力学的上下界ではなく、旧版同時実行の時間は現行DLLの速度比較には使えません。

cover_initial_mesh.pngは初版と同じ1,005節点・306Q8の斜面入力メッシュです。解析後の変形図・せん断帯ではありません。
