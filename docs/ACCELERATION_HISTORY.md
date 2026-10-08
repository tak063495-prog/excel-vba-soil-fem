# VBA版 2D Soil FEM 高速化トライアル履歴

作成日: 2026-10-08。ここでは、同梱ログで確認できる値だけを記載する。時間は各版1回の本番測定を含み、再現性を示す平均ではない。数値 `CONVERGED` と力学状態 `LIMIT_STATE` は別の判定である。

## 1. 記号と試験条件

計画資料の呼称 V1～V5 と、後続のファイル名 `policy` は同じ番号体系ではない。`policy` は適応選別の版名であり、内部ログの `ADAPT_03` などが正式な履歴番号である。対応は次のとおり。

| 表記 | 意味 | 実測・位置付け |
|---|---|---|
| V0 | 全OFF基準 | 本番基準 4130.398 s。SRM: PASS 2.2625、FAIL 2.275、最終保存 Fs=2.2625、幅0.0125。 |
| V1 | 同一ステージ内の変位予測子 | 単独実測は `V1_V4_evaluation_20261006.txt` に総時間 3208.773 s として引用。元の試行表は単独の精密列を含まない。 |
| V2a | 増分間の旧LUを近似Jacobianとして使う限定再利用 | V1+V2aで 3096.354 s。小モデルで長反復再利用も確認。 |
| V2b | 残差改善率・実測費用による更新選別 | 前回V2系列で有効化されたが採用0。単独の現存元ログは欠落。 |
| V3 | 保護付きFs探索 | 当初試行済み。現存資料で総時間を再確認できない。 |
| V4 | 深さ1 Anderson 加速 | V1+V4、V2a+V4、V1+V2a+V4を実測。単独も当初試行したが、現存資料では総時間を再確認できない。 |
| V5 | 旧LU前処理付きVBA GMRES | 当初途中試行ログを受領。完走総時間は未確認。小さな既知解・残差・反復上限テストはPASS。 |

試験は主に 1005節点・306要素・1224積分点の SRM Case 系列で、V1/V2a/V4 と adaptive は V2b/V3/V5 OFF、`BAND_LU NEW`。同じ構成で小モデル（85節点・20要素、4条件、固定Fs）も境界回帰に使った。収束許容値、材料更新、メッシュ、SRM判定を緩める変更は採用していない。

## 2. 単独・組合せの実測

| 構成 | 総時間 | 実測した判定・最終値 | 採否 |
|---|---:|---|---|
| V0 | 4130.398 s | PASS 2.2625 / FAIL 2.275、保存 Fs 2.2625、幅0.0125 | 基準 |
| V1 | 3208.773 s（別評価からの引用） | 単独の詳細CSVは未同梱。 | 候補。単独値は単発引用。 |
| V2（V2a+V2b） | 3538.613 s（組合せ評価からの引用） | LU6589。V2b採用0。単独の元ZIPは現存しない。 | V2aのみとの条件差を含むため分離評価できない。 |
| V1+V2a | 3096.354 s | PASS 2.275 / FAIL 2.2875、幅0.0125、LU5496、Newton13082 | 速度優先の暫定候補 |
| V1+V4 | 3568.928 s | PASS 2.275 / FAIL 2.2875、幅0.0125。V4採用440、棄却986 | 見送り。V1+V2aより472.574 s遅い |
| V2a+V4 | 3680.353 s | PASS 2.2625 / FAIL 2.275、幅0.0125、LU6707、Newton12973 | 見送り。V2系列との厳密比較ではない |
| V1+V2a+V4 | 3664.000 s | PASS 2.275 / FAIL 2.2875、LU6318、Newton13175 | 見送り。V1+V2aより567.646 s、18.33%遅い |

V4は局所的な候補採用があっても、追加の反復・LU・材料更新を含む総時間では利益を確認できなかった。V1+V4の最終出力は V1+V2a と完全一致せず、変位最大差 7.734542E-06 m、応力最大差 0.00607799、yielded差1/1224点だった。V2a+V4は精密 3680.353 s だが、前回V2系列はV2b有効・採用0であり、V4だけの差とは分離できない。

小モデルの ADAPT_02 分離試験では、V2a単独の長反復再利用により Case2: LU 14→12、Case3: 26→21、Case4: 13→11。変位成分最大差 2.603108256727E-09 m、全40実行の境界確認はPASS。ただし本番1005節点の速度・破壊近傍の依存性は未測定である。

## 3. adaptive / policy の履歴

`policy` のファイル名と内部番号を混同しないよう、表では両方を書く。

| 外部ファイル上の呼称 | 内部番号 | 変更の要点 | 総時間・結果 | 判断 |
|---|---|---|---:|---|
| `adaptive` | ADAPT_01 | V1/V2a/V4の自動選別。V5は未使用。 | 3099.930 s、保存 Fs 2.275、PASS 2.275 / FAIL 2.2875 | V1+V2aと同等。差3.576 sは単発では優劣不明。 |
| `adaptive_split` | ADAPT_02 | V1履歴とV2a長反復資格を分離。 | 3127.797 s | ADAPT_01比で採用理由なし。長反復再利用の実装確認版。 |
| `adaptive_policy` | ADAPT_03 | SHORT/LONG別休止、進捗だけの一律休止廃止、経路別ログ。 | 2929.445 s、Newton 11825、LU 5198、保存 Fs 2.275、PASS 2.275 / FAIL 2.2875 | 系列最速。今回モデルの運用候補。ただし各1回。 |
| `adaptive_policy2` | ADAPT_04 | LONG休止を16/32/64へ段階化、有限費用式と診断。 | 3124.238 s、Newton 12762、LU 5453 | ADAPT_03より194.793 s遅く、速度候補から見送り。診断版として保持。 |
| `adaptive_policy3` | ADAPT_05 | LONG休止を固定16へ戻した対照版。費用式・全増分ログを保持。 | 2973.844 s、Newton 11825、LU 5198。ADAPT_03と演算回数・探索結果一致 | 対照版として採用価値。測定時間はADAPT_03より44.399 s遅い。 |
| `adaptive_policy4` | ADAPT_06 | 小幅を1.1倍へ回復する増分幅制御。 | 3969.719 s、Newton 15033、LU 7162、回復85増分、基準復帰24 | 見送り。ADAPT_05より995.875 s増。最終塑性フラグ624→623で完全一致せず。 |
| `adaptive_policy5` | ADAPT_07 | 回復前後の費用窓を観測し、不良なら幅を戻す。累積塑性診断も追加。 | 3770.916 s、Newton14966、LU6638、cutback54、基準復帰20、最終塑性点623 | 見送り。ADAPT_03より841.471 s（28.72%）遅い。 |
| `practical` 最新既定 | BEST_03 | 実測最速ADAPT_03の選別・採算式を復元。追加回復OFF、実務修正・新しい診断を保持。 | 小モデル24条件・施工/載荷9条件PASS。今回の実務ブックのフル速度は未測定。 | ADAPT_03の実測を根拠に採用。2929.445 sの再測定値とは扱わない。 |

ADAPT_03の詳細は、ADAPT_02比で198.352 s短縮、LU398回減。最終出力は ADAPT_02 と出力精度内で一致し、塑性点624/1224。長経路は109回利用・不良25回、短経路は704回利用・不良8回。ADAPT_04は不良率を下げても Fs=2.2875 の仕事量が 996.836→1164.771 s と増え、総時間短縮にはならなかった。ADAPT_05は ADAPT_03 と全Fsの増分/Newton/LU/再利用数、最終CSV出力が一致したが、単発時間は44.399 s遅い。

## 4. 保持する収束判定と採用条件

各版で保持した判定は、非線形残差・補正比・拘束変位条件、材料更新成功、同じ目標への基準方式再試行、最小増分・再試行上限、SRMの数値 PASS/FAIL である。近限界の `numerical_status=CONVERGED` でも `mechanical_status=LIMIT_STATE` を併記する。失敗側の時間を早く打ち切ったことだけで高速化とは扱わない。

採用判断は、(1) 全体時間、Newton/LU回数、基準復帰・切戻し、(2) Fsごとの PASS/FAIL と最終幅、(3) 変位・応力・塑性状態、(4) 同条件の再測定、をそろえて行う。今回の既定には最速ADAPT_03を復元したBEST_03を採用した。V1/V2a/V4/適応選別ON、V2b/V3/V5/TRACE/追加回復/fresh LU試験OFFで、最速試行の候補設定を保持する。固定 V1+V2a は測定した固定組合せの最速である。V4の固定組合せ、ADAPT_04、ADAPT_06/07は今回の実測理由で見送り。

ADAPT_07の悪化は特にFs2.30 FAILEDで、ADAPT_05の771.398 sから1547.785 sへ増えた。近限界域の経路・仕事量の増加であり、全区間で一律に遅くなった結果ではない。最終Fs区間は同じでも、ADAPT_03との変位・応力出力は完全一致せず、yieldedは624→623。局所回復の採算判定だけでは、後続の経路を含む総時間の改善を保証できなかった。

## 5. 参照元・証拠ファイル

以下の評価資料と試行表を `docs/evidence` に同梱している。元の単独ZIPが欠けている版は、残っている評価で確認できる範囲を記載した。

- [高速化計画](evidence/FEM_VBA_only_acceleration_plan_20261004.md)
- [ベンチマーク条件](evidence/SRM_Case1_Case4_Benchmark_Conditions_20261004.md)
- [V1/V2a分離メモ](evidence/V1_V2a_split_notes_20261006.txt)
- [組合せ評価](evidence/combination_evaluation_20261006.txt)、[組合せ比較CSV](evidence/combination_trial_comparison_20261006.csv)
- [V1+V4評価](evidence/V1_V4_evaluation_20261006.txt)
- [adaptive評価](evidence/adaptive_full_evaluation_20261006.txt)、[split評価](evidence/adaptive_split_full_evaluation_20261006.txt)
- [ADAPT_03評価](evidence/adaptive_policy_full_evaluation_20261007.txt)、[ADAPT_03メモ](evidence/adaptive_policy_v3_notes_20261006.txt)
- [ADAPT_04評価](evidence/adaptive_policy4_full_evaluation_20261007.txt)、[ADAPT_04メモ](evidence/adaptive_policy_v4_notes_20261007.txt)
- [ADAPT_05評価](evidence/adaptive_policy5_full_evaluation_20261007.txt)、[ADAPT_05メモ](evidence/adaptive_policy_v5_notes_20261007.txt)
- [ADAPT_06評価](evidence/adaptive_policy6_evaluation_and_plan_20261008.txt)、[ADAPT_06メモ](evidence/adaptive_policy_v6_notes_20261007.txt)
- [ADAPT_07メモ](evidence/adaptive_policy_v7_notes_20261008.txt)
- [ADAPT_07の本番評価とBEST_03復元](evidence/adaptive_policy7_evaluation_and_restore_20261008.txt)、[受領した元ログ](evidence/2DSoilFEM_20261008_adaptive_policy5_out.zip)
- [UI整理メモ](evidence/UI_organized_notes_20261008.txt)

関連する試行CSV（値照合用）: [V0/V5比較](evidence/V0_V5_trial_comparison_20261005.csv)、[V1/V4比較](evidence/V1_V4_trial_comparison_20261006.csv)、[組合せ比較](evidence/combination_trial_comparison_20261006.csv)、[adaptive policy比較](evidence/adaptive_policy5_trial_comparison_20261007.csv)、[ADAPT_06比較](evidence/adaptive_policy6_trial_comparison_20261008.csv)。出典側にない精密値は本文・CSVへ補わず、測定済みでも元ログ欠落・詳細再確認不可、または単発測定と明記した。
