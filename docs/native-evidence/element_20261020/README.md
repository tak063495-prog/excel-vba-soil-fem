# 要素行列積・ILU構築 検証記録

前版6292076との比較。final_pairs01の全5組は同一のコンパイル済みVBA・同一native DLL・同じ設定で、材料DLLだけを交換。
集計はfinal_summary_all5.json。final_summary01.jsonは追加確認前の3組で、全5組へ含めています。
final_logs_all5.jsonは全CSVを独立に再読し、GMRES行を除外せず全バイト比較した結果です。
solver_events_pair1.zipには両側の実CSVを保存。各JSONの元ローカルパスは再現時には利用者の出力先へ読み替えてください。

O2試作版(unroll_pairs01)とILU単独の単発(ilu_pilot)は全体改善を確認できなかった記録で、O3配布版の5組と別集計です。
profile0c/profile1o3は使い捨て計測DLLの記録で、配布版の全体時間の根拠へ混ぜていません。
独立ILUと残差融合の単体試験もExcel全体の短縮率ではありません。残差融合は製品へ不採用です。

point/abiのDLLは制御語変更等のテスト用エクスポートを含む使い捨てDLLです。配布DLLには含まれません。
各build.jsonでソース・DLL・コンパイル設定を記録。テスト専用ソースはtest_material_points.py / test_element_exact.pyの--prepareで生成できます。
材料点20万件、要素5,200件、ILU両制御語2,024件、材料GMRES480件、配布用ILU1,012件、モデル17組、正則化回復112照合を保存。
再現コマンドと数値結果は[本文](../../INCONSISTENT_ELEMENT_PERF.md)を参照してください。
