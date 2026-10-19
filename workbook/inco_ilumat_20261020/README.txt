2DSoilFEM INCONSISTENT ILU・材料評価高速化版 20261020_INCO_ILUMAT_01

XLSMとDLL2個を同じフォルダに置き、64bit Windows Excelで開いてください。
保存済み結果は元ブックのものです。入力を確認して再解析してください。
Double・収束条件・材料則・SRM探索条件を維持しています。
GMRESのILU対角逆数を一度準備して再利用します。材料評価の重複計算も削減しています。
直交化2回・従来の真残差確認・反復上限・許容値を維持します。
除算を逆数乗算へ変えるため、途中のGMRES残差・再スタート回数には丸め差が出ます。
NATIVE_GMRES_PREPARED=0で前版のGMRES経路へ戻せます。キー欠落時も0です。
旧DLL・非対応環境では従来経路へ戻ります。材料DLLはx87・MSVCR100演算を維持しています。
検証条件はdocs/INCONSISTENT_ILUMAT_PERF.mdを参照してください。
