2DSoilFEM INCONSISTENT GMRESカーネル高速化版 20261020_INCO_GMRES_03

XLSMとDLL2個を同じフォルダに置き、64bit Windows Excelで開いてください。
保存済み結果は元ブックのものです。入力を確認して再解析してください。
Double・収束条件・材料則・SRM探索条件を維持しています。
GMRESのILU前処理・疎行列積を倍精度SIMDで計算します。
直交化2回・従来の真残差確認・反復上限・許容値を維持します。
加算の順序が変わるため途中のGMRES残差・再スタート回数には差が出ます。
NATIVE_GMRES_TUNED=0で従来のGMRES経路へ戻せます。キー欠落時も0です。
旧DLL・非対応環境では従来経路へ戻ります。材料DLLは前版と同一です。
検証条件はdocs/INCONSISTENT_GMRES_PERF.mdを参照してください。
