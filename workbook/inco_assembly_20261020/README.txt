2DSoilFEM INCONSISTENT 行列組立高速化版 20261020_INCO_ASSEMBLY_01

XLSMとDLL2個を同じフォルダに置き、64bit Windows Excelで開いてください。
保存済み結果は元ブックのものです。入力を確認して再解析してください。
Double・収束条件・材料則・SRM探索条件を維持しています。
非ゼロ項からCSR散布表を作成する処理を計算DLLへ移しました。
要素・局所行・局所列の順序、加算順序、診断・ログを維持します。
NATIVE_ASSEMBLY=0で従来のVBA経路に戻せます。キー欠落時も0です。
旧DLL・非対応環境では従来経路へ戻ります。材料DLLは前版と同一です。
検証条件はdocs/INCONSISTENT_ASSEMBLY_PERF.mdを参照してください。
