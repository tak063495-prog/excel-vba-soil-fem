2DSoilFEM INCONSISTENT高速化版 20261020_INCO_NATIVE_01

XLSMとDLL 2個を同じフォルダに置き、64bit Excelで開いてください。
保存済みの結果は旧版の結果です。入力を確認し、再解析してください。
「高速化設定」の「DLLの演算方式（倍精度）」はAUTOが既定です。
AUTO: AVX2/FMA対応CPUでは同命令、それ以外はSSE2。SSE2も選択可能。
STRICT: 従来のVBA互換演算。GMRESは従来のx87 DLLを使用します。
Double・収束許容値・材料則・SRM探索条件を維持しています。
GMRESと状態復元を高速化。材料DLLとログ出力処理は従来版です。
微小な丸め差を許容します。すべてのモデルで同じ高速化率を保証しません。
検証条件・比較結果・再現手順は docs/INCONSISTENT_NATIVE_PERF.md を参照。
