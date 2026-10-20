# 出力設定・実画面の根拠

12枚は計算済み擁壁例の専用コピーを実際のExcelで開いて撮影した未加工のウィンドウ画像です。PDFではセルと操作箇所が読める範囲を切り出して配置しています。表示値は例の保存条件であり、全案件の推奨値ではありません。

`capture_manifest.json` に元コピーと画像のSHA-256、撮影操作を記録しています。撮影コピーは保存せず閉じ、元XLSM・VBA・DLLは変更していません。原版計算の再実行は不要な説明書改訂です。

リポジトリのルートから `python docs/manuals/construction-evidence/20261021/output-guide/rebuild_output_guide.py` を実行すると、公開済み圧縮記録から2組の集計CSVと図を再作成します（Python、NumPy、Matplotlib、日本語フォント）。道路は工程1/4/6/11の全節点変位ノルムと反力、SRMはsrm_stage=4のtrial_idで照合した12試行です。非収束試行のumaxは途中値であり、収束曲線へ結びません。CSVのumaxは自由DOF変位の最大絶対成分で、道路の変位ノルムとは別量です。
