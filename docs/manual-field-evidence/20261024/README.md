# 1024説明書・実務例と工程別出力の確認記録

[変更と確認](../../MANUAL_FIELD_EXAMPLES.md)／[監視点・有効接続・工程別数値](field_examples.json)／[PDF構造・リンク](pdf_checks.json)／[目視確認](visual_checks.json)／[画像](visual_review.zip)／[配布同一性](package.json)／[実Excelの操作経路](../../validation/issue35_stage_routes/README.md)

道路11工程・トンネル通常9工程の掲載図は公開済みの保存計算から再集計する。初期有効材料とBIRTH/DEATH/MATSETを追跡して有効要素IDを結果要素と照合し、有効要素に接続しない期間を欠測として扱う。道路の恒久監視点は節点19、トンネル高さ方向の基準は工程1/5/9、幅・地表は工程1。出典SHA256はfield_examples.jsonに記録。

再生成：`python -X utf8 tests/build_manual_field_examples.py`。PDF検査：`tests/review_manual_practice.py`、`tests/check_manual24_pdf.py --ref INDEX`。最新版XLSMでのALL保存・元表への移動・抽出は別の実Excel記録へ分離する。過去版と最新版の計算値を同一の実測として扱わない。
