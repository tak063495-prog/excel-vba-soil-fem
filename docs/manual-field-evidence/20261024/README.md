# 1024説明書・実務例追補の記録

[変更と検査](../../MANUAL_FIELD_EXAMPLES.md)／[集計・入力記録SHA256](field_examples.json)／[PDF構造とリンク](pdf_checks.json)／[目視確認](visual_checks.json)／[画像](visual_review.zip)／[配布同一性](package.json)

道路11工程・トンネルの通常施工9工程は公開した結果節点とステージ結果から再集計。ボックス2ケースは水理の合格記録、遮水体は独立検算の保存値を使用する。元データのパスとSHA256はfield_examples.jsonに記録。計算と時点を混在させず、通常施工の差分からSRM枝を除く。

再生成：`python -X utf8 tests/build_manual_field_examples.py`。描画・章参照検査は`tests/review_manual_practice.py`、`tests/check_manual24_pdf.py --ref INDEX`。PDF内のリンクはmainのパスへ向け、公開前はindex内のファイルと照合する。
