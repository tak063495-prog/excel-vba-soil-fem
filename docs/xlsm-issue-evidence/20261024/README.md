# XLSM issue修正の検証記録

[検証方法と結果](../../XLSM_ISSUE_VALIDATION.md)、[配布照合値と候補別範囲](validation.json)。

- 候補6：全時間連成17入力、issue再現17入力、表示5種類、従来の乾燥・水理・SRM回帰。ソースは `full_regression_source.zip`、ブックSHAは `full_regression_package.json`。
- 最終候補7：水理的に分離したJOINT面を許可するHEAD検査だけを追加修正。厳密コンパイル・再オープン52モジュール照合、境界・JOINT15入力、RESETの表示、塑性を含む時間連成SRM・保存再開・水理状態復元。
- 全記録ZIPには `archive_manifest.json` を付け、各ファイルのSHA256とZIPのCRCを確認した。DLLは重複収録せず配布ZIPの1個を使用する。保存結果ブックは元の例題名で収録し、同名の `_out` ディレクトリーと一緒に展開する。
- 画面は最終候補7の検証済み保存状態を専用コピーで開いて確認した。確認時の再計算は行っていない。

`focus7.log` の末尾には、CONTINUOUSの表の行数を独立DOF数と誤って同一視した初回検算の停止を残している。実出力は独立13DOFの水頭を別名へ展開した16節点行であり、共有面の水頭一致・反力の非重複を確認する検算へ訂正した。SOILの地盤・構造体例では、出力16行中の構造体側8行は未計算の水頭0・流量0である。[訂正後の再検算ログ](final_analytic_check.log)と[joint_analytic.json](joint_analytic.json)が合格記録。アプリの収束条件を変更したものではない。
