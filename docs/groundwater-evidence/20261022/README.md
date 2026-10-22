# 水条件追加版の検証根拠

配布 `2DSoilFEM_ui_20261022.xlsm`（`20261022_WATER_01`）そのものを使ったacceptance20の記録です。比較元はmainの`a444247`、Excel 16.0 build20430、倍精度・既存の材料則・収束/SRM条件を保持しています。

- [比較値](comparison.json)：水なし10比較の全170変位・Fs区間・残差がmainと完全一致。水ありDLL OFF/AUTO、空洞内水、施工・初期OFF/BIRTH・再開も照合。
- [独立検算](independent.json)／[分離水域](separated_independent.json)：弾性沈下、浮力、Q8の別実装による水頭・水圧・流束と水収支。
- [全記録ZIP](records.zip)：各fixtureの全変位・全応力/ひずみ、Excel出力表、perf_summary.csv、solver_events.csv、run.log、チェックポイントなど。ZIP内のフォルダー名がケース名です。
- [ブックの保存・コンパイル・原入力保持](package.json)、[埋込コード照合ログ](compile.log)、[図形回帰](figure.json)、[実Excel画面確認](ui_checks.json)。
- [ブック・DLL・入力・ソースのSHA256](provenance.json)。DLLは既存と同一の1個。
- [PDFの本文保持・章立て・リンク検査](manual_checks.json)／[全ページ概観と追加ページの確認画像](manual_review.zip)。原稿の第2～24章は、水条件への参照を加えたC9の説明以外を保持しています。

再現方法と数値の解釈は[検証記録](../../GROUNDWATER_VALIDATION.md)、利用手順と適用範囲は[操作説明書](../../manuals/2DSoilFEM_Manual_20261022_WATER_01.pdf)第25章を参照してください。実計算例の入力済みブックは[例題ZIP](../../manuals/examples/groundwater_examples.zip)です。

水圧の一方向連携であり、自由水面・不飽和・過渡圧密・fixed-stress連成を検証した記録ではありません。Fsは数値収束境界です。1回の実行時間は速度保証ではありません。専用コピーだけを使用し、Excelを時間制限で停止していません。
