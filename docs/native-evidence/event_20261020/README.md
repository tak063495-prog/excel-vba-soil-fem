# ログ整形・バッファ追加の比較記録

基点は `687cdef51b94c0c07e39a3dcb1608934b1b8a25b`、採用候補は `20261020_INCO_EVENT_02`。
説明・再現・測定範囲は [INCONSISTENT_EVENT_PERF.md](../../INCONSISTENT_EVENT_PERF.md) を参照してください。

- `pair{1,2,3}_enabled{0,1}.json` / `benchmark_summary.json`: 計測コードなしの同じ公開用ブック・同じDLL・AUTOで、EVENT_LOG_FASTだけを0/1にした3組。順序は0→1、1→0、0→1。丸め前Fs6項目・全170変位・最終自由残差・主要回数・ネイティブの経路と回数・イベントCSVのSHA-256を保持します。
- `profile{0,1}.csv` / `profile{0,1}.json` / `profile{0,1}_summary.json` / `profile_comparison.json` / `profile_injection.json`: 使い捨ての計測ブック、47手続きQPC内訳。排他的時間の総和を検査します。高速化率の主張には上記の計測なしの結果を使用します。
- `discovery.*` / `discovery_summary.json` / `discovery_injection.json`: 基点687cdefの42手続き計測。FEMIoCsvとP6CsvHeaderEventを追加し、最初の費用を分解しました。
- `buffer_micro.json`: 65,536行×3回、128/512文字の独立バッファ比較。method=0は従来連結、1は予約領域へMid$で追記、2は行配列とJoin。QPC秒数と出力文字数を記録。全探索時間とは別です。
- `format_boundaries.json`: 2,200回の変化するコンテキスト・特殊文字、長い行、512行境界、ログOFF、1,100行の排他ロック中の蓄積と解除後の再試行。従来・高速双方の実CSV全バイトと行数・ヘッダーを比較します。
- `native_event_replay.json`: GMRESコード1〜6・未知コード・空配列件数・ログOFF・カーネル無効・入力配列不変・参照引数の前処理状態、および1,100行のロック中の再生と解除後の追記。通常の再生経路と実CSV全バイトを比較します。
- `io_retry_{fast,legacy}.json`: EVENT_LOG_FAST=1/0それぞれで、IO_FAST_PATH=0/1のロック・SaveAs・次回開始拒否・削除フォルダ再作成など24項目ずつ。
- `event_{INCONSISTENT,DAVIS}_{0,1}.json` / `event_{INCONSISTENT,DAVIS}_summary.json`: Case1〜4×off/adaptiveの固定Fs比較、両流れ則で16組・32解析。
- `event_builtin_{0,1}.json` / `event_builtin_summary.json`: 同梱306要素の固定Fs=1、全2,010変位の比較。
- `previous_release_check.json`: 前版公開ブックと全変位・丸め前Fs・残差・主要回数が一致。CSVはビルド番号列を除く109,847イベントが完全一致。同一ビルドの0/1比較は先頭列も含めて全バイト一致を必須にしています。
- `environment.json` / `verification.json`: CPU・Excel・タイマー・ソース、構築時の保存前後照合と設定の開き直し確認、DLL/ブック/ZIPのハッシュとCRC。
- `initial_attempt/`: 試作01の3組。数値とログは一致しましたが、全体時間の中央値が改善しなかったため、その結果を採用版の速度改善として扱っていません。02にはGMRESイベント列の共通準備を追加し、別の3組で再検証します。

JSONの絶対パスは実行時の所在です。使い捨てコピーや生CSVはGit管理外に保持します。
配布ブックには計測コードやテスト用コードを含めていません。
コンパイルに失敗した途中のブックは配布せず、修正後の保存前後のソース一致・コンパイルを確認したブックを使用します。
