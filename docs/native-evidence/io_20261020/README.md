# INCONSISTENT I/O比較の記録

対象は `20261020_INCO_IO_01`。変更前の基点は `040c84c4bc9f38e42520c817327b6a292f458cf0`。
本体の差分はFEMIoとFEMCoreのビルド番号だけで、DLL2個は前版の配布バイナリと同一です。
説明・再現手順・ログ障害時の制限は [INCONSISTENT_IO_PERF.md](../../INCONSISTENT_IO_PERF.md) を参照してください。

- `pair{1,2,3}_enabled{0,1}.json`: 同じソース・DLLでIO_FAST_PATHだけを切り替えた、計測コードなしの全探索6回。各170変位、丸め前Fs6項目、残差・反復回数、ログのSHA-256とバイト数を含みます。
- `benchmark_summary.json`: 3組すべての完全一致・速度改善を検査した集計。実行順は0→1、1→0、0→1です。
- `final_package.json` / `final_package_summary.json`: 再構築・設定保存・開き直しを確認した公開ブックの追加全探索と、基準との完全一致検査。
- `previous_release_check.json`: 前版の公開ブックとも全変位・残差・Fs・主要回数が一致。通常ログは先頭のビルド番号列を除いて完全一致です。同じビルドのペア比較では全バイト一致を必須にしています。
- `profile{0,1}.csv` / `profile{0,1}.json` / `profile{0,1}_summary.json`: 使い捨ての計測用ブックでの40手続きQPC内訳と全探索結果。親子を二重加算しない排他的時間の総和を検査しています。
- `profile_comparison.json` / `profile_injection.json`: 手続き別の呼び出し回数・秒数の比較と挿入対象。高速化率の主張には計測なしの3組を使用します。
- `io_{INCONSISTENT,DAVIS}_{0,1}.json` / `io_{INCONSISTENT,DAVIS}_summary.json`: Case1〜4×off/adaptiveの固定Fs比較。両流れ則を合わせて16組・32解析です。
- `io_builtin_{0,1}.json` / `io_builtin_summary.json`: 同梱306要素モデル、固定Fs=1の比較。
- `io_lock_retry.json`: 実Excel・Win32排他ロックによる24項目。保留、再試行、SaveAs、次回開始拒否、削除フォルダの再作成、ヘッダーと行順を確認。最終保留は0です。
- `environment.json` / `verification.json`: CPU・Excel・タイマー、ソース/ブック/DLL/ZIPのハッシュ、構築・静的検査・比較ゲートの単体試験。

JSONのパスは実行時の所在を示します。`tests/tmp` / `tests/results` と生CSVはGit管理外です。
ここには追試に必要な入力fixtureのハッシュ・全変位・丸め前Fs・回数・ログハッシュを保持しています。
過去の失敗した診断試験や、修正前の暫定時間を採用結果へ混ぜていません。
最終I/O試験では本体に診断用の書き換えを行っていません。
