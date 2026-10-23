# 2D Soil FEM 1024版

`2DSoilFEM_ui_20261024.xlsm` と `soilfem_native_x64.dll` を同じ通常フォルダーへ展開して使用します。DLLは1個です。1023版と同じDLLを使用しています。

通常の画面は、ホーム → モデル → 材料・接合 → 境界・水条件 → 施工工程 → 解析・保存 → 結果の7つです。各画面の「表示する項目」で入力表を切り替えます。青いセルを編集すると元の入力表へ即時反映します。

施工工程は行を選んで追加・複製・削除・上下移動・直前の取消しができます。水理計画は工程IDとともに移動します。最後の有効行のSRMはFs探索、途中のSRMは必要Fsの照査です。

形状や材料・水・工程等を変更した後は、解析前メッシュを確認してから解析します。既存メッシュを使う場合、メッシュを作り直す必要はありません。「結果」では最終保存状態の図と工程別の数値を確認できます。

詳細な操作・画面例：
https://github.com/tak063495-prog/excel-vba-soil-fem/blob/main/docs/STAGE_UI_20261024.md

操作・理論・解析例：1024版PDF説明書。
https://github.com/tak063495-prog/excel-vba-soil-fem/blob/main/docs/manuals/2DSoilFEM_Manual_20261024.pdf

コンパイル・保存再読込のソース一致、入力操作、数値回帰と実画面を確認しています。検証環境は64bit Excel 16.0 build 20430です。工程別の飽和fixed-stress時間連成に対応します。未知の自由水面・不飽和解析は対象外です。Double、材料則、収束条件、SRM探索条件は従来どおりです。

Q8/T6混在対応（初回導入時の計算識別 20261024_MIXED_01）：
「モデル」→「メッシュ生成・支持条件」→「不適合Q8をT6へ置換」で0=OFF／1=ONを指定します。初期値0です。品質修正後も目標未達のQ8だけをT6へ置換します。T6の輪郭は解析前メッシュで緑色、要素形式は詳細表K列です。辺の水理指定はT6の1・2・4を使います。再メッシュ後は支持・荷重・接合・水理境界を確認してください。
https://github.com/tak063495-prog/excel-vba-soil-fem/blob/main/docs/MESH_MIXED_T6.md

時間圧密・水位変動（計算識別 20261024_FIXED_STRESS_07）：
施工工程の「時間・連成設定」「時間の詳細設定」、材料・接合の「圧密材料特性」を設定します。結果は「時間履歴」「水量収支」「時刻別の図」で確認します。途中SRMは親の水圧を保持し、照査後は施工本線へ戻ります。入力手順・制限と検証記録：
https://github.com/tak063495-prog/excel-vba-soil-fem/blob/main/docs/FIXED_STRESS_STAGES.md
https://github.com/tak063495-prog/excel-vba-soil-fem/blob/main/docs/FIXED_STRESS_VALIDATION.md

水条件・SRM保存復元・RESET後の図表示・塑性結果更新の修正と検証記録：
https://github.com/tak063495-prog/excel-vba-soil-fem/blob/main/docs/XLSM_ISSUE_VALIDATION.md
計算識別06以前の保存途中データは元ブックで継続するか、07で入力から再解析してください。

07では保存再開の荷重・水位、時刻別の節点力・反力、頂点の材料点計算と塑性乗数、入力検査を修正しました。
https://github.com/tak063495-prog/excel-vba-soil-fem/blob/main/docs/ISSUE49_57_VALIDATION.md

中断の伝播とQ8内部の折返し検査を修正しました。受理済み状態を保持して中断し、全域のJacobian正値を確認できないQ8は要素番号付きで停止します。入力欄・保存形式・DLLは共通です。
https://github.com/tak063495-prog/excel-vba-soil-fem/blob/main/docs/ISSUE3_59_VALIDATION.md

内側の塑性履歴・連成・水量台帳・水圧CSVでも中断を伝播し、保存障害とユーザー中断を区別します。
https://github.com/tak063495-prog/excel-vba-soil-fem/blob/main/docs/ISSUE3_INNER_WRITERS.md


時間詳細行を省略したWATER工程の再開でも、通常実行と同じSTARTの水位境界・外水圧を復元します。明示HOLD／空欄モードのHOLD、LINEAR、時系列は保持します。計算識別・保存形式07、倍精度、収束条件は共通です。
https://github.com/tak063495-prog/excel-vba-soil-fem/blob/main/docs/ISSUE50_DEFAULT_RESUME.md
