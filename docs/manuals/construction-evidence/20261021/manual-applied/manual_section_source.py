import json
from reportlab.platypus import Paragraph,Table

def integrate(a):
    root=a['ROOT'];book=a['book'];P=a['P'];tab=a['tab'];img=a['img'];add=a['add']
    oldids={k:id(book[n-1]) for k,n in a['source_map'].items()}
    oldplan=json.loads((root/'tests/tmp/manual_applied/baseline_validation.json').read_text(encoding='utf8'))['page_plan']
    titles={p['page']:p['title'] for p in oldplan}
    def row(title):return next(r for r in book if r['title']==title)
    def insert(after,title,parts):
        ref=row(after);add(title,parts,ref['chapter']);new=book.pop();book.insert(book.index(ref)+1,new)
    insert(titles[16],'保存メッシュの拘束を変更する',[
        '形状を生成する場合と、保存された節点・要素表を使う場合では、拘束の入力先が違います。設定を変更しただけで既存節点の拘束が更新されたと扱わないでください。',
        tab([['ルート','操作と確認'],['形状から再生成','メッシュ設定C10:C14で支持条件を指定→形状メッシュを生成→節点データD/Eの実値と支持記号を確認。再生成で節点IDが変わるので載荷対象も選び直す。'],['保存メッシュを維持','形状生成は押さず、節点データB/Cの座標とA列IDで対象を特定。D列=X拘束、E列=Y拘束を0（自由）/1（拘束）へ直接変更する。'],['拘束と指定値','節点F/GのX/Y指定変位、載荷の同じDOFのFIX/DISPも確認。D/Eを0にする拘束解除と、F/GやDISPの値を0にする操作は別。DISP=0はゼロ変位拘束。'],['変更後','メッシュ確認を再表示→対象IDと支持記号を照合→EXPORT_LOADは空欄→初期状態から解析。旧結果と再開状態を流用しない。支持不足や剛体移動が生じないかはモデルとして別途判断。']],[1.15,4.05]),
        '例：保存メッシュの右側をX方向ローラーへ変えるなら、X最大の節点をB列で特定し、対象ID一覧を保存してD=1を指定します。底面との共有隅点では既存のY拘束を保持するなど、片方の列だけを意図どおり変えたか確認します。',
        '解析が読むのは節点D/Eの実値です。C10:C14は生成時の設定で、保存メッシュの節点を書き換えるボタンではありません。設定を変えてプレビューが無効化されても、拘束が更新された証拠にはなりません。',
        'この節は現行コードに基づく操作説明です。右側拘束変更の新しい計算例や、支持条件ごとの工学的妥当性を検証したものではありません。'])
    work=root/'tests/tmp/manual_applied/road_taller';d=json.loads((work/'exercise_summary.json').read_text(encoding='utf8'));b=d['baseline'];n=d['taller']
    insert(titles[55],'形状変更の実習：再盛土の高さを1m増やす',[
        img(work/'geometry_before_after.png',150),
        '道路のUI_05作業コピーを使います。材料8のγは19kN/m³のまま、形状入力の点17・18のY（D列）だけを4→5mへ変更。点ID・X座標・境界点列・材料・11工程は維持します。通常の形状生成ルートを使う実習です。',
        tab([['手順','再生成から照合まで'],['1 入力','元コピーを残す。点17=(5,5)、18=(7,5)mへ変更し、材料8のγ=19を確認。γ=20演習のコピーと混ぜない。'],['2 再生成','形状メッシュを作成→メッシュ確認。材料J列1～8、共有境界、底面/側面の支持、11工程を照合。本例はLOAD工程なし。新しいIDへ旧載荷IDを流用しない。'],['3 品質・計算','角度最適化後は99合格/4WARN/0不合格。WARNを消したとは扱わない。C14を空欄、INCONSISTENT、元の倍精度・収束条件で初期状態から解析。'],['4 結果','11工程PASS、Fs=1。最終反力はステージ結果H列=1899kN。元1842との差57kNを下記の独立計算と照合する。']],[1.05,4.15]),
        tab([['実測量','頂部Y=4m','頂部Y=5m'],['節点 / Q8',f"{b['nodes']} / {b['elements']}",f"{n['nodes']} / {n['elements']}"],['最大変位ノルム m',f"{b['max_displacement_m']:.8f}",f"{n['max_displacement_m']:.8f}"],['最終支持反力Y kN',f"{b['reaction_y_kN']:.3f}",f"{n['reaction_y_kN']:.3f}"],['解析秒（単発実測）',f"{b['seconds']:.6f}",f"{n['seconds']:.6f}"]],[2,1.6,1.6]),
        f"台形上部は上幅2m・下幅4m、高さ2→3mなので面積6→9m²。重量差=(9−6)×19×1=57kN。初期自重反力1440kNは不変。最終相対残差{d['last_relative_residual']:.3E}、全工程の最大鉛直釣合い差{d['max_global_closure_kN']:.2E}kN。11×11点の倍率1最小Jacobian={d['min_deformed_jacobian_11x11']:.6f}>0。",
        P('実時間は別時点の単発記録で高速化率にしません。これは施工変形の実習で、SRMやメッシュ収束確認ではありません。全節点・GP結果、ログ、変更入力と検算を公開資料へ保存しています。','small')])
    pages={r['title']:i+1 for i,r in enumerate(book)}
    def p(old):return pages[titles[old]]
    replacements={
      '今回の新規実計算は道路のγ変更演習です。':'新しい実習は道路のγ変更と形状変更です。',
      '現行原版の専用コピーで11工程を再計算。全節点・要素・増分・釣合いを確認。':'旧UI_04原版の専用コピーで11工程を計算。全節点・要素・増分・釣合いを確認。',
      '旧「設定」シートは互換用の参照式です。用途別設定のC列と二重入力しません。説明書の付録には現行83設定のシート・入力セル・保存値を一覧化しています。':'旧「設定」シートは互換用の参照式です。用途別設定のC列と二重入力しません。付録Aに現行入力欄と互換・不使用項目を区別して掲載しています。',
      '後続LOADをE=0にして':'ステージシートのLOAD行のE列を0（無効）にして',
      '入力変更後の再解析、結果の有効状態、選択ステージ、表示番号、倍率。':f'入力変更後の再解析、結果の有効状態、最後の有効工程と図見出しの工程番号、表示量、倍率。途中工程の表示はp.{p(51)}。',
      '工程1の鉛直支持反力1440kN、工程11は1842kN。Fs=1、工程11までPASS。数値と図の対象工程を対応させる。':f'ステージ結果D列=成否、E列=Fs/λ、H列=反力Y。工程1=1440kN、工程11=1842kN、Fs=1、全PASSを照合。集計の詳細はp.{p(55)}。',
      '結果を確認して.xlsmで保存→閉じる→同じファイルを開く→再解析せず結果図を再描画。ステージ結果の工程11とFs=1を照合する。':f'出力・表示C18で表示量を選び「変形・応力図を表示」（p.{p(84)}）。.xlsm保存→閉じる→再読込→再解析せず同じボタンで再描画。工程11/Fs=1を照合（p.{p(94)}）。',
    }
    def transform(x):
        if isinstance(x,Paragraph):
            s=x.text
            for old,new in replacements.items():s=s.replace(old,new)
            return Paragraph(s,x.style)
        if isinstance(x,Table):x._cellvalues=[[transform(y) for y in r] for r in x._cellvalues]
        return x
    for r in book:r['parts']=[transform(x) for x in r['parts']]
    for part in row(titles[112])['parts']:
        if isinstance(part,Table):
            part._cellvalues=[[Paragraph('既存実計算元',x.style) if isinstance(x,Paragraph) and x.text=='現行ブック' else x for x in rr] for rr in part._cellvalues]
    row(titles[112])['parts'].append(P('追加の道路形状変更は現行UI_05原版の専用コピーで新規に11工程を計算しました。点17/18のY変更、全節点・GP・工程結果と独立検算は公開資料manual-appliedに保存。既存擁壁等の再計算と混同しません。'))
    row(titles[10])['parts'].append(P('保存メッシュの拘束だけを変える場合はp.17。生成用設定と節点D/Eの実値の違いを確認します。','small'))
    row(titles[21])['parts'].append(P('C10:C14の支持設定は生成時に使います。保存メッシュを生成せず使う場合、既存節点D/Eの編集はp.17の手順で行います。','small'))
    row(titles[85])['parts'].append(P(f'USERは縮尺・原点の指定で、色上下限の指定ではありません。共通色範囲の扱いはp.{p(94)}を参照し、図ごとの凡例を併記します。'))
    source={56:'擁壁：slope_wall_shape.jsonは形状生成、旧wall_davis08 ZIPは履歴UI_04。新規計算はUI_05へ入力移行。',69:'アンカー：slope_wall_mesh.jsonは保存メッシュ。旧anchor_inco ZIPは履歴UI_04。等価力の対象IDも確認。',72:'トンネル：tunnel_mesh.jsonは保存メッシュ。旧tunnel10 ZIPは履歴UI_04。再現・物性・工程変更が範囲で、任意径の作成は別のメッシュ作成技能が必要。',78:'支持力：bearing_flat_mesh.json / bearing_slope_mesh.jsonは保存メッシュ。通常の形状生成は使わない。寸法変更には別途Q8メッシュの作成が必要。'}
    for old,text in source.items():row(titles[old])['parts'].insert(0,P(text+f' 取得・UI_05への移行はp.{p(10)}–{p(11)}。','small'))
    target=row(titles[61]);target['parts'].append(P('この断面力はExcelの表示ボタンによる自動出力ではなく、付属Python/NumPyで既存2ケースの保存応力を後処理したものです。一般案件用の自動処理ツールではありません。'))
    url='https://github.com/tak063495-prog/excel-vba-soil-fem/blob/main/docs/manuals/construction-evidence/20261021/wall-moment-guide/README.md'
    target['parts'].append(Paragraph(f'<link href="{url}" color="#236c91"><u>既存例の再現README・実行コマンド</u></link>',a['sty']['body']))
    target['parts'].append(tab([['別案件で確認・置換するもの','本例の固定値'],['入力と対象','wall_davis08/12のresult.json.gz、材料2の壁要素、工程3 BIRTH・Fs=1。対象案件の同じ保存状態へ置換。'],['断面・厚さ・奥行','縦壁x=18.6～19.0、底盤y=0～0.4m、奥行1m。対象の切断位置・法線・単位へ置換。'],['自重・切出し','γ=24kN/m³、縦壁上部/底盤左部の固定領域。対象の材料値・重量・領域を定義し、粗密・局所釣合いを再検査。']],[2.1,3.1]))
    for r in book:r['parts']=[P(x) if isinstance(x,str) else x for x in r['parts']]
    for k,ident in oldids.items():a['source_map'][k]=next(i for i,r in enumerate(book,1) if id(r)==ident)
