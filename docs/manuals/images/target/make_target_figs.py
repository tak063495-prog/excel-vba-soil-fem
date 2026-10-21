# 解析対象の模式図（2026-10-22）。各解析例の保存ブックの節点・要素・材料から描く。
import sys, gzip, json, math, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt, openpyxl
from matplotlib import font_manager as fm
for f in fm.findSystemFonts():
    if 'NotoSansCJK' in f: fm.fontManager.addfont(f)
plt.rcParams['font.family']='Noto Sans CJK JP'
EX='/workspace/soilfem/docs/manuals/examples/'
COL=['#c8b89a','#9fb7c9','#a9c99f','#e3c08d','#c9a0b8','#b0b0d8','#d9a3a3','#9fd0c9','#d8d09a']
def load(f):
    if f.endswith('.gz'):
        t=json.load(gzip.open(f,'rt',encoding='utf-8-sig'))['tables']; N,E,M=t['節点データ'],t['要素データ'],t['材料データ']
    else:
        wb=openpyxl.load_workbook(f,read_only=True,data_only=True); g=lambda s:[list(r) for r in wb[s].iter_rows(values_only=True)]
        N,E,M=g('節点データ'),g('要素データ'),g('材料データ')
    num=lambda v:isinstance(v,(int,float))
    xy={int(r[0]):(float(r[1]),float(r[2]),r[3],r[4]) for r in N[1:] if r and num(r[0])}
    kind={int(r[0]):(r[10],r[9],r[1]) for r in M[1:] if r and num(r[0])}
    el=[(int(r[0]),[int(v) for v in r[1:9] if num(v)],int(r[9])) for r in E[1:] if r and num(r[0]) and num(r[9])]
    return xy,el,kind
def draw(f,out,title,notes=(),arrows=(),labels=(),size=(7.2,3.6),off=(),mlab=None,lim=None):
    xy,el,kind=load(EX+f); fig,ax=plt.subplots(figsize=size,dpi=150)
    for i,ns,m in el:
        k=kind.get(m,('SOIL',))[0]
        if k=='JOINT':
            p=[xy[n] for n in ns[:2]]; ax.plot([p[0][0],p[1][0]],[p[0][1],p[1][1]],color='#c0392b',lw=3,zorder=5); continue
        P=[xy[n][:2] for n in ns[:4]]
        dashed = m in off or (kind.get(m,(0,'ON'))[1]=='OFF')
        ax.add_patch(plt.Polygon(P,fc=COL[(m-1)%len(COL)],ec=COL[(m-1)%len(COL)],lw=.3,alpha=.45 if dashed else .95,hatch='//' if dashed else None))
    X=[v[0] for v in xy.values()];Y=[v[1] for v in xy.values()]; xmin,xmax,ymin,ymax=min(X),max(X),min(Y),max(Y); s=(xmax-xmin)*0.012
    for n,(x,y,fx,fy) in xy.items():
        if fx==1 and fy==1 and abs(y-ymin)<1e-6 and n%3==0: ax.plot([x-s,x,x+s,x-s],[y-2*s,y,y-2*s,y-2*s],color='k',lw=.6)
        elif fx==1 and fy!=1 and (abs(x-xmin)<1e-6 or abs(x-xmax)<1e-6) and n%3==0:
            d=-1 if abs(x-xmin)<1e-6 else 1; ax.add_patch(plt.Circle((x+d*s,y),s*.8,fill=False,lw=.6))
    for (x,y,dx,dy,t) in arrows:
        ax.annotate('',xy=(x+dx,y+dy),xytext=(x,y),arrowprops=dict(arrowstyle='->',color='#c0392b',lw=1.8),zorder=6)
        if t: ax.text(x,y,t,color='#c0392b',fontsize=8,ha='center',va='bottom')
    if mlab is not None:
        from collections import defaultdict
        acc=defaultdict(lambda:[0,0,0])
        for i,ns,m in el:
            if kind.get(m,('SOIL',))[0]=='JOINT': continue
            for n in ns[:4]: a=acc[m]; a[0]+=xy[n][0]; a[1]+=xy[n][1]; a[2]+=1
        for m,(sx,sy,c) in acc.items():
            if m in mlab: labels=list(labels)+[(sx/c,sy/c,mlab[m])]
    for (x,y,t) in labels: ax.text(x,y,t,fontsize=8,ha='center',va='center',bbox=dict(fc='white',ec='none',alpha=.75,pad=1),zorder=7)
    ax.set_aspect('equal'); ax.set_xlim(xmin-(xmax-xmin)*.04,xmax+(xmax-xmin)*.04); ax.set_ylim(ymin-(ymax-ymin)*.12,ymax+(ymax-ymin)*.22)
    if lim: ax.set_xlim(lim[0],lim[1]); ax.set_ylim(lim[2],lim[3])
    ax.set_xlabel('X（m）'); ax.set_ylabel('Y（m）'); ax.set_title(title,fontsize=10); ax.grid(alpha=.2)
    if notes: ax.text(0.01,0.99,'\n'.join(notes),transform=ax.transAxes,fontsize=7.5,va='top',bbox=dict(fc='white',ec='#999',alpha=.9))
    fig.tight_layout(); fig.savefig(out); plt.close(fig)
if __name__=='__main__':
    sup='底面：XY固定（▲）　左右：X固定（○）'
    std={1:'材料1 地盤\nc=10, φ=30°, γ=18'}
    draw('ex01_material_20261021.xlsm','t01.png','解析例1・3・7の解析対象：高さ6 m・勾配1:0.67の斜面',(sup,'GRAVITY → SRM'),mlab={1:'材料1 地盤（解析例1は c=10→20）'})
    draw('ex03_support_20261021.xlsm','t03.png','解析例3の解析対象：同じ斜面で左面のX拘束を外す',('ケースA：左面X固定','ケースB：左面のX拘束を外す（底面以外）','GRAVITYのみ'),labels=[(0.6,3,'←ここの\n拘束を外す')],mlab=std)
    draw('ex04_staged_20261021.xlsm','t04.png','解析例4の解析対象：掘削（DEATH）と盛土（BIRTH）',(sup,'斜線＝初期OFF（工程3でBIRTH）','GRAVITY → DEATH 2 → BIRTH 3'),mlab={1:'材料1 下部地盤',2:'材料2 掘削する土',3:'材料3 盛土'})
    draw('ex05_load_20261021.xlsm','t05.png','解析例5・6の解析対象：幅12 m・高さ6 mの地盤と上面の載荷',(sup,'解析例5：X=4〜8 m の節点に各−10 kN','解析例6：X=5〜7 m を−0.02 mまで押し込む'),arrows=[(x,7.2,0,-1.1,'') for x in (4,5,6,7,8)],mlab=std)
    draw('ex08_road_result_20261021.xlsm','t08.png','解析例8の解析対象：道路盛土の拡幅（11工程）',(sup,'斜線＝初期OFF（BIRTHで施工）','材料8は掘削跡へMATSETで入れる再盛土'),mlab={1:'1',2:'2',3:'3',4:'4',5:'5',6:'6',7:'7',8:'8'})
    draw('ex09_wall_result_20261021.xlsm','t09.png','解析例9・12の解析対象：斜面途中のL型擁壁と背面盛土',(sup,'GRAVITY → BIRTH 2（擁壁）→ BIRTH 3（盛土）→ SRM','解析例12：頭部と定着部へ等大・反対向きの力'),arrows=[(18.6,2.4,-3.2,-0.92,'頭部'),(7.4,-0.8,3.2,0.92,'定着部')],mlab={1:'材料1 地盤',3:'材料3\n背面盛土'},labels=[(21.3,1.6,'材料2\nL型擁壁')],size=(7.6,4.2),lim=(0,30,-6,9))
    draw('ex10_tunnel_result_20261021.xlsm','t10.png','解析例10の解析対象：深さ24 mの円形トンネル（内空径8 m）',(sup,'上半→下半の順に掘削し支保（材料6）をBIRTH','SRMで施工完了後のFs'),mlab={},size=(6.4,4.6),labels=[(0,-17.5,'材料1〜5 地山（掘削範囲を番号で区分）'),(7,-24,'支保（材料6）\n厚0.3 m')],lim=(-14,14,-36,-10))
    draw('ex11_joint_20261021.xlsm','t11.png','解析例11・11-Bの解析対象：地盤上のブロックと接合面（赤線＝JOINT）',(sup,'11：自重のみ　11-B：ブロックへ水平力'),arrows=[(2.6,5.0,1.3,0,'水平力（11-B）')],mlab={1:'材料1 地盤',2:'材料2 ブロック'})
    draw('ex13_bearing_flat_s010_20261021.xlsm','t13f.png','解析例13の解析対象（水平地盤）：幅B=2 mの載荷帯',(sup,'自重のあと追加沈下 s を与える'),arrows=[(x,7.0,0,-1.6,'') for x in (1,2,3)],mlab={1:'粘性土 c=100 kPa, φ=0'})
    draw('ex13_bearing_slope_s010_20261021.xlsm','t13s.png','解析例13の解析対象（斜面上）：法肩から1 mの載荷帯',(sup,'高さ5 mの斜面'),arrows=[(x,7.0,0,-1.6,'') for x in (1,2,3)],mlab={1:'粘性土 c=100 kPa, φ=0'})
