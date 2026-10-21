# 解析例2のメッシュ：NG例（旧、格子間隔2 m、84要素）と修正後（格子間隔1 m、330要素）の比較図（2026-10-22）
import math, matplotlib, openpyxl; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm, cm, colors
for f in fm.findSystemFonts():
    if 'NotoSansCJK' in f: fm.fontManager.addfont(f)
plt.rcParams['font.family']='Noto Sans CJK JP'
def old():
    wb=openpyxl.load_workbook('/workspace/soilfem/docs/manuals/examples/ex02_shape_20261021.xlsm',read_only=True,data_only=True)
    N={int(r[0]):(r[1],r[2]) for r in wb['節点データ'].iter_rows(min_row=2,values_only=True) if isinstance(r[0],(int,float))}
    E=[[int(v) for v in r[1:5]] for r in wb['要素データ'].iter_rows(min_row=2,values_only=True) if isinstance(r[0],(int,float))]
    return N,E
def new():
    rd=lambda f:[l.split('\t') for l in open(f,encoding='utf-8-sig').read().split('\n')[1:] if l.split('\t')[0].strip().isdigit()]
    N={int(c[0]):(float(c[1]),float(c[2])) for c in rd('x2q_nd.tsv')}
    E=[[int(v) for v in c[1:5]] for c in rd('x2q_el.tsv')]
    return N,E
def asp(P):
    L=[math.hypot(P[k][0]-P[k-1][0],P[k][1]-P[k-1][1]) for k in range(4)]; return max(L)/min(L)
nrm=colors.LogNorm(1,50); cmap=matplotlib.colormaps['RdYlGn_r']
fig,ax=plt.subplots(1,2,figsize=(10.5,3.8),dpi=150)
for a,(N,E),tt in ((ax[0],old(),'NG例：格子間隔2 m（84要素）'),(ax[1],new(),'修正後：格子間隔1 m（330要素）')):
    A=[]
    for e in E:
        P=[N[i] for i in e]; r=asp(P); A.append(r)
        a.add_patch(plt.Polygon(P,fc=cmap(nrm(r)),ec='k',lw=.3))
    a.set_aspect('equal'); a.set_xlim(-.3,14.3); a.set_ylim(-.3,8.5); a.set_title(f'{tt}\n最大の縦横比 {max(A):.1f}、5を超える要素 {sum(x>5 for x in A)}',fontsize=9.5)
    a.set_xlabel('X（m）')
ax[0].set_ylabel('Y（m）')
sm=cm.ScalarMappable(norm=nrm,cmap=cmap); cb=fig.colorbar(sm,ax=ax,shrink=.85); cb.set_label('要素の縦横比（最長辺／最短辺）')
fig.savefig('ex02_mesh_ng_ok.png',bbox_inches='tight')
# 解析対象図（修正後メッシュ）
fig,a=plt.subplots(figsize=(7.2,3.6),dpi=150); N,E=new()
for e in E: a.add_patch(plt.Polygon([N[i] for i in e],fc='#c8b89a',ec='#c8b89a',lw=.3))
for x0,y0,x1,y1 in ((0,6,6,6),): a.plot([x0,x1],[y0,y1],ls='--',color='gray')
a.text(3,6.3,'第4章の天端（Y=6 m）',fontsize=8,ha='center',color='gray'); a.text(7,4.2,'材料1 地盤\nc=10, φ=30°, γ=18',fontsize=8,ha='center',bbox=dict(fc='white',ec='none',alpha=.75))
a.annotate('',xy=(3,8),xytext=(3,6),arrowprops=dict(arrowstyle='->',color='#c0392b',lw=1.8)); a.text(3.2,7,'＋2 m',color='#c0392b',fontsize=9)
a.text(0.01,0.99,'底面：XY固定　左右：X固定\nGRAVITY → SRM',transform=a.transAxes,fontsize=7.5,va='top',bbox=dict(fc='white',ec='#999'))
a.set_aspect('equal'); a.set_xlim(-.5,14.5); a.set_ylim(-.5,9.5); a.set_xlabel('X（m）'); a.set_ylabel('Y（m）'); a.grid(alpha=.2)
a.set_title('解析例2の解析対象：天端を8 mに上げた斜面',fontsize=10); fig.tight_layout(); fig.savefig('t02.png')
