# 理論解説図の作図スクリプト（2026-10-22）。旧説明書 図5・p40・p41 の説明図を描き直したもの。
import math, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
for f in fm.findSystemFonts():
    if 'NotoSansCJK' in f or 'NotoSansJP' in f: fm.fontManager.addfont(f)
plt.rcParams['font.family']='Noto Sans CJK JP'
C,PHI=10.0,30.0  # 解析例2・第4章の地盤（c=10 kPa、φ=30°）
# 1 強度低減
fig,ax=plt.subplots(figsize=(7.2,3.8),dpi=150); s=[0,120]
for F,col in ((1,'#1f3b5c'),(1.5,'#2a7f62'),(2.63,'#c0392b')):
    cr=C/F; pr=math.degrees(math.atan(math.tan(math.radians(PHI))/F))
    ax.plot(s,[cr+x*math.tan(math.radians(pr)) for x in s],color=col,lw=2,label=f'Fs={F}：c/Fs={cr:.2f} kPa、φr={pr:.1f}°')
ax.set_xlabel('法線応力 σn（kPa、圧縮正）'); ax.set_ylabel('せん断強度 τf（kPa）'); ax.grid(alpha=.3); ax.legend(loc='upper left',fontsize=9)
ax.set_title('τf = c/Fs + σn·tanφ/Fs（c=10 kPa、φ=30°）',fontsize=10); fig.tight_layout(); fig.savefig('th1_mc_srm.png')
# 2 流れ則
fig,(a,b)=plt.subplots(1,2,figsize=(8.6,3.8),dpi=150,gridspec_kw={'width_ratios':[1.25,1]})
P=30; t=math.tan(math.radians(P)); xs=[0,140]; a.plot(xs,[C+x*t for x in xs],color='#1f3b5c',lw=2,label='降伏面 f=0（φ=30°）')
p0=(80,C+80*t); a.plot(*p0,'ko')
L=30; n=(-math.sin(math.radians(P)),math.cos(math.radians(P)))
a.annotate('',xy=(p0[0]+L*n[0],p0[1]+L*n[1]),xytext=p0,arrowprops=dict(arrowstyle='->',color='#2a7f62',lw=2.2))
a.text(p0[0]+L*n[0]-30,p0[1]+L*n[1]+2,'関連流れ ψ=φ\n（面の法線方向）',color='#2a7f62',fontsize=9)
a.annotate('',xy=(p0[0],p0[1]+L),xytext=p0,arrowprops=dict(arrowstyle='->',color='#c0392b',lw=2.2))
a.text(p0[0]+3,p0[1]+L-4,'非関連流れ ψ=0\n（g=一定の法線）',color='#c0392b',fontsize=9)
a.axhline(p0[1],ls='--',color='#d68910',lw=1.2,label='塑性ポテンシャル g（ψ=0）')
a.set_xlim(0,140); a.set_ylim(0,110); a.set_xlabel('σn（kPa、圧縮正）'); a.set_ylabel('τ（kPa）'); a.grid(alpha=.3); a.legend(fontsize=8,loc='lower right')
a.set_title('(a) 降伏面と塑性ひずみの向き（模式図）',fontsize=10)
# (b) ダイレイタンシー：単純せん断で体積が増えるかどうか
for k,(psi,col,lab) in enumerate(((0,'#c0392b','ψ=0：体積変化なし'),(20,'#2a7f62','ψ=20°：せん断で膨らむ'))):
    y0=k*1.6; h=1.0; dx=0.5; dh=dx*math.tan(math.radians(psi))
    b.add_patch(plt.Polygon([(0,y0),(1.2,y0),(1.2,y0+h),(0,y0+h)],fill=False,ls=':',color='gray'))
    b.add_patch(plt.Polygon([(0,y0),(1.2,y0),(1.2+dx,y0+h+dh),(dx,y0+h+dh)],alpha=.35,color=col))
    b.text(1.85,y0+0.45,lab,fontsize=9,color=col)
b.set_xlim(-.1,3.6); b.set_ylim(-.2,3.0); b.axis('off'); b.set_title('(b) ダイレイタンシー角ψの意味（模式図）',fontsize=10)
fig.tight_layout(); fig.savefig('th2_flow.png')
# 3 SRM探索（解析例12の実記録 run.log）
tr=[(1.0,1),(1.1,1),(1.2,1),(1.3,1),(1.4,1),(1.5,1),(1.6,1),(1.7,1),(1.8,0),(1.75,0),(1.725,1),(1.7375,0)]
fig,ax=plt.subplots(figsize=(7.2,3.4),dpi=150)
for i,(f,ok) in enumerate(tr,1):
    ax.plot(i,f,'o' if ok else 'x',color='#2a7f62' if ok else '#c0392b',ms=8,mew=2)
ax.plot(range(1,13),[f for f,_ in tr],color='gray',lw=.8,zorder=0)
ax.axhspan(1.725,1.7375,color='#f5cba7',alpha=.6); ax.text(1.2,1.745,'Fs区間 1.725〜1.7375（中点1.731）',fontsize=9)
ax.plot([],[],'o',color='#2a7f62',label='収束（PASS）'); ax.plot([],[],'x',color='#c0392b',mew=2,label='非収束（FAIL）')
ax.set_xlabel('試行の順番'); ax.set_ylabel('試したFs'); ax.legend(fontsize=9,loc='lower right'); ax.grid(alpha=.3)
ax.set_title('0.1刻みで上げ、最初の非収束から二分法で区間を狭める（解析例12の実記録）',fontsize=10)
fig.tight_layout(); fig.savefig('th3_srm_search.png')
# 4 弾完全塑性
fig,ax=plt.subplots(figsize=(6,3.2),dpi=150)
ax.plot([0,1,3],[0,1,1],color='#1f3b5c',lw=2.2); ax.plot([1.8,2.6],[0,1],ls='--',color='gray')
ax.annotate('',xy=(1.8,0),xytext=(2.0,0.25),arrowprops=dict(arrowstyle='->',color='gray'))
ax.text(.15,.62,'弾性（E・ν）',fontsize=9); ax.text(1.3,1.06,'降伏後：強度一定（硬化・軟化なし）',fontsize=9)
ax.text(2.0,.35,'除荷は弾性\n（塑性ひずみが残る）',fontsize=8,color='gray')
ax.set_xticks([]); ax.set_yticks([]); ax.set_xlabel('ひずみ'); ax.set_ylabel('せん断応力'); ax.set_ylim(0,1.3)
ax.set_title('弾完全塑性（Mohr–Coulomb）の模式図',fontsize=10); fig.tight_layout(); fig.savefig('th4_elastoplastic.png')
