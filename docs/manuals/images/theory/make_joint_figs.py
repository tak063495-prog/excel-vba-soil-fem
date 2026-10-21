# JOINT要素の解説図（2026-10-22）。FEMEngine.bas の接合要素ルーチン（fn=Kn·gap、τmax=c+max(−fn,0)tanφ）に合わせて描く。
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt, math
from matplotlib import font_manager as fm
for f in fm.findSystemFonts():
    if 'NotoSansCJK' in f: fm.fontManager.addfont(f)
plt.rcParams['font.family']='Noto Sans CJK JP'
fig,ax=plt.subplots(1,3,figsize=(10.5,3.4),dpi=150)
a=ax[0]
a.add_patch(plt.Rectangle((0,0.05),4,1,fc='#9fb7c9')); a.add_patch(plt.Rectangle((0,-1.05),4,1,fc='#c8b89a'))
a.plot([0,4],[0.05,0.05],'k',lw=1); a.plot([0,4],[-0.05,-0.05],'k',lw=1)
for x,t in ((0,'4'),(2,'6'),(4,'5')): a.plot(x,0.05,'o',color='#1f3b5c'); a.text(x,0.18,t,ha='center',fontsize=9,color='#1f3b5c')
for x,t in ((0,'1'),(2,'3'),(4,'2')): a.plot(x,-0.05,'s',color='#7a5c2e'); a.text(x,-0.33,t,ha='center',fontsize=9,color='#7a5c2e')
a.text(2,0.6,'B側（材料B）',ha='center',fontsize=9); a.text(2,-0.65,'A側（材料A）',ha='center',fontsize=9)
for x in (2-1.549,2,2+1.549): a.plot(x,0,'x',color='#c0392b')
a.text(4.1,0,'厚さ0\n×＝積分点3点',fontsize=8,va='center',color='#c0392b')
a.set_xlim(-.4,5.6); a.set_ylim(-1.3,1.3); a.axis('off'); a.set_title('(a) 接合要素の節点（6節点・厚さ0）',fontsize=10)
b=ax[1]; c,phi=5,20
for sn,col in ((0,'#c0392b'),(40,'#2a7f62'),(80,'#1f3b5c')):
    tm=c+sn*math.tan(math.radians(phi)); ks=100000; d=tm/ks
    b.plot([0,d*1e3,1.0],[0,tm,tm],color=col,lw=2,label=f'σn={sn} kPa → τmax={tm:.1f}')
b.set_xlabel('すべり量（mm）'); b.set_ylabel('せん断応力 τ（kPa）'); b.legend(fontsize=7.5); b.grid(alpha=.3)
b.set_title('(b) せん断：傾きKs、上限 c+max(σn,0)·tanφ',fontsize=10)
d=ax[2]
g=[-1,0,1]; d.plot([-1,0],[100,0],color='#1f3b5c',lw=2,label='密着（圧縮）σn=−Kn·gap')
d.plot([0,1],[0,-100],color='#2a7f62',lw=2,ls='--',label='引張許容M=1：開口しても引張を負担')
d.plot([0,1],[0,0],color='#c0392b',lw=2.5,label='M=0：開口で剥離（応力0）')
d.axhline(0,color='gray',lw=.6); d.axvline(0,color='gray',lw=.6)
d.set_xlabel('開口量 gap（＞0で開く）'); d.set_ylabel('法線応力 σn（圧縮正）'); d.set_xticks([]); d.set_yticks([]); d.legend(fontsize=7.5,loc='upper right')
d.set_title('(c) 法線：圧縮側と開口側',fontsize=10)
fig.tight_layout(); fig.savefig('th5_joint.png')
