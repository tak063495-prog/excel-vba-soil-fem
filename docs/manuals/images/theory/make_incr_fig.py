# 増分・反復の図（2026-10-22）。解析例12 SRM記録 increment_summary.csv の工程1（GRAVITY）を描く。
import csv, matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
for f in fm.findSystemFonts():
    if 'NotoSansCJK' in f: fm.fontManager.addfont(f)
plt.rcParams['font.family']='Noto Sans CJK JP'
R=list(csv.DictReader(open('/workspace/soilfem/docs/manuals/examples/ex12_anchor_srm_logs_20261021/case_out/increment_summary.csv')))
fig,ax=plt.subplots(1,2,figsize=(10,3.6),dpi=150,sharey=True)
for a,tid,tt in ((ax[0],'9','Fs=1.70（収束）'),(ax[1],'10','Fs=1.80（非収束）')):
    rs=[r for r in R if r['trial_id']==tid and r['stage']=='1']
    for i,r in enumerate(rs,1):
        ok=r['accepted']=='True'; ls,lt=float(r['lambda_start']),float(r['lambda_target'])
        a.plot([i,i],[ls,lt],color='#2a7f62' if ok else '#c0392b',lw=3 if ok else 1.5)
        if not ok: a.plot(i,lt,'x',color='#c0392b')
    a.set_title(f'工程1 GRAVITY、{tt}：{len(rs)}回の試み、うち不採用{sum(r["accepted"]!="True" for r in rs)}回',fontsize=9.5)
    a.set_xlabel('増分の試み（順番）'); a.grid(alpha=.3); a.axhline(1,color='gray',ls='--',lw=.8)
ax[0].set_ylabel('荷重係数 λ（開始→目標）')
ax[1].plot([],[],color='#2a7f62',lw=3,label='採用（収束）'); ax[1].plot([],[],color='#c0392b',lw=1.5,marker='x',label='不採用（増分を縮めて再試行）'); ax[1].legend(fontsize=8,loc='upper left')
fig.tight_layout(); fig.savefig('th6_increments.png')
