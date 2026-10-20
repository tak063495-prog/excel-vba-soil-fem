from pathlib import Path
import json,gzip,argparse
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
plt.rcParams.update({'font.family':'Yu Gothic','axes.unicode_minus':False,'font.size':10})
ap=argparse.ArgumentParser();ap.add_argument('--out',type=Path,default=Path('tmp/wall-moment-review'));out=ap.parse_args().out;root=Path.cwd();d=json.loads((out/'summary.json').read_text(encoding='utf8'))
fig,axs=plt.subplots(1,2,figsize=(11,5),layout='constrained')
for ax,source,label in zip(axs,d['sources'],['粗：壁40 Q8','細：壁68 Q8']):
    t=json.loads(gzip.decompress((root/source['path']).read_bytes()))['tables'];n={r[0]:np.array(r[1:3]) for r in t['節点データ'][1:]}
    pp=[];cc=[]
    for e in t['要素データ'][1:]:
        pp.append([n[e[k+1]] for k in [0,4,1,5,2,6,3,7]]);cc.append({1:'#e6ded0',2:'#a4c5d7',3:'#c7d7a9'}[e[9]])
    ax.add_collection(PolyCollection(pp,facecolors=cc,edgecolors='#4e656b',linewidths=.45))
    ax.plot([18.6,19],[1.8,1.8],color='#c43f42',lw=2);ax.plot([18,18],[0,.4],color='#c43f42',lw=2)
    ax.annotate('縦壁 y=1.8',xy=(18.8,1.8),xytext=(16.5,2.5),arrowprops={'arrowstyle':'->'},color='#a52b30')
    ax.annotate('底盤 x=18.0',xy=(18,.2),xytext=(15.4,-.6),arrowprops={'arrowstyle':'->'},color='#a52b30')
    ax.plot(18.6,.4,'o',color='#d87e29',ms=5);ax.set(xlim=(14.5,19.8),ylim=(-.9,3.9),xlabel='X [m]',ylabel='Y [m]',title=label+' / 工程3・Fs=1 入力形状');ax.set_aspect('equal')
fig.savefig(out/'actual_meshes.png',dpi=180);plt.close(fig)
fig,axs=plt.subplots(1,2,figsize=(10,4.6),layout='constrained')
for case,label,col in [('wall_davis08','粗40 Q8','#44768b'),('wall_davis12','細68 Q8','#be6844')]:
    for ax,part in zip(axs,['stem','base']):
        rr=[r for r in d['records'] if r['case']==case and r['part']==part]
        ax.plot([r['station_m'] for r in rr],[r['Mz_kNm_per_m'] for r in rr],'o-',label=label,color=col,ms=4)
for ax,part in zip(axs,['縦壁：断面高さ y [m]','底盤：断面位置 x [m]']):
    ax.axhline(0,color='#888',lw=.7);ax.set(xlabel=part,ylabel='Mz [kN·m/m]');ax.grid(alpha=.2);ax.legend()
fig.suptitle('保存4GP応力からの後処理値：設計用収束値ではありません',fontsize=12)
fig.savefig(out/'moment_comparison.png',dpi=180);plt.close(fig)
fig,axs=plt.subplots(1,3,figsize=(10,3.8),layout='constrained')
for ax,div in zip(axs,[2,4,8]):
    # Schematic regular Q8 element borders; not an executed analysis mesh.
    for x in np.linspace(0,.4,div+1):ax.plot([x,x],[0,.8],color='#467382',lw=.8)
    for y in np.linspace(0,.8,2*div+1):ax.plot([0,.4],[y,y],color='#467382',lw=.8)
    ax.axhline(.4,color='#c43f42',lw=2);ax.set(xlim=(-.03,.43),ylim=(-.04,.84),title=f'厚さ方向 {div} 分割',xlabel='厚さ 0.4 m');ax.set_aspect('equal');ax.set_yticks([])
fig.suptitle('比較計画の模式図：実行済みメッシュ・推奨精度保証ではありません',fontsize=12)
fig.savefig(out/'mesh_plan_schematic.png',dpi=180);plt.close(fig)
print('PASS three figures')
