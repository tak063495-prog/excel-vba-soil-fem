"""Plot actual Excel force integrals; geometry is the undeformed verification mesh."""
from pathlib import Path
import argparse,json
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

def load(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def integral(folder):
    rows=[r for r in load(Path(folder)/'接合結果.json')[1:] if isinstance(r[0],(float,int))]
    stage=max(r[0] for r in rows);rows=sorted([r for r in rows if r[0]==stage],key=lambda r:r[3])
    return [sum(w*r[c] for w,r in zip([5/9,8/9,5/9],rows)) for c in [17,15]]

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--results',required=True);p.add_argument('--out',required=True);a=p.parse_args()
    plt.rcParams.update({'font.family':['Yu Gothic','Meiryo','DejaVu Sans'],'font.size':10})
    root=Path(a.results);dry=integral(root/'dry');wet=integral(root/'wet');reaction=load(root/'wet/result.json')['metrics'][2]
    fig,(ax,b)=plt.subplots(1,2,figsize=(11,4.6),gridspec_kw={'width_ratios':[1.35,1]},layout='constrained')
    ax.add_patch(Rectangle((0,0),2,1,facecolor='#e9ddbf',edgecolor='#84754e',lw=1.5))
    ax.add_patch(Rectangle((0,1),2,1,facecolor='#dce3e7',edgecolor='#445a67',lw=1.5))
    ax.plot([0,2],[1,1],color='#804a9c',lw=3);ax.scatter([0,1,2],[1]*3,c='#804a9c',s=25,zorder=4)
    ax.text(.65,.45,'基盤 SOIL\n自重40 kN',ha='center',va='center')
    ax.text(1.,1.76,'STRUCT',ha='center',va='center')
    ax.plot([-.15,2.2],[2,2],ls='--',color='#2681b8');ax.text(-.15,2.15,'水位線 標高2 m',color='#206a96')
    ax.text(2.1,.98,'JOINT\n標高1 m\np=10 kPa',color='#704589',va='center')
    def arrow(x,y0,y1,color):ax.annotate('',xy=(x,y1),xytext=(x,y0),arrowprops={'arrowstyle':'-|>','color':color,'lw':2,'mutation_scale':16})
    arrow(.55,1.02,1.4,'#428357');ax.text(-.35,1.43,f'有効接触力\n{wet[0]:.0f} kN',color='#326442',ha='center')
    arrow(1.5,1.02,1.4,'#2681b8');ax.text(1.55,1.43,f'水圧支持力\n{wet[1]:.0f} kN',color='#206a96',ha='center')
    arrow(1.55,2.5,2.05,'#b35041');ax.text(1.6,2.45,'重量48 kN',color='#a3473a',va='center')
    arrow(1,-.45,-.02,'#374d48');ax.text(1.18,-.35,f'底面反力 {reaction:.0f} kN',color='#374d48')
    ax.set(xlim=(-.65,3.15),ylim=(-.55,2.7),aspect='equal');ax.axis('off');ax.set_title('接合面の力の釣合い（奥行1 m）',loc='left',fontweight='bold')
    y=[1,0];b.barh(y,[dry[0],wet[0]],color='#428357',label='固体の有効接触力',height=.42)
    b.barh(y,[dry[1],wet[1]],left=[dry[0],wet[0]],color='#2681b8',label='水圧による支持力',height=.42)
    for yy,(solid,water) in zip(y,[dry,wet]):
        b.text(solid/2,yy,f'{solid:.0f}',color='white',ha='center',va='center',fontweight='bold')
        if water:b.text(solid+water/2,yy,f'{water:.0f}',color='white',ha='center',va='center',fontweight='bold')
        b.text(solid+water+.8,yy,f'計{solid+water:.0f}',va='center')
    b.set(yticks=y,yticklabels=['乾燥','接合面p=10 kPa'],xlim=(0,58),ylim=(-.55,1.55),xlabel='支持力 [kN]')
    b.set_title('実計算の合力を比較',loc='left',fontweight='bold');b.legend(loc='upper center',bbox_to_anchor=(.5,-.17),frameon=False)
    b.spines[['top','right']].set_visible(False);b.grid(axis='x',alpha=.2);b.set_axisbelow(True)
    fig.suptitle('水圧を接触ばねから再減算せず、別の外力として作用させる',fontweight='bold')
    Path(a.out).parent.mkdir(parents=True,exist_ok=True);fig.savefig(a.out,dpi=180,facecolor='white');plt.close(fig)
