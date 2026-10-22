"""Regenerate groundwater manual figures from accepted Excel exports and screenshots."""
from pathlib import Path
import json, argparse
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Rectangle
from PIL import Image

ROOT = Path(__file__).resolve().parents[4]
ap = argparse.ArgumentParser()
ap.add_argument('--results', type=Path, default=ROOT/'tests/tmp/groundwater/acceptance20')
a = ap.parse_args()
OUT = Path(__file__).resolve().parent
evidence = ROOT/'docs/groundwater-evidence/20261022'
plt.rcParams.update({'font.family':'Noto Sans JP', 'font.size':10, 'axes.unicode_minus':False})
def read(p): return json.loads(p.read_text(encoding='utf-8-sig'))
for name, box in {'input':(0,220,1305,650), 'boundary':(0,220,830,900),
                  'preview':(0,220,1080,1090), 'result':(0,1090,1715,1310)}.items():
    Image.open(evidence/(name+'.png')).crop(box).save(OUT/(name+'.png'))

fixture=read(ROOT/'tests/fixtures/groundwater_cutoff.json')
nodes={r[0]:r[1:3] for r in fixture['nodes']}
rows=read(a.results/'cutoff/水圧結果.json')[35:]
gp=np.array([[r[3],r[4],r[5],r[12],r[13]] for r in rows if isinstance(r[0],(int,float)) and r[1] is not None],dtype=float)
fig,ax=plt.subplots(figsize=(8,6),layout='constrained')
for r in fixture['elements']:
    xy=np.array([nodes[n] for n in r[1:5]])
    ax.add_patch(Polygon(xy,fc='#e4e7e9' if r[9]==2 else '#fdf9ed',ec='#929a9c',lw=.7))
soil=gp[gp[:,2]>0]
sc=ax.scatter(soil[:,0],soil[:,1],c=soil[:,2],cmap='Blues',s=48,zorder=3,vmin=0)
ax.quiver(soil[:,0],soil[:,1],soil[:,3],soil[:,4],angles='xy',scale_units='xy',scale=6e-6,color='#bf5217',width=.004,zorder=4)
ax.plot([0,2],[5,5],'--',color='#147eac');ax.text(.1,5.12,'左：全水頭 5 m')
ax.plot([4,6],[4,4],'--',color='#147eac');ax.text(4.05,4.12,'右：全水頭 4 m')
ax.text(3,5.35,'遮水構造体\nK=0',ha='center')
ax.text(3,.4,'下端を回り込む浸透',ha='center',bbox={'fc':'white','alpha':.8,'ec':'none'})
ax.set(xlim=(-.2,6.2),ylim=(-.2,6.4),xlabel='X [m]',ylabel='標高 Y [m]',title='定常浸透の実計算：GP水圧とDarcy流束')
ax.set_aspect('equal');fig.colorbar(sc,ax=ax,label='GP水圧 [kPa]',shrink=.8)
fig.savefig(OUT/'flow.png',dpi=180);plt.close(fig)

boxrecord=read(a.results/'box_inside/result.json')
fig,(ax,bx)=plt.subplots(1,2,figsize=(9,4.7),layout='constrained',gridspec_kw={'width_ratios':[1.1,1]})
ax.add_patch(Rectangle((-.3,-.1),3.6,4.1,fc='#e5f4fd',ec='none'))
ax.add_patch(Rectangle((0,0),3,3,fc='#c8cdd1',ec='#515f69'))
ax.add_patch(Rectangle((1,1),1,1,fc='white',ec='#515f69'))
ax.add_patch(Rectangle((1,1),1,.5,fc='#84c7ef',ec='none'))
ax.plot([-.3,3.3],[4,4],'--',color='#147eac');ax.text(0,4.08,'外水位 4 m')
ax.text(1.5,1.7,'空洞',ha='center');ax.text(1.5,1.16,'内水位 1.5 m',ha='center',fontsize=9)
ax.text(1.5,2.6,'外形 3×3 m／空洞 1×1 m',ha='center',fontsize=9)
ax.annotate('外水圧合力 +90 kN',xy=(.45,.8),xytext=(.45,-.7),ha='center',arrowprops={'arrowstyle':'->','lw':2,'color':'#147eac'},fontsize=9)
ax.set(xlim=(-.5,3.5),ylim=(-.9,4.55),xlabel='X [m]',ylabel='標高 Y [m]',title='空洞を持つ固定支持ボックス')
ax.set_aspect('equal')
labels=['外水圧','内水荷重','自重','支持反力']
values=[90,-5,-boxrecord['metrics'][0],boxrecord['metrics'][2]]
colors=['#2685ab','#77b8dd','#777','#386f4e']
bx.bar(labels,values,color=colors);bx.axhline(0,color='#777',lw=.7)
for i,v in enumerate(values):bx.text(i,v+(5 if v>0 else -6),f'{v:+.0f}',ha='center',va='bottom' if v>0 else 'top')
bx.set(ylim=(-230,140),ylabel='鉛直力 [kN]（上向き正）',title='保存結果と独立の力の釣合い')
fig.suptitle('静水圧の実計算：γw=10 kN/m³、奥行1 m、構造体はTOTAL',fontsize=12)
fig.savefig(OUT/'box.png',dpi=180);plt.close(fig)
print('Generated input, boundary, preview, result, flow, box figures.')
