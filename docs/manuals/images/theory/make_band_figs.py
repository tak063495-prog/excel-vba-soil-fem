# 7.8 せん断帯の見方：保存結果（INCO_24 Griffiths φ20）の全ひずみ図に、Bishop簡易法の臨界円を重ねる。
# 入力：docs/inconsistent-evidence24/fixtures.json, practical_comparison.json, srm_cases_INCO24.zip（読むだけ）
import json,zipfile,csv,io,glob
from pathlib import Path
import numpy as np, matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.collections import PolyCollection
for f in fm.findSystemFonts():
    if 'NotoSansCJK' in f or 'NotoSansJP' in f: fm.fontManager.addfont(f)
plt.rcParams['font.family']='Noto Sans CJK JP'
root=Path(__file__).resolve().parents[4]; base=root/'docs/inconsistent-evidence24'
fi={f['name']:f for f in json.loads((base/'fixtures.json').read_text('utf8'))['fixtures']}
comp=json.loads((base/'practical_comparison.json').read_text('utf8'))
z=zipfile.ZipFile(base/'srm_cases_INCO24.zip')
# --- Bishop simplified for the Griffiths slope (crest y=10 for x<=12, slope to (32,0), rigid base y=0)
c,phi,g=10.0,np.deg2rad(20),20.0
def ground(x): return np.where(x<=12,10.0,np.where(x>=32,0.0,10-(x-12)*0.5))
def bishop(xc,yc,R,n=60):
    # entry/exit where circle meets ground
    xs=np.linspace(xc-R,xc+R,4001); yb=yc-np.sqrt(np.maximum(R**2-(xs-xc)**2,0))
    inside=(yb<ground(xs))&(np.abs(xs-xc)<R)
    if inside.sum()<10: return None
    x1,x2=xs[inside][0],xs[inside][-1]
    if (yc-R)<-1e-6: return None  # below rigid base
    if x2>32.01 or x1<0: return None
    b=(x2-x1)/n; xm=x1+b*(np.arange(n)+.5); ybm=yc-np.sqrt(R**2-(xm-xc)**2); h=ground(xm)-ybm
    if (h<0).any(): return None
    W=g*h*b; al=np.arcsin((xc-xm)/R); F=1.5
    for _ in range(100):
        ma=np.cos(al)*(1+np.tan(al)*np.tan(phi)/F)
        Fn=np.sum((c*b+W*np.tan(phi))/ma)/np.sum(W*np.sin(al))
        if abs(Fn-F)<1e-7: break
        F=Fn
    return F,x1,x2
best=(9,None)
for xc in np.linspace(14,34,41):
    for yc in np.linspace(12,40,57):
        for R in np.linspace(5,45,81):
            r=bishop(xc,yc,R)
            if r and 0.5<r[0]<best[0]: best=(r[0],(xc,yc,R,r[1],r[2]))
Fle,(xc,yc,R,x1,x2)=best
print('Bishop FS=%.3f circle xc=%.2f yc=%.2f R=%.2f entry x=%.2f exit x=%.2f'%(Fle,xc,yc,R,x1,x2))
def field(name):
    f=fi[name]; cc=next(k for k in comp['cases'] if k['fixture']==name)
    d=json.loads(z.read(f'cases/{name}/results/{cc["case"]}.json')); m=dict(zip(d['metric_names'],d['metrics']))
    st=list(csv.DictReader(io.StringIO(z.read(f'cases/{name}/analysis/{cc["case"]}_strain.csv').decode('utf-8-sig'))))
    by={}
    for x in st: by.setdefault(int(x['element']),[]).append(np.hypot(float(x['ex'])-float(x['ey']),float(x['engineering_shear'])))
    xy={int(n[0]):np.array(n[1:3]) for n in f['nodes']}
    polys=[np.array([xy[int(e[k+1])] for k in [0,4,1,5,2,6,3,7]]) for e in f['elements']]
    vals=[max(by[int(e[0])]) for e in f['elements']]
    return polys,vals,m['fos_pass'],len(f['elements'])
fig,axs=plt.subplots(1,2,figsize=(11,3.6))
for ax,name,lab in zip(axs,['griffiths_1999_coarse','griffiths_1999_fine'],['(a) 粗いメッシュ','(b) 細かいメッシュ']):
    P,V,fs,ne=field(name)
    pc=PolyCollection(P,array=np.array(V),cmap='YlOrRd',edgecolors='#888',linewidths=.2); pc.set_clim(0,.012); ax.add_collection(pc)
    t=np.linspace(0,2*np.pi,721); cx,cy=xc+R*np.cos(t),yc+R*np.sin(t); ok=(cy<=ground(cx)+1e-9)&(cx>=x1-1e-6)&(cx<=x2+1e-6)&(cy<yc)
    ax.plot(cx[ok],cy[ok],'--',color='#1f4e9c',lw=1.6,label=f'Bishop簡易法の臨界円（Fs={Fle:.2f}）')
    ax.set_xlim(-0.5,33);ax.set_ylim(-0.5,11);ax.set_aspect('equal');ax.set_title(f'{lab}：{ne}要素、最後に成功したFs={fs:.4f}',fontsize=10)
    ax.set_xlabel('X [m]');ax.set_ylabel('Y [m]');ax.legend(loc='upper right',fontsize=8)
cb=fig.colorbar(pc,ax=axs,shrink=.8); cb.set_label('面内最大せん断ひずみ（全変位から、要素4点の最大）',fontsize=8)
fig.savefig(Path(__file__).with_name('th7_band_slope.png'),dpi=170,bbox_inches='tight')
