from pathlib import Path
import json,zipfile,hashlib,csv,io,sys
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
plt.rcParams.update({'font.family':'Yu Gothic','axes.unicode_minus':False,'font.size':10})
root=Path.cwd();out=Path(__file__).resolve().parent;out.mkdir(exist_ok=True)
base=root/'docs/inconsistent-evidence24';fixtures=json.loads((base/'fixtures.json').read_text(encoding='utf8'))['fixtures'];fi={f['name']:f for f in fixtures}
comp=json.loads((base/'practical_comparison.json').read_text(encoding='utf8'));audit=json.loads((base/'material_force_audit.json').read_text(encoding='utf8'));au={x['case']:x for x in audit['cases']}
z=zipfile.ZipFile(base/'srm_cases_INCO24.zip'); records=[]
labels={'griffiths_1999_coarse':'G φ20 粗','griffiths_1999_fine':'G φ20 細','griffiths_1999_phi0_coarse':'G φ0 粗','griffiths_1999_phi0_fine':'G φ0 細','griffiths_1999_foundation':'G 基礎地盤','pruska_h7_phi10':'P H7 φ10','pruska_h7_phi20':'P H7 φ20','pruska_h10.5_phi10':'P H10.5 φ10'}
order=list(labels)
p10,p20=fi['pruska_h7_phi10'],fi['pruska_h7_phi20']
assert p10['nodes']==p20['nodes'] and p10['elements']==p20['elements'] and p10['stages']==p20['stages']
assert all(a==b for i,(a,b) in enumerate(zip(p10['materials'][0],p20['materials'][0])) if i!=5)
def polygons(f,u=None):
    xy={int(n[0]):np.array(n[1:3]) for n in f['nodes']}
    return [np.array([xy[int(e[k+1])]+(u[int(e[k+1])] if u else 0) for k in [0,4,1,5,2,6,3,7]]) for e in f['elements']]
def mesh(ax,f):
    ax.add_collection(PolyCollection(polygons(f),facecolors='#f3f7f6',edgecolors='#4f807b',linewidths=.35));ax.autoscale_view();ax.set_aspect('equal');ax.set_xlabel('X [m]');ax.set_ylabel('Y [m]')
for name in order:
    f=fi[name];c=next(c for c in comp['cases'] if c['fixture']==name);a=au[c['case']]
    d=json.loads(z.read(f'cases/{name}/results/{c["case"]}.json'));m=dict(zip(d['metric_names'],d['metrics']))
    assert d['fixture_sha256'].lower()==hashlib.sha256((base/'fixtures.json').read_bytes()).hexdigest()
    assert m['analysis_ok'] and m['fos_bracket'] and c['flow_policy']=='INCONSISTENT'
    assert a['material_force_balance_pass'] and a['full_mc_admissibility_pass']
    assert abs(m['fos_pass']-a['fos_pass'])<1e-12 and abs(m['fos_fail']-a['fos_fail'])<1e-12
    mat=f['materials'][0]
    records.append(dict(fixture=name,label=labels[name],c_kPa=mat[6],phi_deg=mat[5],psi_deg=mat[7],E_kPa=mat[1],nu=mat[2],gamma_kN_m3=mat[4],nodes=len(f['nodes']),elements=len(f['elements']),fs_pass=m['fos_pass'],fs_fail=m['fos_fail'],material_residual=a['material_only_relative_residual'],max_vector_mm=a['max_vector_mm'],max_small_strain=a['max_absolute_small_strain'],seconds=c['elapsed_seconds'],build=c['source_build']))
fig,axs=plt.subplots(2,2,figsize=(10,6.4))
for ax,name in zip(axs.flat,[order[0],order[2],'pruska_h7_phi10','pruska_h10.5_phi10']):mesh(ax,fi[name]);ax.set_title(labels[name]+' / 参照形状')
fig.tight_layout();fig.savefig(out/'meshes.png',dpi=190);plt.close(fig)
fig,ax=plt.subplots(figsize=(9.6,4.4))
for i,r in enumerate(records):ax.plot([r['fs_pass'],r['fs_fail']],[i,i],lw=4,color='#277d79');ax.scatter([r['fs_pass'],r['fs_fail']],[i,i],marker='|',s=100,color='#244b70')
ax.set_yticks(range(len(records)),[r['label'] for r in records]);ax.invert_yaxis();ax.set_xlabel('表示Fsの数値区間（収束下端〜非収束上端）');ax.grid(axis='x',alpha=.25);fig.tight_layout();fig.savefig(out/'brackets.png',dpi=190);plt.close(fig)
fig,axs=plt.subplots(1,2,figsize=(10,3.8));sig=np.linspace(0,200,100)
for c,phi in [(10,20),(20,10),(20,20),(50,0)]:axs[0].plot(sig,c+sig*np.tan(np.deg2rad(phi)),label=f'c={c}kPa, φ={phi}°')
axs[0].set_xlabel('圧縮側の法線応力 σ [kPa]');axs[0].set_ylabel('せん断強度 τ [kPa]');axs[0].legend(fontsize=8);axs[0].grid(alpha=.2)
for i,name in enumerate(['pruska_h7_phi10','pruska_h7_phi20','pruska_h10.5_phi10']):
    r=next(r for r in records if r['fixture']==name);axs[1].plot([r['fs_pass'],r['fs_fail']],[i,i],lw=5);axs[1].scatter([r['fs_pass']],[i],color='#244b70')
axs[1].set_yticks(range(3),['H7 / φ10','H7 / φ20','H10.5 / φ10']);axs[1].invert_yaxis();axs[1].set_xlabel('実計算の表示Fs区間');axs[1].grid(axis='x',alpha=.2)
fig.tight_layout();fig.savefig(out/'strength_and_phi.png',dpi=190);plt.close(fig)
fig,ax=plt.subplots(figsize=(9,3.6));x=np.linspace(0,160,100);phi=np.deg2rad(20);px=95;py=20+px*np.tan(phi)
ax.plot(x,20+x*np.tan(phi),label='f=0：強度面（φ=20°）',color='#244b70');ax.plot(x,x*0+py,'--',label='g一定：非関連 ψ=0°',color='#c59438')
ax.scatter([px],[py],color='black')
ax.annotate('',xy=(px-12,py+33),xytext=(px,py),arrowprops=dict(arrowstyle='->',lw=2,color='#277d79'));ax.text(px-53,py+35,'関連：ψ=φ',color='#277d79')
ax.annotate('',xy=(px,py+33),xytext=(px,py),arrowprops=dict(arrowstyle='->',lw=2,color='#b84d47'));ax.text(px+4,py+27,'非関連：ψ=0',color='#b84d47')
ax.set_xlim(0,165);ax.set_ylim(0,110);ax.set_xlabel('法線応力（圧縮正）の概念座標');ax.set_ylabel('せん断応力の概念座標');ax.legend(loc='upper left',fontsize=8);ax.grid(alpha=.2);fig.tight_layout();fig.savefig(out/'flow_concept.png',dpi=190);plt.close(fig)
# Saved total-strain fields on two well-shaped last-converged states; no increment data is invented.
sys.path.insert(0,str(root/'tests'));from check_manual_case import N
fig,axs=plt.subplots(1,2,figsize=(10,4.3));geom=[]
for ax,name in zip(axs,['griffiths_1999_coarse','griffiths_1999_fine']):
    f=fi[name];r=next(r for r in records if r['fixture']==name);c=next(c for c in comp['cases'] if c['fixture']==name)
    d=json.loads(z.read(f'cases/{name}/results/{c["case"]}.json'));rr=np.array(d['displacements_flat']).reshape(-1,3);u={int(i):np.array([ux,uy]) for i,ux,uy in rr}
    assert abs(max(np.linalg.norm(v) for v in u.values())*1000-r['max_vector_mm'])<1e-8
    strain=list(csv.DictReader(io.StringIO(z.read(f'cases/{name}/analysis/{c["case"]}_strain.csv').decode('utf-8-sig'))));values={}
    for e in f['elements']:
        q=[x for x in strain if int(x['element'])==int(e[0])];assert len(q)==4
        values[int(e[0])]=max(np.hypot(float(x['ex'])-float(x['ey']),float(x['engineering_shear'])) for x in q)
    xy={int(n[0]):np.array(n[1:3]) for n in f['nodes']};mindet=1e30
    for e in f['elements']:
        pts=np.array([xy[int(i)]+u[int(i)] for i in e[1:9]])
        for aa in np.linspace(-1,1,11):
            for bb in np.linspace(-1,1,11):
                dr=N(aa+1e-30j,bb).imag/1e-30;ds=N(aa,bb+1e-30j).imag/1e-30;mindet=min(mindet,np.linalg.det(np.array([dr@pts,ds@pts])))
    assert mindet>0
    pc=PolyCollection(polygons(f,u),array=np.array([values[int(e[0])] for e in f['elements']]),cmap='YlOrRd',edgecolors='#666666',linewidths=.25);pc.set_clim(0,.012);ax.add_collection(pc);ax.autoscale_view();ax.set_aspect('equal');ax.set_title(f'{labels[name]} / Fs={r["fs_pass"]:.4f} / 倍率1');ax.set_xlabel('X [m]');ax.set_ylabel('Y [m]');fig.colorbar(pc,ax=ax,shrink=.65,label='面内最大せん断ひずみ（全変位由来）')
    geom.append(dict(fixture=name,min_deformed_jacobian_11x11=mindet,max_vector_mm=r['max_vector_mm']))
fig.tight_layout();fig.savefig(out/'total_shear.png',dpi=190);plt.close(fig)
summary={'scope':'Historical completed INCO24 runs, not new UI04 runs. All values from public archives.','cases':records,'display_geometry':geom,'source_files':[{ 'path':p.relative_to(root).as_posix(),'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in [base/'fixtures.json',base/'practical_comparison.json',base/'material_force_audit.json',base/'srm_cases_INCO24.zip']]}
(out/'summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
with (out/'cases.csv').open('w',encoding='utf8',newline='') as fp:
    writer=csv.DictWriter(fp,fieldnames=list(records[0]));writer.writeheader();writer.writerows(records)
print('PASS 8 existing cases; input/result/audit consistency; 2 plotted states positive dense Jacobian')
