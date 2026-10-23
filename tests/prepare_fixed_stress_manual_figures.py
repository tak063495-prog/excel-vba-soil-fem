"""Reproduce manual figures from accepted fields and independent Mandel series."""
import argparse,csv,hashlib,json,sys,zipfile
from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT/'tests'))
from check_fixed_stress_mandel import pressure,constants,q8_point
ap=argparse.ArgumentParser(description=__doc__)
ap.add_argument('--records-root',type=Path,default=ROOT/'tests/tmp/manual_fs/reproduced_fields')
args=ap.parse_args()
OUT=ROOT/'docs/manuals/images/fixed_stress';E=ROOT/'docs/fixed-stress-evidence/20261024';R=args.records_root.resolve()
assert R.is_relative_to(ROOT), 'Keep reproduction files within this checkout.'
if not R.exists():
 with zipfile.ZipFile(ROOT/'docs/xlsm-issue-evidence/20261024/fixed_stress_records.zip') as archive:
  for member in archive.infolist():
   if member.filename.split('/')[0] not in ('fill_30day','wet_birth_ramp'):continue
   target=(R/member.filename).resolve()
   assert target.is_relative_to(R)
   archive.extract(member,R)
OUT.mkdir(exist_ok=True,parents=True);sources={};data={}
def load(p):
 sources[p.relative_to(ROOT).as_posix()]=hashlib.sha256(p.read_bytes()).hexdigest()
 return json.loads(p.read_text(encoding='utf-8-sig'))
def water(p):
 sources[p.relative_to(ROOT).as_posix()]=hashlib.sha256(p.read_bytes()).hexdigest()
 s=p.read_text(encoding='utf-8-sig').splitlines();i=next(i for i,x in enumerate(s) if x=='element,gauss,pressure_kPa')
 return {(int(r['element']),int(r['gauss'])):float(r['pressure_kPa']) for r in csv.DictReader(s[i:])}
def nodes(p):
 sources[p.relative_to(ROOT).as_posix()]=hashlib.sha256(p.read_bytes()).hexdigest()
 return {int(x['node']):{k:float(v) for k,v in x.items() if k!='node'} for x in csv.DictReader(p.open(encoding='utf-8-sig'))}
def save(fig,name):
 fig.tight_layout();fig.savefig(OUT/name,dpi=190,bbox_inches='tight');plt.close(fig)
mono=load(E/'final_monolithic.json');ms=mono['states'];fig,ax=plt.subplots(figsize=(7.4,3.8));ax2=ax.twinx()
ax.plot([r['time_s']/86400 for r in ms],[r['max_u_difference_m'] for r in ms],color='#1e5a7a');ax2.plot([r['time_s']/86400 for r in ms],[r['max_p_difference_kPa'] for r in ms],color='#b73d25')
ax.set(xlabel='Time [day]',ylabel='Maximum displacement difference [m]',title='Fixed-stress vs independent monolithic discrete solution');ax2.set_ylabel('Maximum pressure difference [kPa]',color='#b73d25');ax.grid(alpha=.2);save(fig,'monolithic_reference_error.png');data['column']=mono
fixture=load(ROOT/'tests/fixtures/fixed_stress/mandel_hybrid_refine.json');fine=load(E/'mandel30_final.json');coarse=load(E/'mandel23_final.json');mid=load(E/'mandel25_final.json');policy=load(E/'mandel.json')
xy={int(n[0]):(n[1],n[2]) for n in fixture['nodes']};coords={int(e[0]):[xy[int(n)] for n in e[1:9]] for e in fixture['elements']}
# Select the same nearest-symmetry GP used by the independent checker.
point=min([(e,g,*q8_point(v,g)) for e,v in coords.items() for g in range(4)],key=lambda z:(z[2],abs(z[3]-fixture['mandel']['b_m']/2)))
e,g,x,y=point;fr=[z for z in fine['frames'] if z['time_s']>0];theory=[pressure(x,z['time_s'],fixture) for z in fr]
assert all(abs(z['centre_excess_pressure_kPa']-v)<=z['max_pressure_error_kPa']+1e-9 for z,v in zip(fr,theory))
fig,ax=plt.subplots(figsize=(7.4,3.8));ax.plot([r['time_s'] for r in fr],theory,color='#212121',label='Independent Mandel series at the same GP');ax.plot([r['time_s'] for r in fr],[r['centre_excess_pressure_kPa'] for r in fr],linestyle='none',marker='.',ms=3,color='#1e688c',label='Accepted FEM pressure')
ax.scatter([0],[fine['frames'][0]['centre_excess_pressure_kPa']],color='#b73d25',s=20,label='t=0 discrete projection');ax.set(xlabel='Time [s]',ylabel='Excess pore pressure [kPa]',title=f'Mandel pressure near symmetry: x={x:.5f} m');ax.legend(fontsize=8);ax.grid(alpha=.2);save(fig,'mandel_central_pressure.png')
common=[]
for c in coarse['frames']:
 if c['time_s']<=0:continue
 f=min(fine['frames'],key=lambda r:abs(r['time_s']-c['time_s']));m=min(mid['frames'],key=lambda r:abs(r['time_s']-c['time_s']))
 if max(abs(f['time_s']-c['time_s']),abs(m['time_s']-c['time_s']))<1e-7:common.append((c,m,f))
assert len(common)==50 and policy['status']=='PASS' and all(policy['checks'].values())
fig,ax=plt.subplots(figsize=(7.4,3.8))
for i,label in enumerate(('Coarse space/time','Finer space/time','Refined early time step')):ax.semilogy([r[0]['time_s'] for r in common],[r[i]['max_pressure_error_kPa'] for r in common],label=label)
ax.set(xlabel='Time [s]',ylabel='Maximum GP pressure error [kPa]',title='Mandel: errors at the same 50 accepted times');ax.legend(fontsize=8);ax.grid(alpha=.2);save(fig,'mandel_pressure_error.png')
assert abs(max(r[2]['max_pressure_error_kPa'] for r in common)-policy['common_time_pressure_error_kPa']['early_max'])<1e-12
data['mandel']={'gp_element':e,'gp_number':g+1,'gp_xy_m':[x,y],'series_at_same_gp':theory,'computed_at_same_gp':[r['centre_excess_pressure_kPa'] for r in fr],'time_s':[r['time_s'] for r in fr],'common_times':len(common),'comparison':policy,'material_constants':constants(fixture)}
fig,axs=plt.subplots(2,1,figsize=(7.4,6.2));meshes,ma=plt.subplots(1,2,figsize=(7.4,4.5))
for ax,meshax,case in zip(axs,ma,('fill_30day','wet_birth_ramp')):
 f=load(ROOT/f'tests/fixtures/fixed_stress/{case}.json');root=R/case;outs=list(root.glob('*_out'));assert len(outs)==1,(case,outs);out=outs[0];run=load(root/'result.json');assert run['status'].startswith('PASS|') and run['fixture_sha256'].lower()==sources[f'tests/fixtures/fixed_stress/{case}.json'];hist=load(root/'時間履歴.json');base=water(out/'stg01_GRAVITY/water.csv');bn=nodes(out/'stg01_GRAVITY/node.csv')
 elems=[e for e in f['elements'] if int(e[9])==1];allowed={int(e[0]) for e in elems};ids={int(n) for e in elems for n in e[1:9]};coord={int(n[0]):(n[1],n[2]) for n in f['nodes']};topY=max(coord[n][1] for n in ids);topids=[n for n in ids if abs(coord[n][1]-topY)<1e-12];top=min(topids,key=lambda n:abs(coord[n][0]-.5*(min(coord[k][0] for k in ids)+max(coord[k][0] for k in ids))))
 assert abs(bn[top]['y']-coord[top][1])<1e-9
 result=[{'stage':1,'time_s':0,'settlement_mm':0,'max_dp_kPa':0,'min_dp_kPa':0}]
 for r in hist[1:]:
  if r[9]!='MAINLINE' or not r[10]:continue
  folder=root/str(r[10]).replace('\\','/');assert folder.resolve().is_relative_to(out.resolve());nn=nodes(folder/'node.csv');wp=water(folder/'water.csv');dp=[v-base[k] for k,v in wp.items() if k[0] in allowed]
  assert len(dp)==4*len(elems)
  result.append({'state':r[0],'stage':int(r[1]),'time_s':r[2],'settlement_mm':-(nn[top]['uy']-bn[top]['uy'])*1000,'max_dp_kPa':max(dp),'min_dp_kPa':min(dp)})
 times=[r['time_s']/86400 for r in result];ax2=ax.twinx();ax.plot(times,[r['settlement_mm'] for r in result],'-o',ms=3,color='#1e688c',label='Foundation-top additional settlement');ax2.plot(times,[r['max_dp_kPa'] for r in result],'--',color='#b73d25',label='Foundation max excess pressure');ax2.fill_between(times,[r['min_dp_kPa'] for r in result],[r['max_dp_kPa'] for r in result],color='#b73d25',alpha=.12)
 ax.set(xlabel='Cumulative construction time [day]',ylabel='Settlement, downward + [mm]',title=case);ax2.set_ylabel('Excess pressure vs initial [kPa]',color='#b73d25');ax.grid(alpha=.2);ax.legend(fontsize=7,loc='upper left');ax2.legend(fontsize=7,loc='lower right')
 assert all(result[i]['time_s']<=result[i+1]['time_s'] for i in range(len(result)-1));assert all(np.isfinite(v) for r in result for k,v in r.items() if isinstance(v,(int,float)))
 data[case]={'foundation_node':top,'foundation_node_xy_m':coord[top],'gp_elements':sorted(allowed),'states':result}
 for el in f['elements']:
  ids8=[int(n) for n in el[1:9]];xx=[coord[ids8[k]][0] for k in (0,4,1,5,2,6,3,7,0)];yy=[coord[ids8[k]][1] for k in (0,4,1,5,2,6,3,7,0)];meshax.fill(xx,yy,color='#e8d4ad' if el[9]==1 else '#b8ccaa',ec='#344f4f',lw=.6)
 meshax.set_aspect('equal');meshax.set(xlabel='X [m]',ylabel='Y [m]',title=case+' (reference mesh)')
save(fig,'practical_stage_history.png');save(meshes,'practical_mesh_schematic.png')
for name in ('monolithic_reference_error.png','mandel_central_pressure.png','mandel_pressure_error.png','practical_stage_history.png','practical_mesh_schematic.png'):
 from PIL import Image
 with Image.open(OUT/name) as im:im.verify()
(OUT/'figure_data.json').write_bytes((json.dumps({'status':'PASS','sources':sources,'data':data},ensure_ascii=False,indent=2)+'\n').encode());print('PASS: exact series, 50 common times, mainline physical histories, mesh connectivity')
