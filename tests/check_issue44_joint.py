"""Analytic steady-flow checks for physically adjacent, hydraulically separate JOINT faces."""
import argparse,json,math
from pathlib import Path

def load(p): return json.loads(p.read_text(encoding='utf-8-sig'))
def main():
 p=argparse.ArgumentParser();p.add_argument('--root',required=True);p.add_argument('--out',required=True);a=p.parse_args();root=Path(a.root);report={}
 for name in ('issue44_joint_soil','issue44_joint_prescribed'):
  d=root/name;j=load(d/'result.json');resume=load(d/'resume.json')
  assert j['status'].startswith('PASS|') and resume['status'].startswith('PASS|'),name
  rows=load(d/'水圧結果.json')[35:]
  gp=[r for r in rows if isinstance(r[0],(int,float)) and r[1] in ((1,) if name=='issue44_joint_soil' else (1,2))]
  nodes=[r for r in rows if isinstance(r[31],(int,float))]
  faces=[r for r in rows if isinstance(r[40],(int,float))]
  exported_nodes=len(nodes)
  if name=='issue44_joint_soil':
   unused=[r for r in nodes if r[33]==0]
   assert len(unused)==8 and all(r[35]==0 and r[34] is False and r[38]>=1 for r in unused),name
   nodes=[r for r in nodes if r[33]!=0]
  ng=4 if name=='issue44_joint_soil' else 8;nn=8 if name=='issue44_joint_soil' else 16;nf=4 if name=='issue44_joint_soil' else 8
  assert len(gp)==ng and len(nodes)==nn and len(faces)==nf,(name,len(gp),len(nodes),len(faces))
  pe=qe=he=0.
  for r in gp:
   assert all(isinstance(r[k],(int,float)) and math.isfinite(r[k]) for k in (3,4,5,12,13)),name
   y=r[4];pe=max(pe,abs(r[5]-10*(3-1.5*y)));qe=max(qe,abs(r[12]),abs(r[13]-5e-6))
  for r in nodes:
   assert math.isfinite(r[33]) and math.isfinite(r[35]),name
   he=max(he,abs(r[33]-(3-.5*r[38])))
  # Width=2 m: each independent layer admits 1e-5 m3/s at its bottom,
  # discharges the same at its top; the coincident interface remains separate.
  re=max(abs(sum(r[35] for r in nodes if abs(r[38]-y)<1e-9)-expected) for y,expected in (((0,1e-5),(1,-1e-5)) if name=='issue44_joint_soil' else ((0,1e-5),(1,0),(2,-1e-5))))
  fe=max(abs(r[43]-(-1e-5 if r[42]==1 else 1e-5 if r[42]==3 else 0)) for r in faces)
  assert pe<1e-8 and qe<1e-12 and he<1e-9 and re<1e-12 and fe<1e-12,(name,pe,qe,he,re,fe)
  du=max(abs(x-y) for x,y in zip(j['disp'],resume['disp']))
  assert len(j['disp'])==len(resume['disp']) and du<=1e-9,name
  report[name]={'status':'PASS','gp_count':ng,'hydraulic_nodes':nn,'exported_nodes':exported_nodes,'faces':nf,'pressure_error':pe,'flux_error':qe,'head_error':he,'reaction_error':re,'face_error':fe,'resume_displacement_error':du,'metrics':j['metrics']}
 name='joint_water_continuous';d=root/name;j=load(d/'result.json');rows=load(d/'水圧結果.json')[35:]
 gp=[r for r in rows if isinstance(r[0],(int,float)) and r[1] in (1,2)];nodes=[r for r in rows if isinstance(r[31],(int,float))];faces=[r for r in rows if isinstance(r[40],(int,float))]
 # CONTINUOUS has 13 independent hydraulic DOFs, expanded to 16 mesh-node rows.
 aliases={}
 for r in nodes:
  xy=(r[37],r[38]);aliases.setdefault(xy,[]).append(r)
 assert len(gp)==8 and len(nodes)==16 and len(aliases)==13 and len(faces)==8,(name,len(gp),len(nodes),len(aliases),len(faces))
 assert all(max(r[33] for r in group)-min(r[33] for r in group)<1e-12 for group in aliases.values()),name
 pe=max(abs(r[5]-10*(3+.2*r[3]-r[4])) for r in gp);qe=max(max(abs(r[12]+2e-6),abs(r[13])) for r in gp)
 he=max(abs(r[33]-(3+.2*r[37])) for r in nodes)
 re=max(abs(sum(r[35] for r in nodes if abs(r[37]-x)<1e-9)-v) for x,v in ((0,-4e-6),(2,4e-6)))
 fe=max(abs(r[43]-(-2e-6 if r[42]==2 else 2e-6 if r[42]==4 else 0)) for r in faces)
 assert pe<1e-8 and qe<1e-12 and he<1e-9 and re<1e-12 and fe<1e-12,(name,pe,qe,he,re,fe)
 report[name]={'status':'PASS','gp_count':8,'hydraulic_nodes':13,'faces':8,'pressure_error':pe,'flux_error':qe,'head_error':he,'reaction_error':re,'face_error':fe}
 for name in ('issue44_joint_continuous','issue44_joint_conflict'):
  x=load(root/name/'input_check.json');assert x['status']=='PASS' and x['expected_error'] in x['actual_error'],name
  report[name]=x
 Path(a.out).write_bytes((json.dumps({'status':'PASS','cases':report},ensure_ascii=False,indent=2)+'\n').encode());print('PASS: JOINT hydraulic boundary checks')
if __name__=='__main__':main()
