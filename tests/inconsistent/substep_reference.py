"""Compose the independent six-facet oracle, then differentiate the complete map."""
import argparse,json
from pathlib import Path
import numpy as np
from spectral_reference import fixtures,oracle,elastic,rotation

def compose(case):
 current=dict(case,strain=(np.array(case['strain'])/case['steps']).tolist())
 ep=np.zeros(4)
 for _ in range(case['steps']):
  r=oracle(current)
  if not r['ok']:return r
  ep+=r['ep'];current['prev']=r['stress']
 return dict(ok=True,stress=r['stress'],ep=ep.tolist())

if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--generate',type=Path);p.add_argument('--analyze',type=Path);p.add_argument('--output',type=Path);args=p.parse_args()
 if args.generate:
  cases=[dict(c,name=c['name']+f'_steps{n}',steps=n) for c in fixtures() if c['name'].startswith(('elastic_r','mixed_r','shear_r','compression_r','face_psi0_z0_r','edge12_psi0_z0_r','edge23_psi0_z0_r')) for n in (4,8)]
  apex=next(c for c in fixtures() if c['name']=='apex_isochoric')
  cases += [dict(rotation(apex,t),name=f'apex_domain_r{t}_steps{n}',steps=n) for t in (0,30,60) for n in (4,8)]
  for n in (4,8):
   cases.append(dict(name=f'confinement_resolution_rejected_steps{n}',E=1000,nu=.3,c=1,phi=0,psi=0,prev=[1e8,1e8,0,1e8-5],strain=[0,0,0],steps=n,reject_tangent=True))
  args.generate.write_bytes(json.dumps(cases,indent=2).encode())
 if args.analyze:
  rows=[]
  for r in json.loads(args.analyze.read_text(encoding='utf-8-sig')):
   c=r['case'];v=r['v'];ref=compose(c)
   if c.get('reject_tangent'):
    assert ref['ok'] and not v[0] and not v[1],(c['name'],'unresolved wide secant accepted')
    rows.append(dict(name=c['name'],expected='TANGENT_PRECISION_REJECTION',accepted=False));continue
   expected_steps=1 if oracle(c).get('elastic',False) else c['steps']
   assert bool(v[0])==ref['ok'] and v[1] and v[2]==expected_steps,(c['name'],'status/step count',v[:3])
   stress=np.array(v[3:7]);ep=np.array(v[7:11]);T=np.array(v[11:20]).reshape(3,3)
   sg=float(np.max(abs(stress-ref['stress'])));pg=float(np.max(abs(ep-ref['ep'])))
   assert sg<1e-7 and pg<1e-9,(c['name'],sg,pg)
   _,C=elastic(c['E'],c['nu']);reconstruction=np.array(c['prev'])+C@(np.r_[c['strain'],0]-ep)
   assert np.max(abs(reconstruction-stress))<1e-7 and abs(ep[0]+ep[1]+ep[3])<1e-10
   J=np.zeros((3,3));h=1e-8
   for j in range(3):
    plus=dict(c,strain=list(c['strain']));minus=dict(c,strain=list(c['strain']))
    plus['strain'][j]+=h;minus['strain'][j]-=h
    rp,rm=compose(plus),compose(minus)
    if rp['ok'] and rm['ok']:J[:,j]=(np.array(rp['stress'][:3])-np.array(rm['stress'][:3]))/(2*h)
    elif rp['ok']:J[:,j]=(np.array(rp['stress'][:3])-np.array(ref['stress'][:3]))/h
    elif rm['ok']:J[:,j]=(np.array(ref['stress'][:3])-np.array(rm['stress'][:3]))/h
    else:raise AssertionError((c['name'],'no admissible derivative probe'))
   tg=float(np.linalg.norm(T-J)/max(1.,np.linalg.norm(J)))
   assert tg<1e-5,(c['name'],'composite tangent',tg)
   rows.append(dict(name=c['name'],preferred_steps=c['steps'],used_steps=v[2],stress_gap=sg,plastic_gap=pg,tangent_relative_gap=tg))
  args.output.write_bytes(json.dumps(dict(status='PASS',cases=len(rows),checks=rows),indent=2).encode());print('PASS',len(rows),'composite substep stress/flow/tangent checks')
