# Independent checks for a completed manual-case result.json.
from pathlib import Path
import json,math,collections,sys
import numpy as np
ROOT=Path('.')
def load(case):return json.loads((ROOT/case/'result.json').read_text(encoding='utf8'))
def N(r,s):return np.array([-(1-r)*(1-s)*(1+r+s)/4,-(1+r)*(1-s)*(1-r+s)/4,-(1+r)*(1+s)*(1-r-s)/4,-(1-r)*(1+s)*(1+r-s)/4,(1-r*r)*(1-s)/2,(1+r)*(1-s*s)/2,(1-r*r)*(1+s)/2,(1-r)*(1-s*s)/2])
def jac(xy):
 vals=[]
 for r in [-1/3**.5,1/3**.5]:
  for s in [-1/3**.5,1/3**.5]:
   dr=N(r+1e-30j,s).imag/1e-30;ds=N(r,s+1e-30j).imag/1e-30;J=np.array([dr@xy,ds@xy]);vals.append(np.linalg.det(J))
 return min(vals)
def state(d,stage=None):
 tables=d['tables'];el={int(e[0]):e for e in tables['要素データ'][1:]};ns={int(n[0]):n for n in tables['節点データ'][1:]}
 stage=stage or max(r[0] for r in tables['結果節点'][1:] if isinstance(r[0],(int,float)))
 u={int(r[1]):np.array(r[4:6],float) for r in tables['結果節点'][1:] if r[0]==stage}
 er={int(r[1]):r for r in tables['結果要素'][1:] if r[0]==stage and isinstance(r[4],(int,float))}
 inc=collections.defaultdict(list)
 for r in tables['結果増分'][1:]:
  if r[0]==stage:inc[int(r[1])].append(r)
 return el,ns,u,er,inc,int(stage)
def check(case,stage=None,strict_deformed=True):
 d=load(case);assert d['status']=='PASS';el,ns,u,er,inc,stage=state(d,stage)
 orig={int(n[0]):np.array(n[1:3],float) for n in d['tables']['節点データ'][1:]}
 referr=max(float(np.linalg.norm(np.array(r[2:4],float)-orig[int(r[1])])) for r in d['tables']['結果節点'][1:] if isinstance(r[0],(int,float)))
 assert referr<1e-10,(case,'saved/reference coordinate mismatch',referr)
 mindet0=min(jac(np.array([orig[int(n)] for n in e[1:9]])) for e in el.values())
 mindet1=min(jac(np.array([orig[int(n)]+u[int(n)] for n in el[eid][1:9]])) for eid in er)
 assert mindet0>0,(case,'reference',mindet0)
 assert all(np.isfinite(v).all() for v in u.values()),(case,'nonfinite displacement')
 if strict_deformed:assert mindet1>0,(case,'deformed',mindet1)
 grid=[(r,s) for r in np.linspace(-1,1,11) for s in np.linspace(-1,1,11)]
 dR=np.array([N(r+1e-30j,s).imag/1e-30 for r,s in grid]);dS=np.array([N(r,s+1e-30j).imag/1e-30 for r,s in grid]);dense={};bad_deformed=[]
 for mode,ids in [('reference',list(el)),('deformed',list(er))]:
  xy=np.array([[orig[int(n)]+(u[int(n)] if mode=='deformed' else 0) for n in el[e][1:9]] for e in ids])
  a=np.einsum('qn,end->eqd',dR,xy);b=np.einsum('qn,end->eqd',dS,xy);det=a[:,:,0]*b[:,:,1]-a[:,:,1]*b[:,:,0]
  dense[mode]=float(det.min())
  if mode=='deformed':bad_deformed=[ids[k] for k in np.flatnonzero(np.any(det<=0,axis=1))]
  if mode=='reference' or strict_deformed:assert dense[mode]>0,(case,mode,'11x11 Jacobian',dense[mode])
 fixed=max([abs(u[i][j]) for i,n in ns.items() for j in [0,1] if n[3+j]!=0],default=0);assert fixed<1e-10
 diag={r[0]:r[1] for r in d['tables']['診断'][1:] if r[0]}

 for eid,rows in inc.items():
  assert len(rows)==4 and sorted(int(r[2]) for r in rows)==[1,2,3,4],(case,eid,'Gauss points')
  for r in rows:
   ev=np.linalg.eigvalsh([[r[5],r[7]/2,0],[r[7]/2,r[6],0],[0,0,0]])
   assert abs(float(ev[-1]-ev[0])-r[12])<1e-10,(case,eid,'shear definition')
 maxu=max(float(np.linalg.norm(v)) for v in u.values());mxgamma=max((r[12] for rs in inc.values() for r in rs),default=0)
 sr=d['tables']['ステージ結果'];last=next(r for r in reversed(sr[1:]) if isinstance(r[0],(int,float)))
 return {'case':case,'status':d['status'],'stage':stage,'nodes':len(ns),'elements':len(el),'active_elements':len(er),'seconds':d['seconds'],'max_displacement_m':maxu,'relative_residual':diag.get('RelativeResidualFree'),'reaction_y_kN':last[7],'min_reference_jacobian':mindet0,'min_deformed_jacobian_scale1':mindet1,'dense_11x11_jacobian':dense,'deformed_geometry_valid':not bad_deformed,'invalid_deformed_elements':bad_deformed,'max_shear_increment':mxgamma,'fixed_disp_max':fixed}
def quadrature(xy):
 vals=[]
 for r in [-1/3**.5,1/3**.5]:
  for s in [-1/3**.5,1/3**.5]:
   dr=N(r+1e-30j,s).imag/1e-30;ds=N(r,s+1e-30j).imag/1e-30
   vals.append(float(np.linalg.det(np.array([dr@xy,ds@xy]))))
 return sum(vals)

def verify(case):
 d=load(case);tables=d['tables'];el,ns,_,_,_,_=state(d)
 area=collections.defaultdict(float)
 for e in el.values():area[int(e[9])]+=quadrature(np.array([ns[int(n)][1:3] for n in e[1:9]]))
 mats={int(r[0]):list(r) for r in tables['材料データ'][1:] if isinstance(r[0],(int,float))}
 active={i:str(m[9]).upper()=='ON' for i,m in mats.items()};records=[];types_by_stage={}
 outcomes={int(r[0]):r for r in tables['ステージ結果'][1:] if isinstance(r[0],(int,float))}
 for st in tables['ステージ'][1:]:
  if not isinstance(st[0],(int,float)) or not st[4] or int(st[0]) not in outcomes:continue
  i=int(st[0]);cmd=str(st[1]).upper();target=int(st[2]) if isinstance(st[2],(int,float)) else None
  if cmd=='BIRTH':active[target]=True
  elif cmd=='DEATH':active[target]=False
  elif cmd=='MATSET':
   prop=mats[int(st[3])];mats[target][1:9]=prop[1:9];mats[target][10:13]=prop[10:13]
  expected=sum(a*mats[m][3]*mats[m][4] for m,a in area.items() if active[m])
  types_by_stage[i]={m:str(row[10]).upper() for m,row in mats.items()}
  result=outcomes[i];error=float(result[7])-expected
  # The result-node Reaction columns contain residuals at free DOFs too.
  # A support-only weight gap is therefore NOT a global equilibrium error:
  # support_R + sum(free residual) must equal independently integrated weight.
  connected={int(n) for e in el.values() if active[int(e[9])] for n in e[1:9]}
  nr=[r for r in tables['結果節点'][1:] if r[0]==i]
  free_y=sum(float(r[9]) for r in nr if ns[int(r[1])][4]==0 and int(r[1]) in connected)
  free_x=sum(float(r[8]) for r in nr if ns[int(r[1])][3]==0 and int(r[1]) in connected)
  closure_y=error+free_y;closure_x=float(result[6])+free_x
  tol=max(1e-5,abs(expected)*1e-8)
  assert abs(closure_y)<tol,(case,i,'weight + free-residual closure',closure_y)
  assert abs(closure_x)<tol,(case,i,'horizontal + free-residual closure',closure_x)
  records.append({'stage':i,'type':cmd,'expected_weight_kN':expected,'reaction_y_kN':result[7],'difference_kN':error,'relative_support_weight_gap':abs(error)/max(abs(expected),1e-30),'free_residual_sum_y_kN':free_y,'global_equilibrium_closure_y_kN':closure_y,'global_equilibrium_closure_x_kN':closure_x,'support_only_roundoff_check':abs(error)<tol})
 meta=d.get('anchor');balance=None
 if meta:
  loads=[r for r in tables['載荷'][1:] if isinstance(r[0],(int,float))];fx=fy=moment=0.
  for r in loads:
   x,y=ns[int(r[1])][1:3];v=float(r[4]);axis=r[2]
   if axis=='X':fx+=v;moment-=y*v
   elif axis=='Y':fy+=v;moment+=x*v
  assert max(abs(fx),abs(fy),abs(moment))<1e-9
  assert abs(math.hypot(meta['force_x'],meta['force_y'])-50)<1e-10
  balance={'net_Fx_kN':fx,'net_Fy_kN':fy,'net_M_kNm':moment}
 # Validate plastic increment invariants against independently computed eigenvalues.
 maxerr=0.
 struct_increment=0.;struct_yield=[]
 for r in tables['結果増分'][1:]:
  if not isinstance(r[0],(int,float)):continue
  w=np.linalg.eigvalsh([[r[8],r[10]/2,0],[r[10]/2,r[9],0],[0,0,r[11]]])
  err=abs(float(w[-1]-w[0])-r[13]);maxerr=max(err,maxerr);assert err<1e-10
  if types_by_stage[int(r[0])][int(el[int(r[1])][9])]=='STRUCT':struct_increment=max(struct_increment,max(abs(float(v)) for v in r[8:12]))
 for r in tables['結果要素'][1:]:
  if isinstance(r[0],(int,float)) and isinstance(r[4],(int,float)) and types_by_stage[int(r[0])][int(el[int(r[1])][9])]=='STRUCT':
   struct_yield.extend(float(r[13+17*k]) for k in range(4))
 assert struct_increment<1e-12,(case,'STRUCT plastic increments',struct_increment)
 assert not struct_yield or max(struct_yield)<1e-6,(case,'STRUCT yielding',max(struct_yield))
 out={'case':case,'material_areas_m2':dict(area),'stage_balance':records,'anchor_balance':balance,'plastic_shear_eigen_max_error':maxerr,'struct_max_plastic_increment':struct_increment,'struct_max_yield_function':max(struct_yield) if struct_yield else None}
 (ROOT/case/'balance_checks.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf8');print(json.dumps(out,ensure_ascii=False,indent=2))
 return out

if __name__=='__main__':
 import argparse
 parser=argparse.ArgumentParser(description='Independent Q8 geometry, displacement, strain and force checks')
 parser.add_argument('result');parser.add_argument('--report-invalid-geometry',action='store_true',help='Write invalid deformed-shape diagnostics instead of stopping immediately; exit status remains 2')
 args=parser.parse_args()
 result=Path(args.result).resolve();assert result.name=='result.json'
 ROOT=result.parent.parent;case=result.parent.name
 rec=check(case,strict_deformed=not args.report_invalid_geometry);(result.parent/'independent_checks.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2),encoding='utf8')
 verify(case)
 if not rec['deformed_geometry_valid']:
  print('FAIL deformed geometry at scale 1; numerical convergence does not validate the deformed shape. See independent_checks.json. Other checks completed.');sys.exit(2)
 print('PASS geometry, displacement constraints, increment eigenvalues, global equilibrium including free residuals, and elastic structure checks')
