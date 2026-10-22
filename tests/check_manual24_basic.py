"""Independent checks for the 1024 manual's basic slope, without Excel."""
from pathlib import Path
import argparse, hashlib, json, importlib.util
import numpy as np

def main():
 ap=argparse.ArgumentParser();ap.add_argument('fixture',type=Path);ap.add_argument('run',type=Path);ap.add_argument('report',type=Path);a=ap.parse_args()
 f=json.loads(a.fixture.read_text(encoding='utf-8'));r=json.loads((a.run/'result.json').read_text(encoding='utf-8'));rows=json.loads((a.run/'結果節点.json').read_text(encoding='utf-8'))
 assert r['status'].split('|')[0]=='PASS' and r['flow']=='INCONSISTENT'
 assert hashlib.sha256(a.fixture.read_bytes()).hexdigest().upper()==r['fixture_sha256'].upper()
 spec=importlib.util.spec_from_file_location('q8_checks',Path(__file__).with_name('check_manual_case.py'));q8=importlib.util.module_from_spec(spec);spec.loader.exec_module(q8)
 ns={int(n[0]):n for n in f['nodes']};els=f['elements'];stage=max(x[0] for x in rows[1:] if isinstance(x[0],(int,float)));rr=[x for x in rows[1:] if x[0]==stage];u={int(x[1]):np.array(x[4:6],float) for x in rr}
 assert len(u)==len(ns) and all(np.isfinite(v).all() for v in u.values())
 coord_error=max(np.linalg.norm(np.array(x[2:4])-ns[int(x[1])][1:3]) for x in rr);assert coord_error<1e-12
 grid=[(r,s) for r in np.linspace(-1,1,11) for s in np.linspace(-1,1,11)];dr=np.array([q8.N(r+1e-30j,s).imag/1e-30 for r,s in grid]);ds=np.array([q8.N(r,s+1e-30j).imag/1e-30 for r,s in grid]);minimum={}
 for mode in ('reference','scale1'):
  xy=np.array([[np.array(ns[int(n)][1:3],float)+(u[int(n)] if mode=='scale1' else 0) for n in e[1:9]] for e in els]);aa=np.einsum('qn,end->eqd',dr,xy);bb=np.einsum('qn,end->eqd',ds,xy);det=aa[:,:,0]*bb[:,:,1]-aa[:,:,1]*bb[:,:,0];minimum[mode]=float(det.min());assert minimum[mode]>0
 mats={int(x[0]):x for x in f['materials']};area=sum(q8.quadrature(np.array([ns[int(n)][1:3] for n in e[1:9]],float)) for e in els);weight=sum(q8.quadrature(np.array([ns[int(n)][1:3] for n in e[1:9]],float))*mats[int(e[9])][3]*mats[int(e[9])][4] for e in els)
 fixed=max(abs(u[i][j]) for i,n in ns.items() for j in (0,1) if n[3+j]);assert fixed<1e-12
 reaction=sum(x[9] for x in rr if ns[int(x[1])][4]);free=sum(x[9] for x in rr if not ns[int(x[1])][4]);closure=reaction+free-weight;assert abs(closure)<1e-7
 assert abs(r['metrics'][4]-2.625)<1e-12 and abs(r['metrics'][6]-2.6375)<1e-12 and r['metrics'][7]<=.02
 out={'status':'PASS','fixture_sha256':r['fixture_sha256'],'release_book_sha256':r['book_sha256'],'nodes':len(ns),'elements':len(els),'stage':stage,'area_m2':area,'weight_kN':weight,'support_reaction_y_kN':reaction,'free_residual_y_kN':free,'global_closure_y_kN':closure,'reference_coordinate_error_m':float(coord_error),'fixed_displacement_max_m':float(fixed),'max_displacement_m':max(float(np.linalg.norm(v)) for v in u.values()),'dense_11x11_min_jacobian':minimum,'fs_pass':r['metrics'][4],'fs_fail':r['metrics'][6],'relative_residual':r['metrics'][3],'seconds':r['seconds']}
 a.report.parent.mkdir(parents=True,exist_ok=True);a.report.write_text(json.dumps(out,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');print(json.dumps(out,ensure_ascii=False))
if __name__=='__main__':main()
