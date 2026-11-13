"""Real Excel validation of the optional nonsymmetric GMRES solver."""
from pathlib import Path
import subprocess,json,numpy as np,sys,argparse
ap=argparse.ArgumentParser(description="Compare explicit CSR-GMRES with native OFF and nonsymmetric BAND on real Excel fixtures.");ap.add_argument("--book",type=Path,required=True);ap.add_argument("--out",type=Path,required=True);ap.add_argument("--runner",type=Path);args=ap.parse_args()
from workbook_xml import read
root=Path(__file__).resolve().parents[1];out=args.out.resolve();assert not out.exists(),'Preserve previous evidence';out.mkdir(parents=True)
runner=(args.runner or root/'tests/run_groundwater_fixture.ps1').resolve();book=args.book.resolve();records=[]
for case in ['gmres_water_plastic','gmres_mixed_gravity']:
 for mode in ['AUTO','OFF','BAND']:
  fixture=root/'tests/fixtures/linear_solver'/f'{case}.json';native=mode
  if mode=='BAND':
   f=json.loads(fixture.read_text(encoding='utf-8'));f['settings']['NONSYM_SOLVER']='BAND';fixture=out/f'{case}_band.json';fixture.write_text(json.dumps(f),encoding='utf-8');native='AUTO'
  dest=out/f'{case}_{mode.lower()}';log=out/f'{case}_{mode.lower()}.log'
  with log.open('w',encoding='utf-8') as f:r=subprocess.run(['pwsh','-NoProfile','-File',str(runner),'-Book',str(book),'-Fixture',str(fixture),'-Out',str(dest),'-Flow','INCONSISTENT','-Native',native,'-Resume'],cwd=root,stdout=f,stderr=subprocess.STDOUT)
  if r.returncode:print(log.read_text(encoding='utf-8')[-3000:],flush=True);sys.exit(1)
  data=json.loads((dest/'result.json').read_text(encoding='utf-8-sig'));resume=json.loads((dest/'resume.json').read_text(encoding='utf-8-sig'));assert data['status']=='PASS|' and resume['status']=='PASS|'
  d=read(dest/'fresh_results.xlsm');solver=d['診断']['B95'];assert solver==('NONSYM_BAND_LU' if mode=='BAND' else 'CSR_GMRES'),solver
  assert data['metrics'][3]<=1e-5
  if case=='gmres_water_plastic':assert sum(abs(np.array(data['stress']).reshape(-1,17)[:,9])<1e-6)>0,'Fixture must exercise plastic return'
  record={'case':case,'mode':mode,'status':'PASS','solver':solver,'metrics':data['metrics'],'native':data['native'],'seconds':data['seconds']};records.append(record)
  (out/'results.json').write_text(json.dumps(records,indent=2),encoding='utf-8');print('PASS',case,mode,flush=True)
 aa=json.loads((out/f'{case}_auto/result.json').read_text(encoding='utf-8-sig'))
 for mode in ['off','band']:
  bb=json.loads((out/f'{case}_{mode}/result.json').read_text(encoding='utf-8-sig'));du=np.max(abs(np.array(aa['disp'])-bb['disp']));ds=np.max(abs(np.array(aa['stress'])-bb['stress']));assert du<1e-7 and ds<1e-4,(case,mode,du,ds);print('MATCH',case,mode,du,ds,flush=True)
