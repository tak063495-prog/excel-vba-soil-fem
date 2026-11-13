"""Compare static stepping through real Excel analyses and checkpoint resume."""
from pathlib import Path
import argparse,json,subprocess,sys
import numpy as np

ap=argparse.ArgumentParser(); ap.add_argument('--book',type=Path,required=True)
ap.add_argument('--out',type=Path,required=True); ap.add_argument('--runner',type=Path)
a=ap.parse_args(); root=Path(__file__).resolve().parents[1]
a.out=a.out.resolve(); a.out.mkdir(parents=True,exist_ok=False)
runner=(a.runner or root/'tests/run_groundwater_fixture.ps1').resolve(); records=[]
for case in ('face_support/static','linear_solver/gmres_associated_plastic'):
 original=json.loads((root/'tests/fixtures'/f'{case}.json').read_text(encoding='utf-8'))
 baseline=None
 for label,keep,count in (('legacy',0,0),('keep10',1,10),('keep20',1,20)):
  data=json.loads(json.dumps(original)); settings=data.setdefault('settings',{})
  settings.update(STATIC_KEEP_ACCEPTED_STEP=keep,STATIC_INITIAL_STEPS=count)
  name=original['name']+'_'+label; data['name']=name
  fixture=a.out/f'{name}.json'; fixture.write_text(json.dumps(data),encoding='utf-8')
  dest=a.out/name; log=a.out/f'{name}.log'; print('START',name,flush=True)
  with log.open('w',encoding='utf-8') as stream:
   result=subprocess.run(['pwsh','-NoProfile','-File',str(runner),'-Book',str(a.book.resolve()),'-Fixture',str(fixture),'-Out',str(dest),'-Flow','INCONSISTENT','-Native','AUTO','-Resume'],cwd=root,stdout=stream,stderr=subprocess.STDOUT)
  if result.returncode:print(log.read_text(encoding='utf-8')[-3000:],flush=True);sys.exit(1)
  fresh=json.loads((dest/'result.json').read_text(encoding='utf-8-sig'))
  resume=json.loads((dest/'resume.json').read_text(encoding='utf-8-sig'))
  assert fresh['status']=='PASS|' and resume['status']=='PASS|',(name,fresh['status'],resume['status'])
  assert np.isfinite(fresh['disp']).all() and np.isfinite(fresh['stress']).all()
  if baseline is None:baseline=fresh
  du=float(np.max(abs(np.array(fresh['disp'])-baseline['disp'])))
  ds=float(np.max(abs(np.array(fresh['stress'])-baseline['stress'])))
  if case=='face_support/static':assert du<1e-8 and ds<1e-5,(name,du,ds)
  record={'case':case,'policy':label,'status':'PASS','resume':'PASS','metrics':fresh['metrics'],'max_disp_difference_from_legacy':du,'max_stress_difference_from_legacy':ds}
  records.append(record);(a.out/'results.json').write_text(json.dumps(records,indent=2,allow_nan=False),encoding='utf-8')
  print('PASS',name,'du',du,'ds',ds,flush=True)
print('PASS',len(records),'real analyses and resumes; plastic load-path differences are reported, not assumed zero',flush=True)
