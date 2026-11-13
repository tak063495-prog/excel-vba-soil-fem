"""Compare actual Excel sparse-direct, band, VBA and missing-DLL paths with resume."""
from pathlib import Path
import argparse,json,shutil,subprocess,sys
import numpy as np
from workbook_xml import read
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--runner',type=Path);ap.add_argument('--start-case');ap.add_argument('--case',dest='selected_cases',action='append');a=ap.parse_args()
root=Path(__file__).resolve().parents[1];a.out=a.out.resolve();assert not a.out.exists();a.out.mkdir(parents=True)
runner=(a.runner or root/'tests/run_groundwater_fixture.ps1').resolve();records=[]
cases=a.selected_cases or ['face_support/static','linear_solver/gmres_associated_plastic','mixed_t6/mixed_gravity','mixed_t6/mixed_water','linear_solver/gmres_water_plastic','face_support/opening_cycle']
if a.start_case:cases=cases[cases.index(a.start_case):]
for case in cases:
 name=case.split('/')[-1];original=json.loads((root/'tests/fixtures'/f'{case}.json').read_text(encoding='utf-8'))
 for mode in ('SPARSE','BAND','OFF','MISSING'):
  data=json.loads(json.dumps(original));data.setdefault('settings',{})['NATIVE_SPARSE_DIRECT']=0 if mode=='BAND' else 1
  if case=='mixed_t6/mixed_gravity':
   # This mixed-element fixture is elastic. Explicit association exercises the symmetric selection.
   for material in data['materials']:material[7]=material[5]
  fixture=a.out/f'{name}_{mode.lower()}.json';fixture.write_text(json.dumps(data),encoding='utf-8');book=a.book.resolve()
  if mode=='MISSING':
   directory=a.out/f'{name}_missing_dll';directory.mkdir();shutil.copy2(book,directory/book.name);shutil.copy2(book.parent/'soilfem_native_x64.dll',directory/'soilfem_native_x64.dll');book=directory/book.name
  dest=a.out/f'{name}_{mode.lower()}';log=a.out/f'{name}_{mode.lower()}.log'
  print('START',case,mode,flush=True)
  with log.open('w',encoding='utf-8') as stream:
   r=subprocess.run(['pwsh','-NoProfile','-File',str(runner),'-Book',str(book),'-Fixture',str(fixture),'-Out',str(dest),'-Flow','INCONSISTENT','-Native','OFF' if mode=='OFF' else 'AUTO','-Resume'],cwd=root,stdout=stream,stderr=subprocess.STDOUT)
  if r.returncode:print(log.read_text(encoding='utf-8')[-4000:],flush=True);sys.exit(1)
  result=json.loads((dest/'result.json').read_text(encoding='utf-8-sig'));resume=json.loads((dest/'resume.json').read_text(encoding='utf-8-sig'))
  assert result['status']=='PASS|' and resume['status']=='PASS|'
  wants_sparse=mode=='SPARSE' and case!='linear_solver/gmres_water_plastic'
  solver=read(dest/'fresh_results.xlsm')['診断']['B95'];assert solver==('CSR_GMRES' if case=='linear_solver/gmres_water_plastic' else 'SPARSE_LDLT' if wants_sparse else 'BAND_LU'),(case,mode,solver)
  event=(dest/'initial_run_logs'/(data['name']+'_out')/'solver_events.csv').read_text(encoding='cp932')
  assert ('SPARSE_DIRECT_FACTOR' in event)==wants_sparse,(case,mode)
  assert np.isfinite(result['disp']).all() and np.isfinite(result['stress']).all()
  record={'case':case,'mode':mode,'status':'PASS','solver':solver,'metrics':result['metrics'],'seconds':result['seconds']};records.append(record)
  (a.out/'results.json').write_text(json.dumps(records,indent=2,allow_nan=False),encoding='utf-8');print('PASS',case,mode,flush=True)
  if mode=='SPARSE':reference=result
  else:
   du=float(np.max(abs(np.array(reference['disp'])-result['disp'])));ds=float(np.max(abs(np.array(reference['stress'])-result['stress'])))
   assert du<1e-7 and ds<1e-4,(case,mode,du,ds)
   record.update(max_disp_difference=du,max_stress_difference=ds)
   (a.out/'results.json').write_text(json.dumps(records,indent=2,allow_nan=False),encoding='utf-8')
print('PASS',len(records),'real Excel/resume comparisons',flush=True)
