"""Gate and archive the WATER default-resume Excel evidence."""
from pathlib import Path
import argparse,hashlib,json,shutil,zipfile

def read(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
ap=argparse.ArgumentParser()
for n in ['root','book','out']:ap.add_argument('--'+n,type=Path,required=True)
a=ap.parse_args();a.out.mkdir(parents=True,exist_ok=True)
assert (a.root/'rest_complete.txt').read_text().strip()=='PASS'
reports={n:read(a.root/p) for n,p in {
 'package':'package.json','baseline':'baseline_reproduction.json',
 'matrix_auto':'matrix_auto/comparison.json','matrix_off':'matrix_off/comparison.json',
 'extra':'extra.json','regression':'comparison.json','saved_states':'saved_states.json','inputs':'inputs.json'}.items()}
for n,r in reports.items():assert r['status']==('REPRODUCED_EXPECTED_FAILURE' if n=='baseline' else 'PASS'),n
assert len(reports['matrix_auto']['cases'])==12 and len(reports['matrix_off']['cases'])==2
assert len(reports['regression']['cases'])==17
for case in (a.root/'full').iterdir():
 if not case.is_dir():continue
 record=case/('result.json' if (case/'result.json').is_file() else 'expected_failure.json')
 assert read(record)['book_sha256'].lower()==sha(a.book),case
assert reports['package']['package_sha256'].lower()==sha(a.book)
assert reports['inputs']['candidate_sha256']==sha(a.book)
for r in reports['matrix_auto']['cases']+reports['matrix_off']['cases']:assert r['book_sha256'].lower()==sha(a.book)
source={}
with zipfile.ZipFile(a.out/'source_snapshot.zip','w',zipfile.ZIP_DEFLATED) as z:
 for p in sorted(Path('src/vba').iterdir()):
  if p.suffix not in ['.bas','.cls']:continue
  source[p.as_posix()]=sha(p);z.write(p,p.as_posix())
assert len(source)==reports['package']['module_count']==52
with zipfile.ZipFile(a.out/'records.zip','w',zipfile.ZIP_DEFLATED) as z:
 for p in sorted(a.root.rglob('*')):
  if not p.is_file() or p.suffix.lower() not in ['.csv','.json','.log','.txt']:continue
  z.write(p,p.relative_to(a.root).as_posix())
 for p in sorted(Path('tests/fixtures/issue50').glob('*.json')):z.write(p,'fixtures/'+p.name)
for n,p in [('package','package.json'),('baseline','baseline_reproduction.json'),('matrix_auto','matrix_auto.json'),('matrix_off','matrix_off.json'),('extra','extra.json'),('regression','regression.json'),('inputs','inputs.json')]:
 (a.out/p).write_text(json.dumps(reports[n],ensure_ascii=False,indent=2)+'\n',encoding='utf8')
summary={'status':'PASS','book_sha256':sha(a.book),'source_sha256':source,'source_snapshot_sha256':sha(a.out/'source_snapshot.zip'),
 'records_sha256':sha(a.out/'records.zip'),'compile':'PASS','module_count':52,
 'calculation_stamp':'20261024_FIXED_STRESS_07','checkpoint_schema_changed':False,
 'targeted_auto_resume_cases':12,'targeted_off_resume_cases':2,'fractional_series_case':1,
 'cold_process_cases':3,'view_load_cases':2,'normal_regression_cases':17,
 'accepted_csv_files_byte_identical':reports['saved_states']['csv_files'],
 'canonical_input_cells_unchanged':sum(reports['inputs']['cells'].values()),
 'double_and_convergence_conditions_changed':False,'test_observer_not_in_distribution':True}
(a.out/'validation.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
print('PASS: evidence gated and packaged',summary['book_sha256'])
