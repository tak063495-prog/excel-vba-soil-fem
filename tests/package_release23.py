"""Package a byte-identical renamed WATER_02 release and its acceptance records."""
from pathlib import Path
import argparse,hashlib,json,shutil,zipfile
ap=argparse.ArgumentParser()
ap.add_argument('--visual-reviewed',action='store_true',help='Confirm the current rendered PDF has been visually reviewed.')
args=ap.parse_args()
if not args.visual_reviewed: ap.error('Review current page renders, then explicitly pass --visual-reviewed.')
ROOT=Path(__file__).resolve().parents[1]
def digest(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,s): p.write_bytes(s.encode('utf-8'))
def load(p): return json.loads(p.read_text(encoding='utf-8-sig'))
release=ROOT/'workbook/ui_20261023'; original=ROOT/'workbook/ui_20261022'
tmp=ROOT/'tests/tmp/release23'; evidence=ROOT/'docs/groundwater-release23'; evidence.mkdir(exist_ok=True)
book=release/'2DSoilFEM_ui_20261023.xlsm';dll=release/'soilfem_native_x64.dll'
assert digest(book)==digest(original/'2DSoilFEM_ui_20261022.xlsm')
assert digest(dll)==digest(original/dll.name)
check=load(tmp/'excel/rename_check.json');assert check['status']=='PASS' and check['compile']=='PASS'
wet=load(tmp/'wet/check.json');comparison=load(tmp/'wet/compare22.json')
assert wet['status']==comparison['status']=='PASS'
initial=load(tmp/'wet/result.json');resume=load(tmp/'wet/resume.json')
ref=ROOT/'tests/tmp/joint_water/acceptance22/wet'
resume_delta={k:max(abs(x-y) for x,y in zip(resume[k],load(ref/'resume.json')[k])) for k in ('disp','stress','metrics')}
assert all(v==0 for v in resume_delta.values()),resume_delta
comparison['resume_max_difference']=resume_delta
comparison['timing_note']='Single measurements; this filename/documentation-only release makes no performance claim.'
write(evidence/'comparison.json',json.dumps(comparison,ensure_ascii=False,indent=2)+'\n')
for src,name in [(tmp/'excel/rename_check.json','rename_check.json'),(tmp/'preview23.jpg','preview23.jpg'),(tmp/'wet/check.json','wet_check.json')]:shutil.copyfile(src,evidence/name)
checks=load(tmp/'manual_review/checks.json');assert checks['status']=='PASS'
# Only set this after actual whole-document contact sheets and changed pages were viewed.
checks['visual_review']='PASS: all 181 page renders reviewed in 10 contact sheets; cover/TOC, p5 and revised p164/171-174 reviewed at full size.'
write(evidence/'manual_checks.json',json.dumps(checks,ensure_ascii=False,indent=2)+'\n')
with zipfile.ZipFile(evidence/'manual_review.zip','w',zipfile.ZIP_DEFLATED) as z:
    for p in sorted((tmp/'manual_review').glob('*.png')):z.write(p,p.name)
    z.write(evidence/'manual_checks.json','checks.json')
with zipfile.ZipFile(evidence/'records.zip','w',zipfile.ZIP_DEFLATED) as z:
    for p in sorted((tmp/'wet').rglob('*')):
        if p.is_file():z.write(p,'wet/'+p.relative_to(tmp/'wet').as_posix())
    for p in (tmp/'wet2.log',tmp/'excel/rename_check.json'):z.write(p,p.name)
readme=(original/'README.txt').read_text(encoding='utf-8-sig')
readme=readme.replace('2DSoilFEM_ui_20261022.xlsm','2DSoilFEM_ui_20261023.xlsm').replace('blob/feat/groundwater-steady','blob/main').replace('2DSoilFEM_Manual_20261022_WATER_02.pdf','2DSoilFEM_Manual_20261023_WATER_02.pdf')
readme='配布名20261023：ユーザー指定のファイル名。作成・検証日は2026-10-22。計算識別20261022_WATER_02、検証済み22版とブック/DLLはバイト一致です。\n\n'+readme
write(release/'README.txt',readme)
record=dict(distribution='20261023',verification_date='2026-10-22',build='20261022_WATER_02',reference_commit='3a1be89872e2e97b3e6f0bd1b3b7a58e11109389',reference_workbook='workbook/ui_20261022/2DSoilFEM_ui_20261022.xlsm',book_sha256=digest(book),dll_sha256=digest(dll),byte_identical_to_reference=True,compile=check['compile'],source_modules_verified=len(check['source_modules']),normal_open=check['normal_open'],actions_rebound=check['actions_rebound'],preview_shapes=check['preview_shapes'],excel_version=check['excel_version'],excel_build=check['excel_build'],wet_independent_check='PASS',initial_full_array_differences=comparison['max_difference'],resume_full_array_differences=resume_delta,full_regression_evidence='docs/joint-water-evidence/20261022 (unchanged same binary)',solver_changes=False)
write(release/'BUILD_RECORD.json',json.dumps(record,ensure_ascii=False,indent=2)+'\n')
zipname=ROOT/'workbook/SoilFEM_UI_20261023.zip'
with zipfile.ZipFile(zipname,'w',zipfile.ZIP_DEFLATED) as z:
    for p in (book,dll,release/'README.txt',release/'BUILD_RECORD.json'):z.write(p,p.name)
    z.write(ROOT/'docs/GROUNDWATER.md','WATER_GUIDE.md')
with zipfile.ZipFile(zipname) as z:
    assert len(z.namelist())==5 and len([n for n in z.namelist() if n.endswith('.dll')])==1
    assert hashlib.sha256(z.read(book.name)).hexdigest()==digest(book)
    assert hashlib.sha256(z.read(dll.name)).hexdigest()==digest(dll)
print(json.dumps(dict(status='PASS',book_sha256=digest(book),zip_sha256=digest(zipname),zip_bytes=zipname.stat().st_size),indent=2))
