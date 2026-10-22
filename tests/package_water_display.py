"""Package the exact compiled/regression-tested water-view book and evidence."""
from pathlib import Path
import argparse, gzip, hashlib, json, shutil, zipfile

ROOT=Path(__file__).resolve().parents[1]
def read(p): return json.loads(p.read_text(encoding='utf-8-sig'))
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,d): p.write_text(json.dumps(d,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')

def main():
    ap=argparse.ArgumentParser();ap.add_argument('run',type=Path);a=ap.parse_args();run=a.run.resolve()
    assert (run/'queue.log').read_text(encoding='utf-8-sig').strip().endswith('ALL PASS')
    package=read(run/'package.json');checks=read(run/'water_display_check.json');sources=read(run/'sources.json')
    assert package['status']=='PASS' and sources['status']=='PASS'
    assert all(r['pass'] for r in checks['records']) and len(checks['rejections'])==7
    book=run/'release/2DSoilFEM_ui_20261023.xlsm';dll=run/'release/soilfem_native_x64.dll'
    assert sha(book).upper()==package['package_sha256']
    assert sha(dll)==sha(ROOT/'workbook/ui_20261023/soilfem_native_x64.dll')
    ev=ROOT/'docs/water-display-evidence/20261022';ev.mkdir(parents=True,exist_ok=True)
    for name in ('package.json','sources.json','water_display_check.json','independent.json','joint_check.json','queue.log','build.log'):
        shutil.copyfile(run/name,ev/name)
    cases=['cutoff','box_inside','stages','one_way','joint','wet_srm_INCONSISTENT','wet_srm_DAVIS','fixed13']
    for name in cases:
        target=ev/name;target.mkdir(exist_ok=True)
        for p in (run/name).glob('*.json'):
            if p.stem=='結果節点':
                compressed=target/(p.name+'.gz');compressed.write_bytes(gzip.compress(p.read_bytes(),mtime=0))
                assert gzip.decompress(compressed.read_bytes())==p.read_bytes()
                previous=target/p.name
                if previous.exists():previous.unlink()
            else:shutil.copyfile(p,target/p.name)
    for name in ('fixed_stress','time_reserved','missing_stage','conflicting_profile','orphan','duplicate','unknown_profile'):
        target=ev/('reject_'+name);target.mkdir(exist_ok=True)
        shutil.copyfile(run/('reject_'+name)/'input_check.json',target/'input_check.json')
    for folder in ('ui_cutoff','ui_cutoff_checked','ui_stages','ui_hydro'):
        target=ev/folder;target.mkdir(exist_ok=True)
        for p in (run/folder).iterdir():
            if p.suffix in ('.json','.txt') and p.name!='ready.txt':shutil.copyfile(p,target/p.name)
    with zipfile.ZipFile(ev/'calculation_logs.zip','w',zipfile.ZIP_DEFLATED) as z:
        for name in ['basic']+cases:
            for p in (run/name).rglob('*'):
                if p.is_file() and p.suffix.lower() not in ('.xlsm','.dll','.json'):
                    z.write(p,name+'/'+p.relative_to(run/name).as_posix())
    examples=ROOT/'docs/manuals/examples/water_display_examples.zip'
    with zipfile.ZipFile(examples,'w',zipfile.ZIP_DEFLATED) as z:
        z.write(dll,dll.name)
        for name in cases:
            books=list((run/name).glob('*.xlsm'));assert len(books)==1,(name,books)
            z.write(books[0],name+'/'+books[0].name)
        for name in ['groundwater_cutoff','groundwater_box_inside','groundwater_stages','water_stage_one_way','joint_water_wet','groundwater_slope_low_srm','water_slope_fixed13']:
            z.write(ROOT/'tests/fixtures'/f'{name}.json','fixtures/'+name+'.json')
        z.writestr('README.txt','Calculated copies from the water-display regression. Copy the single DLL at archive root beside the chosen XLSM before rerunning. ONE_WAY is implemented; FIXED_STRESS and time/drainage/iteration fields are reserved and rejected. Wet SRM near-limit displacements can be large: plots use reference geometry, not physical collapse predictions. See the manual chapter 8 and docs/WATER_STAGE_INPUT_PLAN.md.\n')
    public=ROOT/'workbook/ui_20261023/2DSoilFEM_ui_20261023.xlsm'
    old=read(public.parent/'BUILD_RECORD.json')
    old.setdefault('prior_mesh_release_book_sha256',old['book_sha256'])
    old.update(book_sha256=sha(book),water_view_build='20261022_WATER_VIEW_03',
               water_view_evidence='docs/water-display-evidence/20261022',
               water_view_package=package,water_view_sources=sources,
               water_view_regression=checks,solver_law_tolerance_changes=False)
    shutil.copyfile(book,public);write(public.parent/'BUILD_RECORD.json',old)
    with zipfile.ZipFile(ROOT/'workbook/SoilFEM_UI_20261023.zip','w',zipfile.ZIP_DEFLATED) as z:
        for p in (public,dll,public.parent/'BUILD_RECORD.json'):z.write(p,p.name)
        z.write(ROOT/'docs/GROUNDWATER.md','GROUNDWATER.md')
        z.write(ROOT/'docs/WATER_STAGE_INPUT_PLAN.md','WATER_STAGE_INPUT_PLAN.md')
    write(ev/'artifacts.json',{'book_sha256':sha(public),'dll_sha256':sha(dll),'examples_sha256':sha(examples)})
    print('Packaged compiled book, calculation records, logs, examples and single-DLL release.')

if __name__=='__main__':main()
