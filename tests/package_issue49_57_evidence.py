"""Seal completed issue runs, independent reports, logs and source identities."""
import argparse, json, shutil
from pathlib import Path
from package_fixed_stress_evidence import archive, sha

def read(p): return json.loads(p.read_text(encoding='utf-8-sig'))
def package(work,out,book):
    assert not out.exists(), 'Preserve existing evidence'
    for marker in ('final_queue_done.txt','supplemental_done.txt','late_done.txt','cold_done.txt'):
        assert (work/marker).read_text().strip()=='PASS',marker
    reports={
        'package':'package3.json','all_inputs':'inputs_final.json','targeted':'targeted6_check.json',
        'material':'material4/check.json','frame_forces':'frame51_2/check_frame_forces.json',
        'coupled_regression':'full3_comparison.json','mixed':'mixed3/independent.json',
        'old_issue_states':'legacy3/states.json','water_analytic':'legacy3/water_analytic.json',
        'monolithic':'monolithic3.json','monolithic_anisotropic':'monolithic_biot3.json',
        'extra':'extra3_check.json','ui':'ui6/ui.json','manual':'manual/qa.json',
        'cold_named_partial':'cold3/named_partial/check.json',
        'cold_named_disp':'cold3/named_disp/check.json',
        'cold_water_hold':'cold3/water_hold/check.json'}
    data={name:read(work/path) for name,path in reports.items()}
    for name,d in data.items():
        if name=='water_analytic':
            assert len(d['cases'])==7
            assert all(c['status'] in ('PASS','REJECTED') for c in d['cases'].values())
        elif isinstance(d,dict): assert d.get('status')=='PASS',(name,d.get('status'))
    assert data['package']['package_sha256'].lower()==sha(book)
    assert len(data['targeted']['cases'])==22
    assert data['material']['records']==50
    assert data['manual']['visual_review']=='PASS'
    for name,source in (('named_partial','partial3'),('named_disp','targeted6/named_disp'),('water_hold','targeted6/water_hold')):
        continued=read(work/'cold3'/name/'continued.json');reference=read(work/source/'result.json')
        assert continued['fresh_excel_process'] and not continued['preloaded_frame']
        assert continued['build']=='20261024_FIXED_STRESS_07'
        assert max(abs(x-y) for x,y in zip(reference['metrics'][1:3],continued['metrics'][1:3]))<=1e-6
        assert max(abs(x-y) for x,y in zip(reference['metrics'][4:],continued['metrics'][4:]))<=1e-10
    for root in ('targeted6','full3'):
        for p in (work/root).glob('*/result.json'):
            assert read(p)['book_sha256'].lower()==sha(book),p
    out.mkdir(parents=True)
    for name,path in reports.items():
        shutil.copyfile(work/path,out/(name+'.json'))
    groups=[(name,work/name) for name in ('targeted6','partial3','reset_final3','full3','legacy3','mixed3',
             'material4','frame51_2','apex_off3','mixed_reset3','mixed_coupled3','ui6','cold3')]
    archives=[archive(out/'records.zip',groups,books=('water_linear',))]
    source_files=[p for p in Path('src/vba').iterdir() if p.suffix in ('.bas','.cls')]
    source={str(p).replace('\\','/'):sha(p) for p in source_files}
    # Include every compiled module, rather than the record archive's data-only
    # extension filter.
    import zipfile
    with zipfile.ZipFile(out/'source_snapshot.zip','w',zipfile.ZIP_DEFLATED) as z:
        for p in sorted(source_files):z.write(p,p.as_posix())
        z.writestr('source_sha256.json',json.dumps(source,indent=2))
    archives.append({'path':'source_snapshot.zip','bytes':(out/'source_snapshot.zip').stat().st_size,
                     'sha256':sha(out/'source_snapshot.zip'),'entries':len(source_files)+1})
    review=[p for p in (work/'manual').rglob('*') if p.suffix.lower() in ('.png','.jpg')]
    with zipfile.ZipFile(out/'manual_review.zip','w',zipfile.ZIP_DEFLATED) as z:
        for p in sorted(review):z.write(p,p.relative_to(work/'manual').as_posix())
    archives.append({'path':'manual_review.zip','bytes':(out/'manual_review.zip').stat().st_size,
                     'sha256':sha(out/'manual_review.zip'),'entries':len(review)})
    for name in ('build4.log','package3.log','final_queue_retry.log','targeted6.log','full3.log','legacy3.log','supplemental.log','mixed3.log'):
        shutil.copyfile(work/name,out/name)
    failed=work/'extra2/named_partial'
    if failed.exists():
        archives.append(archive(out/'partial_before_completion_fix.zip',[('pre_fix_partial',failed)]))
    summary={'status':'PASS','calculation_stamp':'20261024_FIXED_STRESS_07','book_sha256':sha(book),
             'module_sha256':source,'archives':archives,
             'reports':{name:{'path':name+'.json','sha256':sha(out/(name+'.json'))} for name in reports},
             'coordinate_collection':'reset_final3 is the accepted coordinate-preservation re-run; targeted6 reset_t6_nonaffine retains the earlier malformed test-driver coordinate collection for provenance',
             'pre_fix_partial':'before the final completed-load inheritance fix; expected mismatch, not an accepted validation run'}
    (out/'validation.json').write_text(json.dumps(summary,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    print('Sealed',len(source_files),'modules; book',sha(book))

if __name__=='__main__':
    ap=argparse.ArgumentParser()
    for name in ('work','out','book'):ap.add_argument('--'+name,type=Path,required=True)
    a=ap.parse_args();package(a.work,a.out,a.book)
