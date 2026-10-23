"""Publish only completed nested-writer Excel evidence; retain the original baselines."""
import argparse
import hashlib
import json
from pathlib import Path
import zipfile
from check_issue3_inner import check


def read(p):
    return json.loads(p.read_text(encoding='utf-8-sig'))


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    for n in ('root', 'book', 'out'):
        ap.add_argument('--'+n, required=True, type=Path)
    a = ap.parse_args()
    assert not a.out.exists(), 'Preserve previous evidence'
    assert (a.root/'complete.txt').read_text(encoding='utf-8-sig').strip() == 'PASS'
    fault = check(a.root/'faults', a.book)
    keys = {'package':'package.json', 'inputs':'inputs.json', 'coupled':'comparison.json',
            'saved_states':'saved_states.json', 'keyboard':'keyboard/cancel.json'}
    reports = {k:read(a.root/v) for k,v in keys.items()}
    for k,r in reports.items():
        assert r['status'] == 'PASS', k
    digest = sha(a.book)
    assert reports['package']['package_sha256'].lower() == digest
    assert reports['package']['module_count'] == 52
    assert reports['inputs']['candidate_sha256'].lower() == digest
    assert sum(reports['inputs']['cells'].values()) == 10115
    assert len(reports['coupled']['cases']) == 17
    assert reports['saved_states']['csv_files'] == 1853
    k = reports['keyboard']
    assert k['injection'] == 'FSCSV_TIME_UI' and k['book_sha256'].lower() == digest
    assert k['state'][1:6] == [True,True,18,False,False]
    assert k['state'][6] == k['state'][8] and k['state'][7] == k['state'][9] and k['state'][10] == k['state'][11]
    assert k['application_before'] == k['application_after'] == [True,True,-4105,1]
    for f in ('u','p','boundary','ledger'):
        assert k['max_accepted_'+f+'_difference'] == 0, f
    assert not list((a.root/'keyboard').rglob('*.tmp'))
    before = read(a.root/'before_following/cancel.json')
    assert before['analysis'].startswith('PASS|') and not before['state'][1]
    assert before['state'][6] > before['state'][8] and before['state'][10] > before['state'][11]
    wrong = read(a.root/'before_fs_time/cancel.json')
    assert wrong['analysis'].startswith('RUNTIME_ERROR|') and wrong['state'][3] != 18
    a.out.mkdir(parents=True)
    reports['cancellation'] = fault
    for key,value in reports.items():
        (a.out/(key+'.json')).write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    for name in ('build.log','package.log','faults.log','full.log','queue.log','keyboard.log'):
        (a.out/name).write_bytes((a.root/name).read_bytes())
    count = 0
    with zipfile.ZipFile(a.out/'records.zip','w',zipfile.ZIP_DEFLATED) as z:
        for folder in ('before_pl_stage','before_fs_time','before_following','faults','full','keyboard'):
            for p in sorted((a.root/folder).rglob('*')):
                if p.is_file() and p.suffix in ('.csv','.json','.log','.txt'):
                    z.write(p,p.relative_to(a.root).as_posix());count+=1
    sources=sorted(Path('src/vba').glob('*.bas'))+sorted(Path('src/vba').glob('*.cls'))
    hashes={p.as_posix():sha(p) for p in sources}
    assert len(hashes)==52
    with zipfile.ZipFile(a.out/'source_snapshot.zip','w',zipfile.ZIP_DEFLATED) as z:
        for p in sources:z.write(p,p.as_posix())
        z.writestr('source_sha256.json',json.dumps(hashes,indent=2))
    result={'status':'PASS','book_sha256':digest,'reference_commit':'3db32347806361ab3d12d18016332da15674d747',
            'reference_book_sha256':'98952de07b4cf389ee4503704a2824a1cfc3a7210d3a04a4fe570df168b2d251',
            'calculation_stamp':'20261024_FIXED_STRESS_07','source_sha256':hashes,'compile_modules':52,
            'inner_cancellation_cases':16,'ordinary_error_controls':8,'coupled_cases':17,
            'accepted_state_csv_identical':1853,'record_files':count,'actual_keyboard':'FSWriteCheckpoint row loop, time save',
            'pdf_sha256_unchanged':sha(Path('docs/manuals/2DSoilFEM_Manual_20261024.pdf'))}
    (a.out/'validation.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print('PASS: nested CSV cancellation and accepted distribution',digest)
