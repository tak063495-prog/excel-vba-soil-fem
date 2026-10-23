"""Package completed Issue3/59 records; refuses incomplete acceptance results."""
import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    for key in ('root', 'book', 'out'):
        ap.add_argument('--' + key, type=Path, required=True)
    a = ap.parse_args()
    if a.out.exists():
        raise FileExistsError('Preserve previous evidence: ' + str(a.out))
    assert (a.root / 'final7_complete.txt').read_text(encoding='utf-8-sig').strip() == 'PASS'
    reports = {
        'package': 'package7.json', 'inputs': 'all_inputs7.json',
        'geometry': 'issue597.json', 'independent_geometry': 'geometry_python6.json',
        'coupled': 'comparison7.json', 'saved_states': 'saved_states7.json',
        'mixed': 'mixed6/independent.json', 'cancellation': 'fault7.json',
        'keyboard': 'keyboard7/cancel.json', 'subdivision': 'subdivision7/geometry.json',
    }
    values = {key: read(a.root / path) for key, path in reports.items()}
    for key, value in values.items():
        assert value['status'] == 'PASS', key
    digest = sha(a.book)
    for key in ('geometry', 'cancellation', 'keyboard', 'subdivision'):
        assert values[key]['book_sha256'].lower() == digest, key
    assert values['package']['package_sha256'].lower() == digest
    assert values['inputs']['candidate_sha256'].lower() == digest
    assert sum(values['inputs']['cells'].values()) == 10115
    assert len(values['cancellation']['cases']) == 20
    assert len(values['coupled']['cases']) == 17
    assert len(values['geometry']['cases']) == 9
    assert values['keyboard']['injection'] == 'FS_UI'
    assert values['keyboard']['ledger_required'] and values['keyboard']['max_accepted_ledger_difference'] == 0
    assert values['keyboard']['application_before'] == values['keyboard']['application_after'] == [True, True, -4105, 1]
    assert values['subdivision']['cases'][0]['case'] == 'curved_subdivision'
    assert values['subdivision']['cases'][0]['certificate']
    mixed_book = read(a.root / 'mixed6/q8_candidate_srm/result.json')['book_sha256'].lower()
    assert len(list((a.root / 'mixed6').glob('*/result.json'))) == 17
    assert not list((a.root / 'mixed6').rglob('fixed_stress.csv'))
    pattern = r'verify\[reopen\] FEM OK sha256=([A-F0-9]+)'
    old_geometry = re.search(pattern, (a.root / 'build6.log').read_text(encoding='utf-8-sig')).group(1)
    new_geometry = re.search(pattern, (a.root / 'build7.log').read_text(encoding='utf-8-sig')).group(1)
    assert old_geometry == new_geometry
    old = read(a.root / 'mixed6/q8_reference_srm/result.json')
    new = read(a.root / 'mixed6/q8_candidate_srm/result.json')
    for field in ('disp', 'stress', 'metrics'):
        assert old[field] == new[field], ('Q8 exact regression', field)
    values['q8_exact'] = {'status': 'PASS', 'max_u_difference': 0, 'max_stress_difference': 0, 'max_metric_difference': 0,
                          'dof': len(new['disp']), 'stress_values': len(new['stress']), 'fs_interval': new['metrics'][5:8],
                          'reference_seconds': old['seconds'], 'candidate_seconds': new['seconds']}
    reports['q8_exact'] = 'derived_from_mixed6'
    a.out.mkdir(parents=True)
    for key, value in values.items():
        (a.out / (key + '.json')).write_text(json.dumps(value, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    for name in ('build7.log', 'package7.log', 'input597.log', 'full7.log', 'mixed6.log', 'final_fault7.log', 'final_queue7.log', 'keyboard7.log'):
        (a.out / name).write_bytes((a.root / name).read_bytes())
    sources = sorted(Path('src/vba').glob('*.bas')) + sorted(Path('src/vba').glob('*.cls'))
    source_hashes = {p.as_posix(): sha(p) for p in sources}
    assert len(sources) == 52
    with zipfile.ZipFile(a.out / 'source_snapshot.zip', 'w', zipfile.ZIP_DEFLATED) as z:
        for p in sources:
            z.write(p, p.as_posix())
        z.writestr('source_sha256.json', json.dumps(source_hashes, indent=2))
    folders = ('before_fs', 'before_ledger6b', 'input597', 'full7', 'mixed6', 'final_fault7', 'keyboard7', 'subdivision7')
    count = 0
    with zipfile.ZipFile(a.out / 'records.zip', 'w', zipfile.ZIP_DEFLATED) as z:
        for folder in folders:
            for p in sorted((a.root / folder).rglob('*')):
                if p.is_file() and p.suffix.lower() in {'.json', '.csv', '.log', '.txt'}:
                    z.write(p, p.relative_to(a.root).as_posix())
                    count += 1
    result = {
        'status': 'PASS', 'book_sha256': digest, 'reference_commit': 'e290fb9e4942546997dc3436c9c7b14f363f1c17',
        'reference_book_sha256': 'c0e75c6c9fad2b47d6b37680fa6e591126031989860e03c9bf55ee81a356840f',
        'calculation_stamp': '20261024_FIXED_STRESS_07', 'compile_modules': 52,
        'source_sha256': source_hashes, 'records_files': count,
        'regression_cases_on_distribution': {'fixed_stress': 17, 'invalid_valid_geometry_input': 9, 'cancel_and_controls': 20},
        'unchanged_path_regression': {'mixed_q8_t6': 17, 'book_sha256': mixed_book, 'geometry_normalized_sha256': new_geometry.lower(),
                                     'scope': 'candidate6: same Q8 certificate, mechanics, material, I/O and DLL. Only FSRunStage ledger rollback changed in candidate7; these fixtures do not activate fixed-stress.'},
        'additional_checks': ['nine direct/runtime geometry probes', 'positive curved Q8 requiring subdivision', 'actual Ctrl+Break with events enabled and automatic calculation restored'],
        'saved_state_csv_files': values['saved_states']['csv_files'],
        'reports': {key: key + '.json' for key in reports},
        'pdf_sha256_unchanged': sha(Path('docs/manuals/2DSoilFEM_Manual_20261024.pdf')),
    }
    (a.out / 'validation.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print('PASS:', count, 'record files; accepted distribution', digest)
