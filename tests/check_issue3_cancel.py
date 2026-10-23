"""Check controlled Excel cancellation evidence, without launching Excel."""
import argparse
import hashlib
import json
from pathlib import Path


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def check(folder, book):
    expected = {f'{native}_{point}' for native in ('OFF', 'AUTO')
                for point in ('first_FS', 'FS', 'FLAG', 'CAPTURE', 'SAVE', 'PLASTIC', 'WRITE_TIME', 'WRITE_STAGE')}
    expected |= {f'ordinary_{point}' for point in ('FS', 'CAPTURE', 'SAVE', 'PLASTIC')}
    actual = {p.parent.name: p for p in folder.glob('*/cancel.json')}
    assert set(actual) == expected, (set(actual) - expected, expected - set(actual))
    sha = hashlib.sha256(book.read_bytes()).hexdigest()
    cases = []
    for name, path in sorted(actual.items()):
        r = read(path)
        assert r['status'] == 'PASS', name
        assert r['book_sha256'].lower() == sha, name
        assert r['application_before'] == r['application_after'], name
        s = r['state']
        assert s[0] == 1, name
        if r['fault_number'] == 18:
            assert s[1:6] == [True, True, 18, False, False], name
            assert s[6] == s[8] and s[7] == s[9] and s[10] == s[11], name
            assert r['analysis'] == 'NONCONVERGED|ユーザーが解析を中断しました。', name
            for field in ('u', 'p', 'boundary'):
                assert r[f'max_accepted_{field}_difference'] == 0, (name, field)
            assert r['ledger_required'] and r['ledger_count'] == 48 and r['max_accepted_ledger_difference'] == 0, name
            assert r['displacement_count'] == r['boundary_count'] == 86 and r['pressure_count'] == 32, name
            assert s[7] == 0 if 'first' in name else s[7] >= 2, name
            if r['injection'] in {'WRITE_TIME', 'WRITE_STAGE'}:
                assert not list(path.parent.rglob('*.tmp')), ('Temporary output remained', name)
        else:
            assert r['fault_number'] == 5 and r['analysis'].startswith('PASS|'), name
            assert not s[1] and not s[2] and s[3] != 18, name
        cases.append({'case': name, 'status': 'PASS', 'fault_number': r['fault_number'],
                      'accepted_time': s[6], 'accepted_steps': s[7], 'successful_increments': s[10]})
    return {'status': 'PASS', 'book_sha256': sha, 'cases': cases,
            'scope': '16 cancellation cases and four ordinary-error controls; accepted displacement, GP pressure, boundary displacement and 48 ledger values; temporary I/O cleanup; Excel application state'}


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--root', type=Path, required=True)
    ap.add_argument('--book', type=Path, required=True)
    ap.add_argument('--out', type=Path, required=True)
    a = ap.parse_args()
    result = check(a.root, a.book)
    a.out.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print('PASS:', len(result['cases']), 'Excel fault/control cases')
