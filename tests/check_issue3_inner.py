"""Verify inner CSV cancellation, ordinary errors, and accepted-state restoration."""
import argparse
import hashlib
import json
from pathlib import Path


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def check(root, book):
    points = {w+'_'+p for w in ('PLCSV', 'FSCSV', 'LEDGERCSV', 'GWCSV') for p in ('TIME', 'STAGE')}
    expected = {native+'_'+p for native in ('OFF', 'AUTO') for p in points}
    expected |= {'ordinary_'+p for p in points}
    actual = {p.parent.name: p for p in root.glob('*/cancel.json')}
    assert set(actual) == expected, (set(actual)-expected, expected-set(actual))
    sha = hashlib.sha256(book.read_bytes()).hexdigest()
    cases = []
    for name, p in sorted(actual.items()):
        r = read(p)
        assert r['status'] == 'PASS' and r['book_sha256'].lower() == sha, name
        assert r['application_before'] == r['application_after'], name
        log = (p.parent/'column_inner_io_out/run.log').read_bytes()
        expected_native = b'NativeDLL=VBA(OFF_SELECTED)' if name.startswith('OFF_') else b'NativeDLL=NATIVE('
        assert expected_native in log, ('Actual native selection differs', name)
        s = r['state']
        assert s[0] == 1, name
        if r['fault_number'] == 18:
            assert s[1:6] == [True, True, 18, False, False], (name, s)
            assert r['analysis'] == 'NONCONVERGED|ユーザーが解析を中断しました。', name
            assert s[6] == s[8] and s[7] == s[9] and s[10] == s[11], name
            assert r['displacement_count'] == r['boundary_count'] == 86 and r['pressure_count'] == 32, name
            for field in ('u', 'p', 'boundary', 'ledger'):
                assert r['max_accepted_'+field+'_difference'] == 0, (name, field)
            assert r['ledger_required'] and r['ledger_count'] == 48, name
            assert not list(p.parent.rglob('*.tmp')), ('Temporary file/folder remained', name)
            assert s[6] < 95040, 'Following WAIT unexpectedly ran: '+name
        else:
            assert r['fault_number'] == 5 and not s[1] and s[3] != 18, name
            if name.endswith('_TIME'):
                assert r['ordinary_failure_expected'] and '時刻別状態の保存に失敗しました' in r['analysis'], name
                assert r['analysis'].startswith('RUNTIME_ERROR|') and s[2] and not s[4] and not s[5], name
            else:
                assert not r['ordinary_failure_expected'] and r['analysis'].startswith('PASS|') and not s[2], name
                assert s[6] == 95040, name
        cases.append({'case': name, 'status': 'PASS', 'error': r['fault_number'],
                      'accepted_time': s[6], 'accepted_steps': s[7],
                      'classification': r['analysis']})
    return {'status': 'PASS', 'book_sha256': sha, 'cases': cases,
            'scope': 'four inner CSV writers x stage/time x VBA/native =16 cancellations; eight Err5 controls; all U/P/boundary/ledger values, no later WAIT, temporary-folder cleanup and Excel settings'}


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    for n in ('root', 'book', 'out'):
        ap.add_argument('--'+n, required=True, type=Path)
    a = ap.parse_args()
    result = check(a.root, a.book)
    a.out.write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
    print('PASS: 16 nested-writer cancellations and eight ordinary-error controls')
