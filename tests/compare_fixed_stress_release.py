"""Check that the caption-only release preserves all recorded numerical fields."""
import argparse
import json
from pathlib import Path


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def compare(reference, candidate):
    results = []
    for folder in sorted(candidate.iterdir()):
        if not folder.is_dir():
            continue
        old = reference / folder.name
        if (folder / 'result.json').exists():
            checks = []
            for name in ('result.json', 'resume.json'):
                a, b = read(old/name), read(folder/name)
                assert a['status'].startswith('PASS|') and b['status'].startswith('PASS|'), folder.name
                if name == 'result.json':
                    assert a['fixture_sha256'] == b['fixture_sha256'], folder.name
                for key in ('disp', 'stress', 'metrics'):
                    assert len(a[key]) == len(b[key]), (folder.name, key)
                    delta = max((abs(x-y) for x, y in zip(a[key], b[key])), default=0)
                    assert delta == 0, (folder.name, name, key, delta)
                    checks.append({'record': name, 'field': key, 'values': len(a[key]), 'max_difference': delta})
            results.append({'case': folder.name, 'status': 'PASS', 'checks': checks})
        else:
            a, b = read(old/'expected_failure.json'), read(folder/'expected_failure.json')
            assert a['expected'] == b['expected'] and b['expected'] in b['status'], folder.name
            assert not b['status'].startswith('PASS|'), folder.name
            for key in ('coupled', 'nodes', 'pressure', 'stress', 'profile'):
                assert a[key] == b[key], (folder.name, key)
            results.append({'case': folder.name, 'status': 'EXPECTED_REJECTION_UNCHANGED'})
    assert len(results) == 17, len(results)
    return {'status': 'PASS', 'scope': 'caption-only final release versus numerical candidate: all recorded displacement/stress slots/metrics, fresh and resume; rejected-state fields unchanged', 'cases': results}


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    for name in ('reference', 'candidate', 'report'):
        ap.add_argument('--'+name, type=Path, required=True)
    args = ap.parse_args()
    result = compare(args.reference, args.candidate)
    args.report.write_text(json.dumps(result, indent=2), encoding='utf-8')
    print(json.dumps({'status': result['status'], 'cases': len(result['cases'])}))
