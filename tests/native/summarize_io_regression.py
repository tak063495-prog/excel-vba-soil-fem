"""Exact numerical/log comparison for paired I/O fixed-Fs regression runs."""
import argparse
import json
import math
from pathlib import Path


def rows(path):
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    return data if isinstance(data, list) else [data]


def compare(reference, candidate, scope='io'):
    assert len(reference) == len(candidate) > 0
    report = []
    for r, c in zip(reference, candidate):
        if scope == 'events':
            assert r['IoFast'] == c['IoFast'] == 1
            assert r['EventFast'] == 0 and c['EventFast'] == 1
        else:
            assert scope == 'io' and r['IoFast'] == 0 and c['IoFast'] == 1
            assert r.get('EventFast', 0) == c.get('EventFast', 0) == 0
        assert all(x['StateCopy'] == x['GmresSimd'] == 1 and x['ExclusiveExcel'] is True for x in (r, c))
        for key in ('Case', 'Variant', 'Mode', 'Precision', 'NativeDllSha256', 'Fos', 'Bracket',
                    'Disp', 'SolverEventsSha256', 'SolverEventsBytes'):
            assert r[key] == c[key], (r['Case'], r['Variant'], key)
        assert r['NativeDllSha256'] and r['SolverEventsSha256'] and r['SolverEventsBytes'] > 0
        assert len(r['Bracket']) == 6 and all(math.isfinite(v) for v in r['Bracket'][:4])
        assert r['Disp'] and all(math.isfinite(v) for v in r['Disp'])
        a, b = r['Run'].split('|'), c['Run'].split('|')
        assert a[1:3] == b[1:3] == ['PASS', 'True']
        assert a[3:9] == b[3:9] and a[14:] == b[14:]
        assert math.isfinite(float(b[3])) and float(b[3]) >= 0 and 'fallback=0' in b[8]
        report.append(dict(case=r['Case'], variant=r['Variant'], displacement_count=len(r['Disp']),
                           max_abs_displacement_delta=0, residual=float(b[3]), bracket=c['Bracket'],
                           reference_seconds=r['WallSeconds'], candidate_seconds=c['WallSeconds'],
                           solver_events_sha256=c['SolverEventsSha256'], logs_exact=True))
    return report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('reference'); ap.add_argument('candidate'); ap.add_argument('--out', required=True)
    ap.add_argument('--scope', choices=('io', 'events'), default='io')
    args = ap.parse_args()
    result = dict(status='PASS', cases=compare(rows(args.reference), rows(args.candidate), args.scope))
    text = json.dumps(result, ensure_ascii=False, indent=2)
    Path(args.out).write_text(text + '\n', encoding='utf-8')
    print(text)


if __name__ == '__main__':
    main()
