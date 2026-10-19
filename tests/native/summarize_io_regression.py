"""Exact numerical/log comparison for paired I/O or assembly fixed-Fs runs."""
import argparse
import json
import math
import re
from pathlib import Path


def rows(path):
    data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
    return data if isinstance(data, list) else [data]


def compare(reference, candidate, scope='io'):
    assert len(reference) == len(candidate) > 0
    report = []
    for r, c in zip(reference, candidate):
        if scope == 'gmres':
            assert r['IoFast'] == c['IoFast'] == r['EventFast'] == c['EventFast'] == r['AssemblyFast'] == c['AssemblyFast'] == 1
            assert r['GmresTuned']==0 and c['GmresTuned']==1 and r['TunedCalls']==0
            assert r['Assembly']==c['Assembly'] and r['Flow']==c['Flow']
            count=int(re.search(r'gmres_simd=(\d+)',c['Run']).group(1))
            assert c['TunedCalls']==count
            assert r['DecisionEventsSha256']==c['DecisionEventsSha256'] and c['DecisionEventsSha256']
            assert r['NonGmresSolverEventsRows']==c['NonGmresSolverEventsRows']>=0
        elif scope == 'assembly':
            assert r['IoFast'] == c['IoFast'] == r['EventFast'] == c['EventFast'] == 1
            assert r['AssemblyFast'] == 0 and c['AssemblyFast'] == 1
            assert r['Flow'] == c['Flow'] and c['Flow'] in ('INCONSISTENT','DAVIS')
            assert r['Assembly'][0] == 0
            if c['Flow'] == 'INCONSISTENT' and c['Mode'] == 'AUTO':
                assert c['Assembly'][0] > 0, 'Native assembly was not exercised'
            else:
                assert c['Assembly'][0] == 0
            assert r['Assembly'][1:] == c['Assembly'][1:]
        elif scope == 'events':
            assert r['IoFast'] == c['IoFast'] == 1
            assert r['EventFast'] == 0 and c['EventFast'] == 1
        else:
            assert scope == 'io' and r['IoFast'] == 0 and c['IoFast'] == 1
            assert r.get('EventFast', 0) == c.get('EventFast', 0) == 0
        assert all(x['StateCopy'] == x['GmresSimd'] == 1 and x['ExclusiveExcel'] is True for x in (r, c))
        keys=['Case','Variant','Mode','Precision','NativeDllSha256','Fos','Bracket','Disp']
        if scope!='gmres':keys+=['SolverEventsSha256','SolverEventsBytes']
        for key in keys:
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
                           solver_events_sha256=c['SolverEventsSha256'], logs_exact=r['SolverEventsSha256']==c['SolverEventsSha256'],
                           decision_events_equal=True))
    return report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('reference'); ap.add_argument('candidate'); ap.add_argument('--out', required=True)
    ap.add_argument('--scope', choices=('io', 'events', 'assembly', 'gmres'), default='io')
    args = ap.parse_args()
    result = dict(status='PASS', cases=compare(rows(args.reference), rows(args.candidate), args.scope))
    text = json.dumps(result, ensure_ascii=False, indent=2)
    Path(args.out).write_text(text + '\n', encoding='utf-8')
    print(text)


if __name__ == '__main__':
    main()
