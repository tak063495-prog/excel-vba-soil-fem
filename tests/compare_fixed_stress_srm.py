"""Check frozen-pressure SRM output and unchanged construction continuation."""
import argparse
import csv
import json
from pathlib import Path


def water(folder):
    rows = (folder/'water.csv').read_text(encoding='utf-8-sig').splitlines()
    first = rows.index('element,gauss,pressure_kPa')+1
    return {(int(a), int(b)): float(c) for line in rows[first:] if line.strip()
            for a, b, c in [line.split(',')]}


def compare(control, srm):
    a = json.loads((control/'result.json').read_text(encoding='utf-8-sig'))
    b = json.loads((srm/'result.json').read_text(encoding='utf-8-sig'))
    if not all(v['status'].startswith('PASS|') for v in (a, b)):
        raise ValueError('both runs must pass')
    if len(a['disp']) != len(b['disp']) or len(a['stress']) != len(b['stress']):
        raise ValueError('field sizes differ')
    du = max(abs(x-y) for x, y in zip(a['disp'], b['disp']))
    ds = max(abs(x-y) for x, y in zip(a['stress'], b['stress']))
    if du > 1e-9 or ds > 1e-6:
        raise ValueError(f'SRM changed construction continuation u={du} stress={ds}')
    history = json.loads((srm/'時間履歴.json').read_text(encoding='utf-8-sig'))
    frames = {row[0]: row for row in history[1:] if row and row[0] and row[10]}
    branch_checks = []
    for row in frames.values():
        if row[9] != 'SRM_BRANCH':
            continue
        parent = frames[row[11]]
        if row[2] != parent[2]:
            raise ValueError('SRM advanced physical time')
        p, q = water(srm/Path(parent[10])), water(srm/Path(row[10]))
        if p.keys() != q.keys():
            raise ValueError('pressure fields differ in size')
        dp = max(abs(p[k]-q[k]) for k in p)
        if dp > 1e-10:
            raise ValueError(f'SRM pressure changed {dp}')
        branch_checks.append({'state': row[0], 'parent': parent[0], 'time_s': row[2],
                              'Fs': row[12], 'max_pressure_difference_kPa': dp})
    if not branch_checks:
        raise ValueError('no saved SRM branch')
    return {'status': 'PASS', 'all_dofs': len(a['disp']),
            'all_stress_slots': len(a['stress']), 'max_u_difference_m': du,
            'max_stress_difference_kPa': ds, 'saved_branch_checks': branch_checks,
            'scope': 'accepted saved branches and construction after interim SRM; failed trial fields are not independently exported'}


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    for name in ('control', 'srm', 'report'):
        ap.add_argument('--'+name, type=Path, required=True)
    args = ap.parse_args()
    record = compare(args.control, args.srm)
    args.report.write_text(json.dumps(record, indent=2), encoding='utf-8')
    print(json.dumps(record, indent=2))
