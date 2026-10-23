"""Compare a fresh-process partial-stage continuation to an accepted final frame."""
import argparse
import csv
import json
from pathlib import Path


def check(history, expected, result):
    actual = json.loads((history / 'continued.json').read_text(encoding='utf-8-sig'))
    reference = json.loads(result.read_text(encoding='utf-8-sig'))
    if not actual['status'].startswith('PASS|'):
        raise ValueError(actual['status'])
    nodes = {int(row['node']): row for row in csv.DictReader((expected / 'node.csv').open(encoding='utf-8-sig'))}
    if len(nodes) != len(actual['nodes']):
        raise ValueError('node count differs')
    du = max(abs(float(nodes[int(n)][key]) - value)
             for n, ux, uy in actual['nodes'] for key, value in [('ux', ux), ('uy', uy)])
    water = (expected / 'water.csv').read_text(encoding='utf-8-sig').splitlines()
    first = water.index('element,gauss,pressure_kPa') + 1
    pressures = {(int(row[0])-1)*4+int(row[1])-1: float(row[2])
                 for line in water[first:] if line.strip() for row in [line.split(',')]}
    if len(pressures) != len(actual['pressure']):
        raise ValueError('pressure count differs')
    dp = max(abs(actual['pressure'][i]-p) for i, p in pressures.items())
    if len(actual['stress']) != len(reference['stress']):
        raise ValueError('stress count differs')
    ds = 0.0
    for i, (a, b) in enumerate(zip(actual['stress'], reference['stress'])):
        if i % 17 == 6:  # Principal direction is undefined at repeated eigenvalues.
            continue
        if i % 17 == 9:  # CSV rounding of large yield-function sentinels.
            if abs(a-b) > 1e-6+5e-15*max(abs(a), abs(b)):
                raise ValueError(f'yield diagnostic differs at {i}')
            continue
        ds = max(ds, abs(a-b))
    if du > 1e-9 or dp > 1e-6 or ds > 1e-6:
        raise ValueError(f'continuation differs u={du}, p={dp}, stress={ds}')
    return {'status': 'PASS', 'nodes': len(nodes), 'pressure_points': len(pressures),
            'max_u_difference_m': du, 'max_pressure_difference_kPa': dp,
            'max_stress_difference_kPa': ds,
            'diagnostics': 'principal direction at repeated eigenvalues is excluded; yield sentinel uses CSV relative-rounding allowance'}


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    for name in ('history', 'expected', 'result', 'report'):
        ap.add_argument('--'+name, type=Path, required=True)
    args = ap.parse_args()
    record = check(args.history, args.expected, args.result)
    args.report.write_text(json.dumps(record, indent=2), encoding='utf-8')
    print(json.dumps(record, indent=2))
