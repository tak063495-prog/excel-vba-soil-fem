"""Compare a fresh-process partial-stage continuation to an accepted final frame."""
import argparse
import csv
import json
import math
from pathlib import Path


def finite(value, location):
    value = float(value)
    if not math.isfinite(value):
        raise ValueError(f'non-finite value at {location}: {value}')
    return value


def validate_json(value, location):
    if isinstance(value, dict):
        for k, v in value.items(): validate_json(v, f'{location}.{k}')
    elif isinstance(value, list):
        for i, v in enumerate(value): validate_json(v, f'{location}[{i}]')
    elif isinstance(value, (int, float)):
        finite(value, location)


def check(history, expected, result):
    actual = json.loads((history / 'continued.json').read_text(encoding='utf-8-sig'))
    reference = json.loads(result.read_text(encoding='utf-8-sig'))
    validate_json(actual, str(history / 'continued.json'))
    validate_json(reference, str(result))
    if not actual['status'].startswith('PASS|'):
        raise ValueError(actual['status'])
    with (expected / 'node.csv').open(encoding='utf-8-sig') as source:
        nodes = {int(row['node']): row for row in csv.DictReader(source)}
    if len(nodes) != len(actual['nodes']):
        raise ValueError('node count differs')
    du = max(abs(finite(nodes[int(n)][key], f'node.csv node={n} {key}') - value)
             for n, ux, uy in actual['nodes'] for key, value in [('ux', ux), ('uy', uy)])
    water = (expected / 'water.csv').read_text(encoding='utf-8-sig').splitlines()
    first = water.index('element,gauss,pressure_kPa') + 1
    pressures = {(int(row[0])-1)*4+int(row[1])-1: finite(row[2], f'water.csv element={row[0]} gauss={row[1]}')
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
    for name, value in [('du', du), ('dp', dp), ('ds', ds)]: finite(value, name)
    return {'status': 'PASS', 'nodes': len(nodes), 'pressure_points': len(pressures),
            'max_u_difference_m': du, 'max_pressure_difference_kPa': dp,
            'max_stress_difference_kPa': ds,
            'diagnostics': 'principal direction at repeated eigenvalues is excluded; yield sentinel uses CSV relative-rounding allowance'}


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    for name in ('history', 'expected', 'result', 'report'):
        ap.add_argument('--'+name, type=Path, required=True)
    args = ap.parse_args()
    try:
        record = check(args.history, args.expected, args.result)
    except (ValueError, KeyError, TypeError, OverflowError) as error:
        args.report.write_text(json.dumps({'status':'FAIL', 'error':str(error)}, allow_nan=False), encoding='utf-8')
        raise
    args.report.write_text(json.dumps(record, indent=2, allow_nan=False), encoding='utf-8')
    print(json.dumps(record, indent=2, allow_nan=False))
