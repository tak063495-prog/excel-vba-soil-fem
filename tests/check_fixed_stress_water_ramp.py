"""Independent global vertical load check for water_external_linear.json."""
import argparse
import csv
import json
from pathlib import Path


def check(run):
    frames = []
    for file in run.rglob('fixed_stress.csv'):
        if not file.parent.name.startswith('C'):
            continue
        lines = file.read_text(encoding='utf-8-sig').splitlines()
        state = next(line.split(',') for line in lines if line.startswith('STATE,'))
        phase = next(line.split(',') for line in lines if line.startswith('PHASE,'))
        if int(phase[1]) != 3:
            continue
        fraction = float(phase[2])/40.0
        load_y = sum(float(row['locked_sw']) for row in csv.DictReader((file.parent/'dof.csv').open(encoding='utf-8-sig')) if int(row['dof']) % 2)
        expected = -(40+9.81*(2+fraction))
        if abs(load_y-expected) > 1e-8:
            raise ValueError(f'water load jump at t={state[1]}: {load_y} != {expected}')
        frames.append({'frame': file.parent.name, 'time_s': float(state[1]), 'fraction': fraction,
                       'load_y_kN': load_y, 'expected_y_kN': expected, 'error_kN': abs(load_y-expected)})
    if len(frames) < 4 or not any(0 < row['fraction'] < 1 for row in frames):
        raise ValueError('no intermediate ramp fields')
    result = json.loads((run/'result.json').read_text(encoding='utf-8-sig'))
    reaction_error = abs(result['metrics'][2]-69.43)
    if reaction_error > 1e-6:
        raise ValueError(f'final reaction error {reaction_error}')
    return {'status': 'PASS', 'frames': sorted(frames, key=lambda r:r['time_s']),
            'final_reaction_error_kN': reaction_error,
            'basis': 'soil volume 2m3 * 20kN/m3 + top water pressure 9.81*(level-2m)*1m2'}


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--run', type=Path, required=True)
    ap.add_argument('--report', type=Path, required=True)
    args = ap.parse_args()
    result = check(args.run)
    args.report.write_text(json.dumps(result, indent=2), encoding='utf-8')
    print(json.dumps(result, indent=2))
