"""Compare accepted-state CSVs across releases, including all saved time steps."""
import argparse
import hashlib
import json
from pathlib import Path

FILES = {'node.csv', 'dof.csv', 'gauss.csv', 'water.csv', 'increment.csv',
         'plastic_cumulative.csv', 'fixed_stress.csv', 'fixed_stress_volume.csv'}


def states(root):
    result = {}
    for p in root.rglob('*.csv'):
        if p.name not in FILES:
            continue
        parts = tuple('time_ROOT' if part.startswith('time_20') else part
                      for part in p.relative_to(root).parts)
        key = '/'.join(parts)
        assert key not in result, key
        result[key] = hashlib.sha256(p.read_bytes()).hexdigest()
    return result


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    for key in ('reference', 'candidate', 'out'):
        ap.add_argument('--' + key, type=Path, required=True)
    a = ap.parse_args()
    old, new = states(a.reference), states(a.candidate)
    assert old and old.keys() == new.keys(), (len(old), len(new), old.keys() ^ new.keys())
    different = [k for k in old if old[k] != new[k]]
    assert not different, different
    result = {'status': 'PASS', 'csv_files': len(new), 'files': new,
              'scope': 'byte-identical accepted-state CSVs: nodes, DOF loads/boundary, stresses, GP water pressure/Fs, increments/plastic state, fixed-stress state/ledger; stage and time outputs, initial and resumed'}
    a.out.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    print('PASS:', len(new), 'accepted-state CSV files')
