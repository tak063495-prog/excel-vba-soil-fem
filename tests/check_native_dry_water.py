"""Compare hybrid zero-pressure evaluation with VBA through wet/dry transitions."""
from pathlib import Path
import argparse, json, re, subprocess
import numpy as np

ap = argparse.ArgumentParser()
ap.add_argument('--book', type=Path, required=True)
ap.add_argument('--out', type=Path, required=True)
a = ap.parse_args()
root = Path(__file__).resolve().parents[1]
a.out = a.out.resolve()
a.out.mkdir(parents=True, exist_ok=False)
# Each invocation owns DispatchEx/New-Object Excel.Application, not an active
# user instance. Other independent validations may continue at the same time.
runner_text = (root/'tests/run_groundwater_fixture.ps1').read_text(encoding='utf-8-sig')
runner_text = runner_text.replace("if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; do not dispatch.'}", '# Isolated owned Excel instance.')
runner_text = runner_text.replace('$PSScriptRoot/', str(root/'tests').replace('\\', '/')+'/')
runner = a.out/'owned_fixture.ps1'
runner.write_bytes(runner_text.encode('utf-8-sig'))
records = []
for case in ('mixed_t6/mixed_water', 'linear_solver/gmres_water_plastic'):
    original = json.loads((root/'tests/fixtures'/f'{case}.json').read_text(encoding='utf-8-sig'))
    if case.startswith('mixed_t6'):
        # Q8 then two T6 elements: dry, partially wet and wet runs, followed
        # by rising and falling water across the dry-kernel eligibility test.
        original['water']['profiles'] = [['W', -1, -.2, ''], ['W', 3, .5, ''], ['B', -1, .7, ''], ['B', 3, .7, ''], ['C', -1, -.2, ''], ['C', 3, -.2, '']]
        original['stages'] += [[2, 'WATER', 'B', 1, 1, 'rise'], [3, 'WATER', 'C', 1, 1, 'fall']]
    outputs = {}
    for label, native, enabled in (('baseline', 'AUTO', 0), ('hybrid', 'AUTO', 1), ('vba', 'OFF', 1)):
        data = json.loads(json.dumps(original))
        data['settings']['NATIVE_DRY_WATER'] = enabled
        data['settings']['NATIVE_WATER_TANGENT'] = 1
        data['name'] = case.split('/')[-1]+'_'+label
        fixture = a.out/(data['name']+'.json')
        fixture.write_text(json.dumps(data), encoding='utf-8')
        dest = a.out/data['name']
        with (a.out/(data['name']+'.log')).open('w', encoding='utf-8') as stream:
            result = subprocess.run(['pwsh', '-NoProfile', '-File', str(runner), '-Book', str(a.book.resolve()), '-Fixture', str(fixture), '-Out', str(dest), '-Native', native, '-Resume'], cwd=root, stdout=stream, stderr=subprocess.STDOUT)
        assert result.returncode == 0, data['name']
        value = json.loads((dest/'result.json').read_text(encoding='utf-8-sig'))
        resume = json.loads((dest/'resume.json').read_text(encoding='utf-8-sig'))
        assert value['status'].startswith('PASS|') and resume['status'].startswith('PASS|')
        count = re.search(r'dry_water_elements=(\d+)', value['native'])
        count = int(count[1]) if count else 0
        assert (count > 0) == (label == 'hybrid'), (data['name'], value['native'])
        outputs[label] = value
        print('PASS', data['name'], 'dry native elements', count, flush=True)
    baseline = outputs['vba']
    for label in ('baseline', 'hybrid'):
        value = outputs[label]
        du = float(np.max(abs(np.array(value['disp'])-baseline['disp'])))
        # Principal direction and yield diagnostic sentinels are not stresses.
        stress = np.array(value['stress']).reshape(-1, 17)
        reference = np.array(baseline['stress']).reshape(-1, 17)
        ds = float(np.max(abs(stress[:, [0, 1, 2, 16]]-reference[:, [0, 1, 2, 16]])))
        assert du < 1e-7 and ds < 1e-4, (case, label, du, ds)
        records.append(dict(case=case, mode=label, max_displacement_difference=du, max_stress_difference=ds, status='PASS'))
    (a.out/'results.json').write_text(json.dumps(records, indent=2), encoding='utf-8')
print('PASS hybrid/VBA transitions and checkpoint resumes', flush=True)
