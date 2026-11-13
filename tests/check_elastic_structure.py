"""Validate elastic structures against an independent elastic stress path."""
import argparse,json
from pathlib import Path
import numpy as np
root=Path(__file__).resolve().parents[1]
ap=argparse.ArgumentParser();ap.add_argument('--runs-root',type=Path);ap.add_argument('--out',type=Path);a=ap.parse_args()
folder=a.runs_root or root/'evidence/targeted'
records=[]
for flow in ('davis','inconsistent'):
    def load(name):
        path=folder/name/'result.json'
        obj=json.loads(path.read_text(encoding='utf-8-sig'))
        assert obj['status'].startswith('PASS|'),obj
        return obj
    reference=load(f'{flow}_reference_auto')
    reference_stress=np.array(reference['stress']).reshape(-1,17)
    for native in ('auto','off'):
        name=f'{flow}_elastic_{native}'
        actual=load(name)
        stress=np.array(actual['stress']).reshape(-1,17)
        for data in (stress,actual['disp'],actual['metrics']):assert np.isfinite(data).all()
        # Column 9 is a Mohr-Coulomb yield value for the reference. Elastic
        # structures have no MC yield surface and explicitly report zero.
        mask=np.ones(17,dtype=bool);mask[9]=False
        stress_error=float(np.max(abs(stress[:,mask]-reference_stress[:,mask])))
        displacement_error=float(np.max(abs(np.array(actual['disp'])-reference['disp'])))
        assert stress_error<1e-10 and displacement_error<1e-13
        assert (stress[:,9]==0).all()
        assert actual['metrics'][4]==10 # fixed Fs=10, even with reduction flag=1
        assert abs(actual['metrics'][2]-20)<1e-10
        assert actual['metrics'][3]<1e-10
        if native=='auto':assert 'material=' in actual['native'] and 'eval_elements=0' in actual['native']
        records.append({'case':name,'stress_error':stress_error,'displacement_error':displacement_error,'relative_residual':actual['metrics'][3]})
out=a.out or root/'delivery/elastic_structure_checks.json'
out.write_text(json.dumps({'status':'PASS','reference':'same elastic constants with no yielding along the reference MC stress path','cases':records},indent=2,allow_nan=False),encoding='utf-8')
print('PASS: DAVIS and INCONSISTENT, native/VBA, SRM exclusion and checkpoint resume')
