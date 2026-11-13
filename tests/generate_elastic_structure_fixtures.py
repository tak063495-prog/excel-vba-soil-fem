"""Elastic law controls: compare against a strictly elastic MC stress path."""
import json
from copy import deepcopy
from pathlib import Path
root=Path(__file__).parent/'fixtures'
base=json.loads((root/'issue62_67/dry_fixed_fs.json').read_text())
out=root/'elastic_structure';out.mkdir(exist_ok=True)
for flow in ('DAVIS','INCONSISTENT'):
    for law,native in (('reference','AUTO'),('elastic','AUTO'),('elastic','OFF')):
        case=deepcopy(base)
        case['name']=f'{flow.lower()}_{law}_{native.lower()}'
        material=case['materials'][0]
        material[10]='STRUCT' if law=='reference' else 'STRUCT_ELASTIC'
        material[8]=0 if law=='reference' else 1 # elastic kind overrides this
        material[6]=1e6 if law=='reference' else 0
        case['settings'].update(FLOW_POLICY=flow,NATIVE_DLL=native,SRM_FIXED_FS=10,SRM_FMAX=12)
        (out/(case['name']+'.json')).write_text(json.dumps(case,indent=2),encoding='utf-8')
print('Created six elastic structural controls')
