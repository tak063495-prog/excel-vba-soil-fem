"""Small real Excel fixtures for checkpoint regressions."""
import json
from pathlib import Path
from copy import deepcopy
root=Path(__file__).parent/'fixtures'
out=root/'issue62_67';out.mkdir(exist_ok=True)
base=json.loads((root/'issues_xlsm/issue45_equivalent.json').read_text(encoding='utf-8'))
dry=deepcopy(base);dry['name']='dry_fixed_fs'
dry['water']={'mode':'NONE','gamma':10,'defaultProfile':'','auto':'OFF','profiles':[],'materials':[],'boundaries':[]}
dry['stages'].append([2,'SRM','',1,1,'fixed Fs'])
dry['settings']={'SRM_FIXED_FS':1.2,'SRM_FMAX':3,'FLOW_POLICY':'DAVIS'}
(out/'dry_fixed_fs.json').write_text(json.dumps(dry,indent=2),encoding='utf-8')
wet=deepcopy(base);wet['name']='minimum_step'
wet['water']['boundaries']=[]
wet['water']['profiles']=wet['water']['profiles'][:2]
wet['water']['coupledMaterials']=[[1,1,1e-6,1e-8,1e-8,'',.35]]
wet['stages'].append([2,'WAIT','',1,1,'three seconds'])
wet['water']['stagePlans']=[[2,'FIXED_STRESS','',3,.75,'',100,1e-8,'']]
wet['water']['timeOptions']=[[2,'S','HOLD','INSTANT',.5,0,0,'','','']]
for n in wet['nodes']:n[3]=n[4]=1
(out/'minimum_step.json').write_text(json.dumps(wet,indent=2),encoding='utf-8')
joint=json.loads((root/'issue49_57/noflow_joint_continuous.json').read_text(encoding='utf-8'))
joint.pop('expectedAnalysisError',None);joint['name']='joint_pressure_split'
joint.setdefault('settings',{})['RCM_POLICY']='OFF'
joint['water']['mode']='HYDROSTATIC';joint['water']['boundaries']=[];joint['water']['drainSets']=[]
joint['water']['stagePlans']=[[2,'FIXED_STRESS','',1,1,'',100,1e-8,'']]
for n in joint['nodes']:n[3]=n[4]=1
(out/'joint_pressure_split.json').write_text(json.dumps(joint,indent=2),encoding='utf-8')
print('Created three checkpoint regression fixtures')
