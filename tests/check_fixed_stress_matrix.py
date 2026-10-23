import json
import sys
from pathlib import Path
import numpy as np
from test_fixed_stress_reference import reference_local

actual=json.loads(Path(sys.argv[1]).read_text(encoding='utf-8-sig'))
checks=[]
for kind in ['q8','t6']:
    xy=np.array([[0,0],[1,0],[1,1],[0,1],[.5,0],[1,.5],[.5,1],[0,.5]],float)
    if kind=='t6': xy=np.array([[0,0],[1,0],[0,1],[0,1],[.5,0],[.5,.5],[0,1],[0,.5]])
    coords=xy@np.array([[2,.2],[.3,1.5]])+np.array([1,-2])
    blocks=reference_local(coords,kind,.8,2e-6,3e-8,1e-8,9.80665,12000,.27,1.7)
    expected=np.concatenate([b.ravel() for b in blocks])
    values=np.array(actual[kind])
    offset=0
    for block in blocks:
        current=values[offset:offset+block.size].reshape(block.shape)
        assert np.max(abs(current-block)) <= 2e-12*np.max(abs(block)),(kind,np.max(abs(current-block)))
        offset+=block.size
    area=2.94*(.5 if kind=='t6' else 1.)
    mass,flow,stab,Q,gravity=blocks
    assert abs(mass.sum()-2e-6*area*1.7)<1e-19
    p=9.80665*(4-coords[:4,1]); p[3]=0 if kind=='t6' else p[3]
    assert np.max(abs(flow@p+gravity))<1e-21
    displacement=np.zeros(16);displacement[0::2]=.002*coords[:,0];displacement[1::2]=-.003*coords[:,1]
    assert abs((Q.T@displacement).sum()-.8*(-.001)*area*1.7)<1e-15
    checks.append(dict(kind=kind,max_abs=float(np.max(abs(values-expected))),mass_volume=float(mass.sum()/2e-6)))
matrix=np.array([[2.,-1,0],[-1,5,-1],[0,-1,3]])
for case in ['closed','dirichlet']:
    expected=np.linalg.solve(matrix,np.array([1.,2.,3.]))
    if case=='dirichlet': expected=np.r_[.7,np.linalg.solve(matrix[1:,1:],np.array([2.,3.])-matrix[1:,0]*.7)]
    values=np.array(actual[case]);assert values[9]==1
    assert np.max(abs(values[:3]-expected))<1e-12
    assert np.max(abs(values[3:6]))<1e-12
    residual=matrix@expected-np.array([1.,2.,3.])
    assert abs(values[6]-residual[0])<1e-12
    checks.append(dict(case=case,max_abs=float(np.max(abs(values[:3]-expected)))))
print(json.dumps(dict(status='PASS',checks=checks),indent=2))
