"""Compare generated-mesh elastic runs and unchanged mesh routes to accepted records."""
from pathlib import Path
import json,math,argparse
import numpy as np

def load(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('folder');a=p.parse_args();r=Path(a.folder);old=Path(__file__).resolve().parent/'tmp/mesh_light'
    assert load(r/'package.json')['status']==load(r/'final_sources.json')['status']=='PASS'
    assert load(r/'final_mesh/independent.json')['status']=='PASS'
    results=[]
    for name in ['step_wall_fill','box_tunnel','octagon_tunnel','gravity_taper_blocks','lwall_taper_blocks','circle64_tunnel']:
        current=f'final_{name}_solve'
        x=load(r/f'reference_{name}_solve/result.json');y=load(r/current/'result.json');fixture=load(r/f'{name}_elastic.json')
        assert x['status'].startswith('PASS|') and y['status'].startswith('PASS|')
        delta={}
        for key in ['disp','stress','metrics']:
            assert len(x[key])==len(y[key]);delta[key]=max(abs(i-j) for i,j in zip(x[key],y[key]));assert delta[key]==0,(name,key,delta[key])
        geo=next(c for c in load(r/'final_mesh/independent.json')['cases'] if c['case']==name)
        gamma={str(m[0]):m[4]*m[3] for m in fixture['materials']}
        weight=sum(geo['areas'][mat]*gamma[mat] for mat in geo['areas'])
        assert math.isclose(y['metrics'][0],weight,abs_tol=1e-7)
        assert abs(y['metrics'][2]-weight)<1e-6 and y['metrics'][3]<1e-7
        results.append(dict(case=name,all_dofs=len(y['disp']),stress_values=len(y['stress']),max_difference=delta,weight_kN=weight,reaction_y_kN=y['metrics'][2],relative_residual=y['metrics'][3],Fs_diagnostics=y['metrics'][4:],reference_seconds=x['seconds'],candidate_seconds=y['seconds']))
    simple=[]
    for case in ['slope2','slope1','flat1','slope2_quality']:
        x=load(old/f'release_benchmark/{case}_0.json');y=load(r/f'simple_regression/{case}_0.json')
        assert x['nodes']==y['nodes'] and x['elements']==y['elements'] and x['result']==y['result'],case
        simple.append(dict(case=case,all_node_element_cells_exact=True,seconds=y['seconds']))
    x=load(old/'baseline_extended2/slope2_delaunay_0.json');y=load(r/'guard_regression/slope2_delaunay_0.json')
    assert x['nodes']==y['nodes'] and x['elements']==y['elements'] and x['result']==y['result']
    guards=[x for x in load(r/'guard_regression/summary.json') if x.get('status')=='EXPECTED_REJECT']
    assert len(guards)==6 and all(x['mesh_unchanged'] for x in guards)
    assert load(r/'wet_joint/check.json')['status']==load(r/'wet_joint/compare_previous.json')['status']=='PASS'
    report=dict(status='PASS',elastic=results,simple_routes=simple,policy1_exact=True,protected_input_rejections=6,water_joint=load(r/'wet_joint/check.json'),water_joint_previous=load(r/'wet_joint/compare_previous.json'),limitations='Elastic self-weight and existing wet-JOINT/resume regression; not a new full SRM search, construction excavation validation or mesh-convergence study.')
    (r/'verification.json').write_bytes((json.dumps(report,indent=2)+'\n').encode());print(json.dumps(report,indent=2))
