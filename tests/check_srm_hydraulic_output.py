"""Verify frozen hydraulic output in the saturated_slope_srm fixture."""
import argparse
import json
from pathlib import Path


def check(folder):
    rows=json.loads((folder/'水圧結果.json').read_text(encoding='utf-8-sig'))[35:]
    parent={(int(r[1]),int(r[2])):r for r in rows if r[0]==2}
    branch={(int(r[1]),int(r[2])):r for r in rows if r[0]==3}
    assert len(parent)==len(branch)==80 and parent.keys()==branch.keys()
    delta=0
    for key,r in branch.items():
        assert all(isinstance(r[i],(int,float)) for i in (5,12,13)), ('Missing SRM pressure/flux',key)
        delta=max(delta,*(abs(r[i]-parent[key][i]) for i in (5,12,13)))
    assert delta==0, ('Frozen parent field changed in SRM',delta)
    return {'status':'PASS','parent_stage':2,'srm_stage':3,'gauss_points':80,
            'nonblank_srm_flux_points':80,'pressure_flux_max_difference':delta}


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--run',type=Path,required=True);p.add_argument('--report',type=Path,required=True)
    a=p.parse_args();result=check(a.run)
    a.report.write_text(json.dumps(result,indent=2),encoding='utf-8')
    print('PASS: 80 SRM hydraulic output points')
