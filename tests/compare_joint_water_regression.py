"""Aggregate JOINT checks and compare the unchanged groundwater cases to WATER_01."""
from pathlib import Path
import argparse
import json
import hashlib
from check_joint_water import run_check

def load(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))

def main():
    p=argparse.ArgumentParser()
    p.add_argument('--joint',required=True)
    p.add_argument('--legacy',required=True)
    p.add_argument('--previous',required=True)
    p.add_argument('--book',required=True)
    p.add_argument('--out',required=True)
    a=p.parse_args()
    joint,legacy,previous=map(Path,[a.joint,a.legacy,a.previous])
    fixtures=Path(__file__).parent/'fixtures'
    digest=hashlib.sha256(Path(a.book).read_bytes()).hexdigest().upper()
    cases={}
    for case in ['dry','wet','prescribed','continuous','prescribed_steady','soil_steady','partial','slip','opening','stages','srm','struct_pair','wet_davis','slip_davis']:
        fixture=fixtures/('joint_water_'+case.removesuffix('_davis')+'.json')
        r=load(joint/case/'result.json')
        assert r['book_sha256']==digest,(case,r['book_sha256'])
        assert r['fixture_sha256']==hashlib.sha256(fixture.read_bytes()).hexdigest().upper(),case
        cases[case]=run_check(fixture,joint/case)
        (joint/case/'check.json').write_text(json.dumps(cases[case],ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    rejected={}
    for case in ['missing_mode','inconsistent_pressure','conflicting_heads','wrong_soil_mode']:
        r=load(joint/case/'input_check.json')
        assert r['status']=='PASS' and r['expected_error'] in r['actual_error'] and r['book_sha256']==digest
        rejected[case]=r
    dry=load(joint/'dry_comparison.json')
    assert dry['status']=='PASS' and all(x==0 for x in dry['max_difference'].values())
    comparison=load(legacy/'comparison.json')
    assert isinstance(comparison,list) and any(r.get('case')=='wet_srm' for r in comparison)
    assert all(r['max_du']==0 for r in comparison if 'flow' in r)
    versions={}
    for case in ['box','cutoff','stages','box_inside','separated','initial_birth','all_birth','wet_srm']:
        old,new=load(previous/case/'result.json'),load(legacy/case/'result.json')
        assert old['fixture_sha256']==new['fixture_sha256'],case
        delta={}
        for key in ['disp','stress','metrics']:
            assert len(old[key])==len(new[key]),(case,key)
            delta[key]=max(abs(x-y) for x,y in zip(old[key],new[key]))
        # These cases have no JOINT and should retain the exact numerical path.
        assert all(v==0 for v in delta.values()),(case,delta)
        versions[case]={'max_difference':delta,'previous_seconds':old['seconds'],'current_seconds':new['seconds'],'dofs':len(new['disp']),'Fs':new['metrics'][4:],'residual':new['metrics'][3]}
    report={'status':'PASS','book_sha256':digest,'joint':cases,'rejected':rejected,'dry_joint_comparison':dry,'unchanged_groundwater':versions,'legacy_comparison':comparison}
    Path(a.out).write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'status':'PASS','positive_cases':len(cases),'rejected_cases':len(rejected),'unchanged_groundwater_cases':len(versions),'book_sha256':digest}))

if __name__=='__main__':
    main()
