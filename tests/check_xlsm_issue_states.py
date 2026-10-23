"""Check actual XLSM issue regression state/output evidence (no Excel automation)."""
import argparse
import json
import math
from pathlib import Path


def read(p):
    return json.loads(p.read_text(encoding='utf-8-sig'))


def maxdiff(a,b):
    assert len(a)==len(b), (len(a),len(b))
    return max((abs(x-y) for x,y in zip(a,b)),default=0)


def pressure_principals(folder):
    result=read(folder/'result.json')
    rows=read(folder/'水圧結果.json')[35:]
    stage=max(r[0] for r in rows if isinstance(r[0],(int,float)))
    final=[r for r in rows if r[0]==stage]
    stress=result['stress']
    assert len(final)*17==len(stress), 'Missing final GP pressures'
    error=0
    for r in final:
        level,gamma=(1,9.80665) if folder.name=='issue43_water_srm' else (2,10)
        assert abs(r[5]-gamma*max(level-r[4],0))<1e-8, ('Wrong accepted water profile',r[:6])
        assert str(r[14]).startswith('HYDROSTATIC:B:'), r[14]
        offset=((int(r[1])-1)*4+int(r[2])-1)*17
        sx,sy,tau=stress[offset:offset+3]
        sz=stress[offset+16]
        center=(sx+sy)/2
        radius=math.hypot((sx-sy)/2,tau)
        values=[center+radius-r[5],center-radius-r[5],sz-r[5]]
        hi,lo=max(values),min(values)
        expected=hi if lo>0 else hi-lo if hi>0 else -lo
        error=max(error,abs(stress[offset+8]-expected))
    assert error<1e-8, ('Effective stress diagnostic stale',folder,error)
    return {'gp_count':len(final),'max_effective_principal_error':error,'stage':stage}


def check(root):
    cases=root/'cases'
    report={'status':'PASS','pressure':{},'resets':{},'views':{}}
    for name in ('issue43_water_ramp_cache','issue43_water_srm'):
        report['pressure'][name]=pressure_principals(cases/name)
    contract=read(cases/'issue43_water_srm'/'resume_residual_contract.json')
    assert contract['status']=='PASS' and contract['production_tolerance_changed'] is False
    assert max(contract['fresh'],contract['resumed'])<=contract['limit'] and contract['delta']<=1e-9
    report['natural_srm_residual']=contract
    for name in ('issue46_reordered','issue46_multiple','issue46_final_reset','issue46_no_reset','issue46_final_output'):
        a,b=[read(cases/name/(n+'.json')) for n in ('result','resume')]
        assert a['status'].startswith('PASS|') and b['status'].startswith('PASS|')
        du=maxdiff(a['disp'],b['disp'])
        assert du<1e-9
        assert max(a['metrics'][3],b['metrics'][3])<1e-7
        assert abs(a['metrics'][2]-b['metrics'][2])<1e-6
        if name=='issue46_final_reset':
            assert max(map(abs,a['stress']))==0 and max(map(abs,a['disp']))==0
            assert a['metrics'][2]==0 and a['metrics'][3]==0
        report['resets'][name]={'max_du':du,'fresh_reaction':a['metrics'][2],'resumed_reaction':b['metrics'][2],'fresh_residual':a['metrics'][3],'resumed_residual':b['metrics'][3]}
        view=read(root/'view'/name/'issue46.json')
        assert view['status']=='PASS' and view['plot']=='PASS' and view['reopen']=='PASS'
        assert view['coordinate_guard'].startswith('ERROR|')
        if name=='issue46_no_reset': assert view['saved_coordinate_guard'].startswith('ERROR|')
        report['views'][name]=view
    rows=read(cases/'issue26_hold'/'結果増分.json')[1:]
    groups={stage:[r for r in rows if r[0]==stage] for stage in (1,2,3,4)}
    first=max((abs(v) for st in (1,3) for r in groups[st] for v in r[8:12]),default=0)
    assert first>1e-10, 'No plastic increment was exercised'
    hold={st:max((abs(v) for r in groups[st] for v in r[8:12]),default=0) for st in (2,4)}
    assert max(hold.values())<1e-12
    report['hold']={'plastic_increment_max':first,'unchanged_stage_row_counts':{st:len(groups[st]) for st in (2,4)},'unchanged_stage_max_plastic':hold,'note':'Unchanged stages may have no accepted increment; no prior increment may be relabelled as that stage.'}
    rerun=read(root/'rerun'/'rerun.json')
    clean=read(cases/'issue38_elastic_control'/'result.json')
    assert rerun['old_max_plastic']>1e-10 and rerun['new_max_plastic']<=1e-12
    assert rerun['rows']==rerun['unique_keys'] and rerun['rows']>0
    errors={key:maxdiff(rerun[key],clean[key]) for key in ('disp','stress','metrics')}
    assert errors['disp']<1e-9 and errors['stress']<1e-6 and errors['metrics']<1e-6, errors
    report['rerun']={'old_plastic':rerun['old_max_plastic'],'new_plastic':rerun['new_max_plastic'],'rows':rerun['rows'],'clean_control_max_differences':errors}
    guard=read(root/'guard'/'guard.json')
    assert guard['status']=='PASS' and guard['settings_unchanged'] and guard['reopen']
    report['formula_guard']=guard
    report['snapshots']={}
    for name in ('snapshot_hydro','snapshot_steady'):
        path=root/name
        record=read(path/'snapshot.json')
        assert record['status']=='PASS'
        assert (path/'accepted.csv').read_bytes()==(path/'restored.csv').read_bytes()
        assert (path/'accepted.csv').read_bytes()!=(path/'last_trial.csv').read_bytes()
        report['snapshots'][name]=record
    return report


if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--root',required=True);ap.add_argument('--out',required=True);ap.add_argument('--pressure-only',action='store_true');args=ap.parse_args()
    root=Path(args.root)
    result=pressure_principals(root) if args.pressure_only else check(root)
    Path(args.out).write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print('PASS')
