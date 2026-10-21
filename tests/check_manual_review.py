"""Reconcile review exercises against the archived original and unchanged inputs."""
from pathlib import Path
import json,gzip,hashlib,sys
ROOT=Path.cwd();out=Path(sys.argv[1]).resolve() if len(sys.argv)>1 else ROOT/'tests/tmp/manual_review'
def source(p):return p if p.exists() else p.with_suffix(p.suffix+'.gz')
def read(p):
    p=source(p);return json.loads(gzip.decompress(p.read_bytes()) if p.suffix=='.gz' else p.read_text(encoding='utf-8-sig'))
pairs=[(ROOT/'docs/manuals/construction-evidence/20261021/road01/result.json.gz',out/'road_ui05/result.json'),(out/'road_gamma20/result.json',out/'road_gamma20_ui05/result.json')]
record=[]
for a,b in pairs:
    a,b=source(a),source(b);u,v=read(a),read(b);assert u['status']==v['status']=='PASS'
    for name in ['材料データ','節点データ','要素データ','ステージ','載荷','結果節点','結果要素','ステージ結果']:
        assert u['tables'][name]==v['tables'][name],(b,name)
    record.append(dict(baseline=str(a.relative_to(ROOT)),candidate=str(b.relative_to(ROOT)),baseline_sha256=hashlib.sha256(a.read_bytes()).hexdigest(),candidate_sha256=hashlib.sha256(b.read_bytes()).hexdigest(),matched_tables=['materials','nodes','elements','stages','loads','all stage nodal displacement/force/reaction','all stage GP stresses/strains','stage results'],seconds_before=u['seconds'],seconds_after=v['seconds']))
base=read(ROOT/'tests/manual_cases/road_widening_shape.json');change=read(out/'road_gamma20.json')
expected=json.loads(json.dumps(base));expected['materials'][7][4]=20;assert expected==change
balance=read(out/'road_gamma20_ui05/balance_checks.json');first,last=balance['stage_balance'][0],balance['stage_balance'][-1]
assert abs(first['reaction_y_kN']-1440)<1e-6 and abs(last['reaction_y_kN']-1864)<1e-6
assert abs(last['global_equilibrium_closure_y_kN'])<1e-6
E,nu,gamma,H=30000,.3,18,17;M=E*(1-nu)/((1+nu)*(1-2*nu));settlement=gamma*H**2/(2*M)
assert abs(settlement-.06440571428571428)<1e-14
summary=dict(regression=record,one_changed_input='road material8 gamma 19->20 only',expected_additional_weight_kN=22,observed_final_reaction_kN=last['reaction_y_kN'],initial_reaction_kN=first['reaction_y_kN'],final_closure_kN=last['global_equilibrium_closure_y_kN'],hand_calculation=dict(constrained_modulus_kPa=M,settlement_m=settlement,assumptions='uniform horizontal elastic layer; laterally constrained; not applicable directly to slope'),validation_scope='Programmatic Excel Validation.Value/STOP metadata and VBE compilation, not novice physical keyboard-dialog user testing. Observer zero threshold tests are synthetic, not a full plastic hold/SRM regression. No convergence/material conditions loosened.')
if '--check-only' not in sys.argv:(out/'review_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
print('PASS two 11-stage exact input/nodal/GP/stage comparisons; gamma exercise and independent hand calculation')
