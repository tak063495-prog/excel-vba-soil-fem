"""Check native fallback equivalence and mixed-T6 RESET state preservation."""
import argparse,json
from pathlib import Path
from check_issue49_57 import read,close,check_reset_gravity_stress

def check(root):
    a=read(root/'targeted6/apex_native/result.json');b=read(root/'apex_off3/result.json')
    assert a['requested_native']=='AUTO' and b['requested_native']=='OFF'
    assert 'point=' in a['native'] and 'OFF_SELECTED' in b['native']
    assert a['book_sha256']==b['book_sha256'] and a['fixture_sha256']==b['fixture_sha256']
    assert a['status'].startswith('PASS|') and b['status'].startswith('PASS|')
    du=close(a['disp'],b['disp'],1e-12);stress=close(a['stress'],b['stress'],1e-12)
    pa=read(root/'targeted6/apex_native/結果塑性.json')[1:];pb=read(root/'apex_off3/結果塑性.json')[1:]
    multiplier=close([r[9] for r in pa],[r[9] for r in pb],1e-12)
    d=read(root/'mixed_reset3/expected_failure.json');m=read(Path('tests/fixtures/issues_xlsm/issue57_mixed_reset.json'))
    assert '曲線T6' in d['status'] and not d['status'].startswith('PASS|')
    coords=close([v for r in d['reference_coordinates'] for v in r[1:]],
                 [v for r in m['nodes'] for v in r[1:3]],1e-12)
    E,nu,gamma=m['materials'][0][1],m['materials'][0][2],m['materials'][0][4]
    md=E*(1-nu)/((1+nu)*(1-2*nu));H=max(r[2] for r in m['nodes'])
    displacement=close([v for r in d['nodes'] for v in r[1:]],
                      [v for r in m['nodes'] for v in (0,-gamma/md*(H*r[2]-.5*r[2]**2))],1e-9)
    reset_stress=check_reset_gravity_stress(m,d['stress'])
    mixed=read(root/'mixed_coupled3/result.json');resumed=read(root/'mixed_coupled3/resume.json')
    assert mixed['status'].startswith('PASS|') and resumed['status'].startswith('PASS|')
    mixed_u=close(mixed['disp'],resumed['disp'],1e-9)
    return {'status':'PASS','native_apex':{'max_u_m':du,'max_stress_kPa':stress,'max_multiplier':multiplier,'native_diagnostic':a['native']},
            'mixed_reset':{'preserved_coordinate_max_m':coords,'preserved_u_vs_self_weight_solution_max_m':displacement,'preserved_stress_vs_self_weight_solution_max_kPa':reset_stress},
            'mixed_coupled':{'resume_u_max_m':mixed_u,'fresh_metrics':mixed['metrics'],'resume_metrics':resumed['metrics']}}

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--root',type=Path,required=True);ap.add_argument('--out',type=Path,required=True)
    a=ap.parse_args();d=check(a.root);a.out.write_text(json.dumps(d,indent=2)+'\n',encoding='utf-8');print('PASS extra issue checks')
