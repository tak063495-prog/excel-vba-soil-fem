"""Compare coupled regressions after accepted-load/output restoration fixes."""
import argparse
import json
from pathlib import Path


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def compare(reference, candidate):
    cases=[]
    for folder in sorted(candidate.iterdir()):
        if not folder.is_dir():
            continue
        old=reference/folder.name
        if (folder/'result.json').exists():
            records=[]
            for filename in ('result.json','resume.json'):
                a,b=read(old/filename),read(folder/filename)
                assert a['status'].startswith('PASS|') and b['status'].startswith('PASS|')
                if filename=='result.json': assert a['fixture_sha256']==b['fixture_sha256']
                # All displacement, stress, and Fs values must remain exact.
                for key in ('disp','stress'):
                    assert a[key]==b[key], (folder.name,filename,key)
                assert a['metrics'][4:]==b['metrics'][4:], (folder.name,filename,'Fs')
                delta=max(abs(x-y) for x,y in zip(a['metrics'][:4],b['metrics'][:4]))
                # CSV load locks and rebuilt internal forces may change roundoff
                # in the output-only reaction/residual sums. No solver gate changes.
                assert delta<=1e-10, (folder.name,filename,'diagnostics',delta)
                records.append({'record':filename,'displacement_values':len(b['disp']),
                                'stress_values':len(b['stress']),'u_stress_Fs_exact':True,
                                'diagnostic_max_difference':delta,
                                'reference_seconds':a['seconds'],'candidate_seconds':b['seconds']})
            cases.append({'case':folder.name,'status':'PASS','records':records})
        else:
            a,b=read(old/'expected_failure.json'),read(folder/'expected_failure.json')
            assert a['expected']==b['expected'] and b['expected'] in b['status']
            for key in ('coupled','nodes','pressure','stress','profile'):
                assert a[key]==b[key], (folder.name,key)
            cases.append({'case':folder.name,'status':'EXPECTED_REJECTION_UNCHANGED'})
    assert len(cases)==17, len(cases)
    return {'status':'PASS','scope':'all coupled displacement/stress/Fs values exact; output diagnostics within 1e-10; rejected states exact','cases':cases}


if __name__=='__main__':
    ap=argparse.ArgumentParser()
    for name in ('reference','candidate','report'): ap.add_argument('--'+name,type=Path,required=True)
    args=ap.parse_args()
    record=compare(args.reference,args.candidate)
    args.report.write_text(json.dumps(record,indent=2),encoding='utf-8')
    print('PASS: 17 coupled cases')
