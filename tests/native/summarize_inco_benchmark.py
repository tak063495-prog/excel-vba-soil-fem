"""Validate paired full-SRM runs and summarize high-resolution wall times."""
import argparse
import json
import math
import statistics
from pathlib import Path

def load(path):
    value=json.loads(Path(path).read_text(encoding='utf-8-sig'))
    if isinstance(value,list):
        assert len(value)==1, 'Expected one full-SRM case per file'
        return value[0]
    return value

def compare(reference,candidate,allow_legacy_bracket=False):
    assert reference['StateCopy']==reference['GmresSimd']==0, 'Reference flags must be OFF'
    assert candidate['StateCopy']==candidate['GmresSimd']==1, 'Candidate flags must be ON'
    for key in ('Case','Variant','Mode','Precision','NativeDllSha256','Fos'):
        assert reference[key]==candidate[key], ('Mismatch',key)
    assert reference['NativeDllSha256'], 'DLL hash missing'
    assert reference['ExclusiveExcel'] is True and candidate['ExclusiveExcel'] is True
    if 'Bracket' in reference or 'Bracket' in candidate:
        assert 'Bracket' in reference and 'Bracket' in candidate, 'Missing raw bracket on one side'
        assert reference['Bracket']==candidate['Bracket'], 'Raw Fs bracket mismatch'
    else:
        assert allow_legacy_bracket, 'Missing raw Fs bracket; explicit legacy opt-in required'
    r=reference['Run'].split('|'); c=candidate['Run'].split('|')
    assert r[1:3]==c[1:3]==['PASS','True'], 'Analysis did not pass'
    assert 'statecopy=0;' in r[8] and 'gmres_simd=0;' in r[8], 'Reference used acceleration'
    assert 'fallback=0' in c[8] and 'statecopy=0' not in c[8] and 'gmres_simd=0' not in c[8]
    residuals=[float(r[3]),float(c[3])]
    assert all(math.isfinite(v) and v>=0 for v in residuals)
    assert residuals[1]<=max(1e-10,10*residuals[0]), 'Residual degraded'
    x=reference['Disp']; y=candidate['Disp']
    assert len(x)==len(y)>0 and all(math.isfinite(v) for v in x+y)
    delta=max(abs(u-v) for u,v in zip(x,y)); scale=max(map(abs,x))
    assert delta<=1e-12+1e-9*scale, 'Displacement changed beyond comparison gate'
    rt=float(reference['WallSeconds']); ct=float(candidate['WallSeconds'])
    assert math.isfinite(rt) and math.isfinite(ct) and 0<ct<rt, 'No measured improvement'
    return dict(reference_seconds=rt,candidate_seconds=ct,speedup=rt/ct,
        displacement_count=len(x),max_abs_displacement_delta=delta,
        relative_to_max_displacement=delta/(scale or 1),displacements_exact=(x==y),
        residuals=residuals,residuals_exact=(residuals[0]==residuals[1]),
        major_counts_equal=(r[4:8]==c[4:8]),raw_bracket=reference.get('Bracket'))

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('directory'); ap.add_argument('--pairs',type=int,default=3)
    ap.add_argument('--out',required=True)
    ap.add_argument('--allow-legacy-bracket',action='store_true',help='Only for archived pairs lacking both raw bracket fields')
    args=ap.parse_args(); directory=Path(args.directory)
    assert args.pairs>0
    rows=[dict(pair=i,**compare(load(directory/f'pair{i}_enabled0.json'),load(directory/f'pair{i}_enabled1.json'),args.allow_legacy_bracket)) for i in range(1,args.pairs+1)]
    baseline=[r['reference_seconds'] for r in rows]; candidate=[r['candidate_seconds'] for r in rows]
    rb=statistics.median(baseline); rc=statistics.median(candidate)
    result=dict(status='PASS',pairs=rows,reference_median_seconds=rb,candidate_median_seconds=rc,
        reference_range_seconds=[min(baseline),max(baseline)],candidate_range_seconds=[min(candidate),max(candidate)],
        median_speedup=rb/rc,median_reduction_percent=100*(1-rc/rb))
    text=json.dumps(result,ensure_ascii=False,indent=2)
    Path(args.out).write_text(text+'\n',encoding='utf-8'); print(text)

if __name__=='__main__': main()
