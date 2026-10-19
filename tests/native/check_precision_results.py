"""Assert finite displacements, relative agreement, status and Fs agreement.

One file compares OFF/AUTO; two files compare matching Case/Variant/Mode.
Counts may differ in floating-point modes and are reported separately.
"""
import argparse
import json
import math
from pathlib import Path

def read(path):
    data=json.loads(Path(path).read_text(encoding='utf-8-sig'))
    return data if isinstance(data,list) else [data]

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('reference')
    parser.add_argument('candidate',nargs='?')
    parser.add_argument('--rtol',type=float,default=1e-9)
    parser.add_argument('--atol',type=float,default=1e-12)
    args=parser.parse_args()
    a=read(args.reference); b=read(args.candidate) if args.candidate else a
    key=lambda row:(row['Case'],row['Variant'],row['Mode'] if args.candidate else '')
    left={key(r):r for r in a if args.candidate or r['Mode']=='OFF'}
    right={key(r):r for r in b if args.candidate or r['Mode']=='AUTO'}
    assert left and left.keys()==right.keys(), 'Missing comparison rows'
    output=[]
    for k,r in left.items():
        c=right[k]; x=list(map(float,r['Disp'])); y=list(map(float,c['Disp']))
        assert len(x)==len(y) and len(x)>0, ('Shape mismatch',k)
        assert all(math.isfinite(v) for v in x+y), ('Nonfinite displacement',k)
        delta=max(abs(u-v) for u,v in zip(x,y)); scale=max(map(abs,x))
        assert delta<=args.atol+args.rtol*scale, ('Displacement mismatch',k,delta,scale)
        rr=r['Run'].split('|'); cr=c['Run'].split('|')
        assert rr[1:3]==cr[1:3] and rr[1]!='ERROR', ('Status mismatch',k)
        assert r['Fos']==c['Fos'], ('Fs bracket/diagnostic mismatch',k)
        assert 'fallback=0' in cr[8], ('Native fallback',k,cr[8])
        output.append(dict(case=k,max_abs=delta,relative_to_max=delta/(scale or 1),counts_equal=rr[4:8]==cr[4:8],reference_seconds=float(rr[0]),candidate_seconds=float(cr[0])))
    print(json.dumps(output,ensure_ascii=False,indent=2))

if __name__=='__main__':
    main()
