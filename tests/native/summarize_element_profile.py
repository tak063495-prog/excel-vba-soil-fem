"""Partition disposable element, ILU-build and regularized-correction QPC roots."""
import argparse,json
from pathlib import Path

def summarize(d,names):
    p=d['NativeProfile'];assert len(names)==10 and p[0]>0 and p[1]==0
    rows=[dict(procedure=name,calls=int(p[2+3*i]),inclusive_seconds=p[3+3*i],exclusive_seconds=p[4+3*i]) for i,name in enumerate(names)]
    roots=sum(rows[i]['inclusive_seconds'] for i in (0,1,8,9))
    assert abs(roots-sum(r['exclusive_seconds'] for r in rows))<1e-7
    assert all(0<=r['exclusive_seconds']<=r['inclusive_seconds']+1e-9 for r in rows)
    return dict(status='PASS',instrumented_roots_seconds=roots,partition_exact=True,procedures=rows,
                note='QPC calls and inhibited inlining affect timings. These times are diagnostic; use production-DLL pairs for total speed.')

if __name__=='__main__':
    a=argparse.ArgumentParser();a.add_argument('record');a.add_argument('--names',required=True);a.add_argument('--out',required=True);x=a.parse_args()
    r=summarize(json.loads(Path(x.record).read_text(encoding='utf-8-sig')),json.loads(Path(x.names).read_text()))
    Path(x.out).write_text(json.dumps(r,indent=2)+'\n',encoding='utf-8');print(json.dumps(r,indent=2))
