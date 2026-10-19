"""Check hierarchical native QPC totals; include counts to expose timer overhead."""
import argparse,json
from pathlib import Path

def summarize(record,names):
    d=record['NativeProfile']
    assert record['ExclusiveExcel'] and d[0]>0 and d[1]==0
    rows=[dict(procedure=name,calls=int(d[2+3*i]),inclusive_seconds=d[3+3*i],exclusive_seconds=d[4+3*i]) for i,name in enumerate(names)]
    root=rows[-1]; total=root['inclusive_seconds']
    assert root['calls']==830 and total>0
    assert abs(sum(r['exclusive_seconds'] for r in rows)-total)<1e-6
    assert all(0<=r['exclusive_seconds']<=r['inclusive_seconds']+1e-9 for r in rows)
    for r in rows:r['exclusive_percent']=100*r['exclusive_seconds']/total
    return dict(status='PASS',qpc_frequency=d[0],total_gmres_seconds=total,exclusive_partition_pass=True,
        note='Per-call QPC overhead is present; compare whole-run speed only with uninstrumented builds.',
        ranking=sorted(rows,key=lambda r:-r['exclusive_seconds']))

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('input');ap.add_argument('--names');ap.add_argument('--out',required=True);a=ap.parse_args()
    names=json.loads(Path(a.names).read_text()) if a.names else ['dot','axpy','csr_matvec','csr_resnorm','ilu_apply','sfn_gmres']
    record=json.loads(Path(a.input).read_text(encoding='utf-8-sig'))
    result=summarize(record,names);s=json.dumps(result,indent=2)
    Path(a.out).write_text(s+'\n',encoding='utf-8');print(s)
