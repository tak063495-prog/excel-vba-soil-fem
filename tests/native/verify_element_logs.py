"""Independently re-read actual full CSV files for each production comparison pair."""
import argparse,hashlib,json
from pathlib import Path

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('directory');p.add_argument('--pairs',type=int,default=3);p.add_argument('--out',required=True);a=p.parse_args()
    root=Path(a.directory);reports=[]
    for i in range(1,a.pairs+1):
        data=[]
        for side in (0,1):
            r=json.loads((root/f'pair{i}_enabled{side}.json').read_text(encoding='utf-8-sig'))
            raw=(Path(r['RunOutputDirectory'])/'solver_events.csv').read_bytes()
            digest=hashlib.sha256(raw).hexdigest()
            assert digest.lower()==r['SolverEventsSha256'].lower()
            assert len(raw)==r['SolverEventsBytes']>0
            # Every GMRES row remains included in this comparison.
            data.append(raw)
        assert data[0]==data[1],f'Full CSV differs in pair {i}'
        reports.append(dict(pair=i,bytes=len(data[0]),sha256=hashlib.sha256(data[0]).hexdigest(),all_bytes_equal=True))
    result=dict(status='PASS',gmres_rows_excluded=False,pairs=reports)
    Path(a.out).write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8');print(json.dumps(result,indent=2))
