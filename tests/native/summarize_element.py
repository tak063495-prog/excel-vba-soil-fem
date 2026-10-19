"""Strict final-output and complete event-byte comparison for element/ILU changes."""
import argparse
import json
import math
import statistics
from pathlib import Path

def read(p):
    return json.loads(Path(p).read_text(encoding='utf-8-sig'))

def compare(a,b):
    for key in ('Case','Variant','Mode','Precision','Flow','NativeDllSha256','Assembly','Fos','Bracket','Disp',
                'SolverEventsSha256','SolverEventsBytes','DecisionEventsSha256','GmresEventRows',
                'GmresPrepared','PreparedCalls','TunedCalls'):
        assert a[key] == b[key], key
    assert a['MaterialDllSha256'] != b['MaterialDllSha256']
    for d in (a,b):
        assert d['ExclusiveExcel'] is True
        assert all(d[k] == 1 for k in ('StateCopy','GmresSimd','IoFast','EventFast','AssemblyFast','GmresTuned','GmresPrepared'))
        assert d['SolverEventsBytes'] > 0 and d['SolverEventsSha256']
        assert d['Disp'] and all(math.isfinite(x) for x in d['Disp'])
        assert d['WallSeconds'] > 0 and math.isfinite(d['WallSeconds'])
    ar,br = a['Run'].split('|'),b['Run'].split('|')
    assert ar[1:3] == br[1:3] == ['PASS','True']
    assert ar[3:9] == br[3:9] and ar[14:] == br[14:], 'Residual or counts changed'
    assert math.isfinite(float(br[3])) and float(br[3]) >= 0
    return dict(flow=a['Flow'],case=a['Case'],variant=a['Variant'],reference_seconds=a['WallSeconds'],candidate_seconds=b['WallSeconds'],
                reduction_percent=100*(1-b['WallSeconds']/a['WallSeconds']),displacement_count=len(b['Disp']),
                max_abs_displacement_delta=0.,final_residual=float(br[3]),bracket=b['Bracket'],
                logs_exact=True,prepared_calls=b['PreparedCalls'])

if __name__ == '__main__':
    ap=argparse.ArgumentParser();ap.add_argument('input');ap.add_argument('--pairs',type=int,default=3)
    ap.add_argument('--regression',action='store_true');ap.add_argument('--out',required=True);a=ap.parse_args()
    root=Path(a.input); reports=[]
    names=('INCONSISTENT','DAVIS','builtin') if a.regression else [f'pair{i}' for i in range(1,a.pairs+1)]
    for name in names:
        r,c=[read(root/f'{name}_enabled{i}.json') for i in (0,1)]
        r=r if isinstance(r,list) else [r];c=c if isinstance(c,list) else [c]
        assert len(r)==len(c)>0
        reports.extend(compare(x,y) for x,y in zip(r,c))
    result=dict(status='PASS',cases=reports)
    if not a.regression:
        rm=statistics.median(x['reference_seconds'] for x in reports)
        cm=statistics.median(x['candidate_seconds'] for x in reports)
        result.update(reference_median_seconds=rm,candidate_median_seconds=cm,median_reduction_percent=100*(1-cm/rm))
        if cm>=rm:result['status']='NO_IMPROVEMENT'
    Path(a.out).write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,indent=2))
