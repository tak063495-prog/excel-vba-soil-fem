"""Compare identical compiled VBA with old/new material DLL and prepared ILU."""
import argparse,json,math,re,statistics
from pathlib import Path

def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def rows(x):return x if isinstance(x,list) else [x]

def compare(a,b,srm=True):
    for key in ('Case','Variant','Mode','Precision','Flow','NativeDllSha256','Assembly','Fos','Bracket','Disp'):
        assert a[key]==b[key],key
    assert a['MaterialDllSha256'] and b['MaterialDllSha256'] and a['MaterialDllSha256']!=b['MaterialDllSha256']
    for r in (a,b):
        assert r['ExclusiveExcel'] is True
        assert all(r[k]==1 for k in ('StateCopy','GmresSimd','IoFast','EventFast','AssemblyFast','GmresTuned'))
        assert r['NativeDllSha256'] and r['SolverEventsSha256'] and r['SolverEventsBytes']>0
        assert r['Disp'] and all(math.isfinite(x) for x in r['Disp'])
        assert len(r['Bracket'])==6 and all(math.isfinite(x) for x in r['Bracket'][:4])
    assert a['GmresPrepared']==0 and a['PreparedCalls']==0 and b['GmresPrepared']==1
    ar=a['Run'].split('|');br=b['Run'].split('|')
    assert ar[1:3]==br[1:3]==['PASS','True']
    assert ar[3:9]==br[3:9] and ar[14:]==br[14:],'Final residual or major/native/material counts changed'
    assert 'fallback=0' in br[8] and math.isfinite(float(br[3])) and float(br[3])>=0
    native=int(re.search(r'gmres_simd=(\d+)',br[8]).group(1))
    assert b['PreparedCalls']==b['TunedCalls']==a['TunedCalls']==native
    assert a['DecisionEventsSha256']==b['DecisionEventsSha256'] and b['DecisionEventsSha256']
    assert a['NonGmresSolverEventsRows']==b['NonGmresSolverEventsRows']>=0
    if srm:
        assert b['Bracket'][4] is True and b['Bracket'][5]>0 and native>0
    else:
        # Fixed-Fs cases do not exercise GMRES; require every event byte too.
        assert native==0 and a['SolverEventsSha256']==b['SolverEventsSha256'] and a['SolverEventsBytes']==b['SolverEventsBytes']
    rt=a['WallSeconds'];ct=b['WallSeconds']
    assert all(math.isfinite(t) and t>0 for t in (rt,ct))
    return dict(case=a['Case'],variant=a['Variant'],reference_seconds=rt,candidate_seconds=ct,
        reduction_percent=100*(1-ct/rt),displacement_count=len(b['Disp']),max_abs_displacement_delta=0.,
        final_residual=float(br[3]),bracket=b['Bracket'],prepared_calls=b['PreparedCalls'],
        decision_events_equal=True,logs_exact=a['SolverEventsSha256']==b['SolverEventsSha256'])

def main():
    ap=argparse.ArgumentParser();ap.add_argument('input');ap.add_argument('candidate',nargs='?')
    ap.add_argument('--pairs',type=int,default=3);ap.add_argument('--out',required=True);a=ap.parse_args()
    if a.candidate:
        reference=rows(read(a.input));candidate=rows(read(a.candidate));assert len(reference)==len(candidate)>0
        result=dict(status='PASS',cases=[compare(r,c,False) for r,c in zip(reference,candidate)])
    else:
        assert a.pairs>0
        report=[dict(pair=i,**compare(read(Path(a.input)/f'pair{i}_enabled0.json'),read(Path(a.input)/f'pair{i}_enabled1.json'))) for i in range(1,a.pairs+1)]
        r=[x['reference_seconds'] for x in report];c=[x['candidate_seconds'] for x in report]
        rm=statistics.median(r);cm=statistics.median(c)
        result=dict(status='PASS' if cm<rm else 'NO_IMPROVEMENT',pairs=report,reference_median_seconds=rm,candidate_median_seconds=cm,
            median_reduction_percent=100*(1-cm/rm),reference_range_seconds=[min(r),max(r)],candidate_range_seconds=[min(c),max(c)])
    Path(a.out).write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8');print(json.dumps(result,indent=2))

if __name__=='__main__':main()
