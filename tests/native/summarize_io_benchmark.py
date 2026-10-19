"""Paired I/O/assembly changes must preserve numerical results and event bytes."""
import argparse
import json
import math
import statistics
from pathlib import Path
from summarize_inco_benchmark import load

def compare(reference,candidate,scope='io'):
    if scope=='assembly':
        assert reference['IoFast']==candidate['IoFast']==reference['EventFast']==candidate['EventFast']==1
        assert reference['AssemblyFast']==0 and candidate['AssemblyFast']==1
        assert reference['Assembly'][0]==0 and candidate['Assembly'][0]>0
        assert reference['Assembly'][1:]==candidate['Assembly'][1:]
    elif scope=='events':
        assert reference['IoFast']==candidate['IoFast']==1
        assert reference['EventFast']==0 and candidate['EventFast']==1
    else:
        assert scope=='io' and reference['IoFast']==0 and candidate['IoFast']==1
        assert reference.get('EventFast',0)==candidate.get('EventFast',0)==0
    assert all(r['StateCopy']==r['GmresSimd']==1 and r['ExclusiveExcel'] is True for r in (reference,candidate))
    for key in ('Case','Variant','Mode','Precision','NativeDllSha256','Fos','Bracket','Disp','SolverEventsSha256','SolverEventsBytes'):
        assert reference[key]==candidate[key], ('Mismatch',key)
    assert reference['NativeDllSha256'], 'Native DLL hash missing'
    assert len(reference['Bracket'])==6 and all(math.isfinite(v) for v in reference['Bracket'][:4]), 'Incomplete raw Fs bracket'
    assert reference['Bracket'][4] is True and reference['Bracket'][5]>0, 'No completed Fs bracket'
    assert reference['SolverEventsSha256'] and reference['SolverEventsBytes']>0
    assert reference['Disp'] and all(math.isfinite(v) for v in reference['Disp'])
    r=reference['Run'].split('|');c=candidate['Run'].split('|')
    assert r[1:3]==c[1:3]==['PASS','True'] and r[3:8]==c[3:8]
    assert r[8]==c[8] and 'fallback=0' in c[8], 'Native path/counts changed'
    assert r[14:]==c[14:], 'Material path/counts changed'
    assert math.isfinite(float(c[3])) and float(c[3])>=0
    rt=float(reference['WallSeconds']);ct=float(candidate['WallSeconds'])
    assert math.isfinite(rt) and math.isfinite(ct) and 0<ct<rt
    return dict(reference_seconds=rt,candidate_seconds=ct,displacement_count=len(reference['Disp']),
        max_abs_displacement_delta=0.0,residual=float(c[3]),raw_bracket=candidate['Bracket'],
        solver_events_sha256=candidate['SolverEventsSha256'],solver_events_bytes=candidate['SolverEventsBytes'])

def main():
    ap=argparse.ArgumentParser();ap.add_argument('directory');ap.add_argument('--out',required=True)
    ap.add_argument('--pairs',type=int,default=3);ap.add_argument('--scope',choices=('io','events','assembly'),default='io');a=ap.parse_args();d=Path(a.directory)
    assert a.pairs>0
    rows=[dict(pair=i,**compare(load(d/f'pair{i}_enabled0.json'),load(d/f'pair{i}_enabled1.json'),a.scope)) for i in range(1,a.pairs+1)]
    r=[x['reference_seconds'] for x in rows];c=[x['candidate_seconds'] for x in rows]
    rm=statistics.median(r);cm=statistics.median(c)
    result=dict(status='PASS',comparison_scope=a.scope,pairs=rows,reference_median_seconds=rm,candidate_median_seconds=cm,
        reference_range_seconds=[min(r),max(r)],candidate_range_seconds=[min(c),max(c)],
        median_speedup=rm/cm,median_reduction_percent=100*(1-cm/rm))
    text=json.dumps(result,indent=2);Path(a.out).write_text(text+'\n',encoding='utf-8');print(text)

if __name__=='__main__':main()
