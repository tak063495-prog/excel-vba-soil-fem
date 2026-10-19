"""Partition the single-thread material QPC roots and shared math wrappers."""
import argparse,json
from pathlib import Path

def summarize(record,names,field='NativeProfile'):
    d=record[field];assert d[0]>0 and d[1]==0 and len(names)==10
    rows=[dict(procedure=name,calls=int(d[2+3*i]),inclusive_seconds=d[3+3*i],exclusive_seconds=d[4+3*i]) for i,name in enumerate(names)]
    assert rows[0]['procedure']=='sfm_eval' and rows[0]['calls']==56583
    assert rows[1]['procedure']=='sfm_tangent' and rows[1]['calls']==7803
    roots=rows[0]['inclusive_seconds']+rows[1]['inclusive_seconds']
    outside=sum(r['exclusive_seconds'] for r in rows)-roots
    assert outside>=-1e-8 and all(0<=r['exclusive_seconds']<=r['inclusive_seconds']+1e-9 for r in rows)
    return dict(status='PASS',material_roots_seconds=roots,math_wrappers_outside_material_roots_seconds=outside,
        note='Shared c_sqrt/c_pow/etc wrappers also time other material-DLL exports. QPC overhead is included; use uninstrumented pairs for whole-run speed.',ranking=sorted(rows,key=lambda r:-r['exclusive_seconds']))

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('record');ap.add_argument('--names',required=True);ap.add_argument('--out',required=True)
    ap.add_argument('--field',choices=('NativeProfile','MaterialNativeProfile'),default='NativeProfile');a=ap.parse_args()
    result=summarize(json.loads(Path(a.record).read_text(encoding='utf-8-sig')),json.loads(Path(a.names).read_text()),a.field)
    text=json.dumps(result,indent=2);Path(a.out).write_text(text+'\n');print(text)
