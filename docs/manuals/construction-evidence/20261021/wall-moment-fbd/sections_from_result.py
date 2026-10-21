"""Section forces (N, V, M) and free-body checks for ANY saved result.json.gz.
New file (2026-10-21). Does not modify rebuild_sections.py, which only
regenerates the archived wall_davis08/12 comparison.

Run from repository root:
  python docs/manuals/construction-evidence/20261021/wall-moment-fbd/sections_from_result.py \
      --result <your_out>/result.json.gz --stage 3 --material 2 --out tmp/my_sections

The script prints the source path, SHA-256, the stage row (kind/status) it read,
and writes sections.csv and fbd.json to --out (never into the source folder).
Wall geometry (stations, rectangles) is the L-wall of analysis example 9;
edit STEM/BASE below for other geometries.
"""
from pathlib import Path
import sys,json,gzip,hashlib,argparse,csv
import numpy as np
HERE=Path(__file__).resolve().parent
sys.path.insert(0,str(HERE.parent/'wall-moment-guide'))
import rebuild_sections as R
STEM=dict(x0=18.6,x1=19.0,ytop=3.4,stations=[.6,.8,1.,1.4,1.8,2.2,2.6,3.0,3.2])
BASE=dict(x0=15.0,y0=0.0,y1=0.4,stations=[15.2,15.6,16.,16.4,16.8,17.2,17.6,18.,18.2,18.4])
GAMMA=24.0
def recover(data,stage,mat):
    t=data['tables'];nodes={int(r[0]):np.array(r[1:3],float) for r in t['節点データ'][1:] if r[0] is not None}
    sr={int(r[1]):r for r in t['結果要素'][1:] if r[0]==stage and isinstance(r[4],(int,float))}
    els=[]
    for e in t['要素データ'][1:]:
        if e[0] is None or e[9]!=mat or int(e[0]) not in sr:continue
        r=sr[int(e[0])];els.append((int(e[0]),np.array([nodes[int(n)] for n in e[1:9]]),np.array([[r[4+17*g+k] for k in range(3)] for g in range(4)],float)))
    row=next(r for r in t['ステージ結果'][1:] if r[0]==stage)
    return els,row
def fbd(els,kind,station):
    if kind=='stem': x0,x1,y0,y1=STEM['x0'],STEM['x1'],station,STEM['ytop'];ref=np.array([(x0+x1)/2,station]);cut=2;csign=-1
    else: x0,x1,y0,y1=BASE['x0'],station,BASE['y0'],BASE['y1'];ref=np.array([station,(y0+y1)/2]);cut=1;csign=1
    faces=[]
    for axis,value,lo,hi,sign in [(0,x0,y0,y1,-1),(0,x1,y0,y1,1),(1,y0,x0,x1,-1),(1,y1,x0,x1,1)]:
        value-=sign*1e-7;g=R.integrate(els,axis,value,lo,hi,24)
        F=np.array([g['N'],g['V']]) if axis==0 else np.array([g['V'],g['N']])
        c=np.array([value,(lo+hi)/2]) if axis==0 else np.array([(lo+hi)/2,value]);a=c-ref
        faces.append(dict(F=(sign*F).tolist(),M=float(sign*(g['M']+a[0]*F[1]-a[1]*F[0])),M_section=g['M']))
    w=GAMMA*(x1-x0)*(y1-y0);Mb=float(((x0+x1)/2-ref[0])*(-w))
    others=sum(f['M'] for i,f in enumerate(faces) if i!=cut)+Mb
    return dict(part=kind,station_m=station,M_cut_integrated=faces[cut]['M_section'],M_cut_from_other_faces_and_weight=-others/csign,
                moment_imbalance_residual=sum(f['M'] for f in faces)+Mb,body_weight_kN_per_m=w)
def main():
    ap=argparse.ArgumentParser();ap.add_argument('--result',type=Path,required=True);ap.add_argument('--stage',type=int,required=True)
    ap.add_argument('--material',type=int,default=2);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args()
    raw=a.result.read_bytes();data=json.loads(gzip.decompress(raw));els,row=recover(data,a.stage,a.material)
    print('source',a.result.resolve());print('sha256',hashlib.sha256(raw).hexdigest());print('stage row',row[:6]);print('elements',len(els))
    assert els,'no elements of that material in that stage';a.out.mkdir(parents=True,exist_ok=True)
    rec=[]
    for kind,axis,sts,lo,hi in [('stem',1,STEM['stations'],STEM['x0'],STEM['x1']),('base',0,BASE['stations'],BASE['y0'],BASE['y1'])]:
        for s in sts:
            sd=[R.integrate(els,axis,s+g*1e-7,lo,hi,24) for g in (-1,1)]
            rec.append(dict(part=kind,station_m=s,N=sum(x['N'] for x in sd)/2,V=sum(x['V'] for x in sd)/2,M=sum(x['M'] for x in sd)/2))
    with (a.out/'sections.csv').open('w',newline='',encoding='utf8') as f:
        w=csv.DictWriter(f,fieldnames=list(rec[0]));w.writeheader();w.writerows(rec)
    res=dict(source=str(a.result.resolve()),sha256=hashlib.sha256(raw).hexdigest(),stage_row=row[:6],checks=[fbd(els,'stem',1.8),fbd(els,'base',18.0)])
    (a.out/'fbd.json').write_text(json.dumps(res,ensure_ascii=False,indent=1,default=str)+'\n',encoding='utf8');print(json.dumps(res['checks'],indent=1))
if __name__=='__main__':main()
