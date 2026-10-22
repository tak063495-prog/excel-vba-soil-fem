"""Independent Q8 Darcy, force-balance and saved-result checks (no Excel)."""
import argparse, json, math
from pathlib import Path
import numpy as np

NAT=np.array([[-1,-1],[1,-1],[1,1],[-1,1],[0,-1],[1,0],[0,1],[-1,0]],float)
def poly(r,s):return np.array([1,r,s,r*s,r*r,s*s,r*r*s,r*s*s])
INV=np.linalg.inv(np.array([poly(r,s) for r,s in NAT]))
def shape(r,s):
    n=poly(r,s)@INV
    dr=np.array([0,1,0,s,2*r,0,2*r*s,s*s])@INV
    ds=np.array([0,0,1,r,0,2*s,r*r,2*r*s])@INV
    return n,np.array([dr,ds])
def load(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def table(folder):return load(Path(folder)/'水圧結果.json')
def rows(data,col):return [r for r in data[35:] if len(r)>col and isinstance(r[col],(int,float))]
def level(f,profile,x):
    if profile=='*':profile=f['water']['defaultProfile']
    pts=[r[1:3] for r in f['water']['profiles'] if r[0]==profile]
    assert len(pts)>=2 and pts[0][0]-1e-9<=x<=pts[-1][0]+1e-9
    return float(np.interp(x,np.array(pts)[:,0],np.array(pts)[:,1]))
def independent(f):
    xy=np.array([r[1:3] for r in f['nodes']],float); nn=len(xy); k=np.zeros((nn,nn))
    conduct={int(r[0]):float(r[2]) for r in f['water']['materials']}; mats={int(r[0]):r for r in f['materials']}
    flowing={int(r[0]):r for r in f['elements'] if mats[int(r[9])][10]=='SOIL'}
    active=np.zeros(nn,bool); gp,wt=np.polynomial.legendre.leggauss(3)
    for e in flowing.values():
        ids=np.array(e[1:9],int)-1; coords=xy[ids]; ke=np.zeros((8,8)); active[ids]=True
        for r,wr in zip(gp,wt):
            for s,ws in zip(gp,wt):
                _,d=shape(r,s); jac=coords.T@d.T; det=np.linalg.det(jac); assert det>0
                grad=np.linalg.solve(jac.T,d)
                ke+=wr*ws*det*conduct[int(e[9])]*float(mats[int(e[9])][3])*(grad.T@grad)
        k[np.ix_(ids,ids)]+=ke
    bc={}
    for b in f['water']['boundaries']:
        if b[0] not in flowing or b[2] not in ('HEAD','EXPOSED'):continue
        e=flowing[b[0]]; s=int(b[1])-1
        for loc in (s,(s+1)%4,s+4):
            n=int(e[loc+1])-1; val=level(f,b[3],xy[n,0]) if b[3] else float(b[4])
            assert n not in bc or abs(bc[n]-val)<1e-9
            bc[n]=val
    assert bc
    locked=np.array(sorted(bc),int); free=np.array([i for i in range(nn) if active[i] and i not in bc],int)
    h=np.zeros(nn);h[locked]=[bc[i] for i in locked]
    h[free]=np.linalg.solve(k[np.ix_(free,free)],-k[np.ix_(free,locked)]@h[locked])
    return xy,h,flowing,k,active,bc

def check_cutoff(folder,fixture):
    f=load(fixture); xy,h,flow,k,active,bc=independent(f); data=table(folder); record=load(Path(folder)/'result.json')
    assert record['status'].startswith('PASS|') and record['metrics'][3]<1e-7
    maxp=maxq=maxh=0.; gp=[(-1,-1),(1,-1),(1,1),(-1,1)]; gamma=f['water']['gamma']
    for row in rows(data,1):
        eid,g=int(row[1]),int(row[2]);
        if eid not in flow:continue
        e=flow[eid]; ids=np.array(e[1:9],int)-1; r,s=np.array(gp[g-1])/math.sqrt(3); n,d=shape(r,s); coords=xy[ids]
        p=gamma*(n@h[ids]-(n@coords)[1]); grad=np.linalg.solve((coords.T@d.T).T,d); q=-float(next(m[2] for m in f['water']['materials'] if m[0]==e[9]))*(grad@h[ids])
        maxp=max(maxp,abs(row[5]-p));maxq=max(maxq,float(np.max(np.abs(np.array(row[12:14])-q))))
    for row in rows(data,32):
        nxy=np.array(row[37:39]); near=np.linalg.norm(xy-nxy,axis=1); idx=int(near.argmin());assert near[idx]<1e-8
        if active[idx]:maxh=max(maxh,abs(h[idx]-row[33]))
    mass={}; qscale=0.
    for row in rows(data,41):
        e=int(row[41]);mass[e]=mass.get(e,0.)+row[43];qscale=max(qscale,abs(row[43]))
    local=max(abs(v) for v in mass.values())/max(qscale,1e-20)
    assert maxp<1e-7 and maxq<1e-12 and maxh<1e-8 and local<1e-8,(maxp,maxq,maxh,local)
    reaction=k@h; supplied=sum(max(reaction[i],0.) for i in bc);removed=-sum(min(reaction[i],0.) for i in bc)
    assert abs(supplied-removed)<1e-12
    return dict(case='cutoff',max_pressure_error_kPa=maxp,max_flux_error_m_s=maxq,max_head_error_m=maxh,local_balance=local,inflow_m3_s=supplied,outflow_m3_s=removed)

def check_box(folder):
    a=load(Path(folder)/'result.json');data=table(folder); fx=fy=moment=0.
    for row in rows(data,23):fx+=row[27];fy+=row[28];moment+=row[29]
    assert abs(fx)<1e-9 and abs(fy-90)<1e-9 and abs(moment-135)<1e-9,(fx,fy,moment)
    assert abs(a['metrics'][0]-192)<1e-9 and abs(a['metrics'][2]-102)<1e-8 and a['metrics'][3]<1e-8
    b=load(Path(folder)/'resume.json');du=float(np.max(np.abs(np.array(a['disp'])-b['disp'])))
    assert b['status'].startswith('PASS|') and du<1e-9 and b['metrics'][3]<1e-7 and abs(a['metrics'][2]-b['metrics'][2])<1e-6,b['metrics']
    return dict(case='box',water_Fx_kN=fx,water_Fy_kN=fy,water_M_kNm=moment,reaction_Y_kN=a['metrics'][2],resume_max_du_m=du,resume_residual=b['metrics'][3])

def check_basic(folder):
    a=load(Path(folder)/'results.json');out=[]; constrained=10000*(1-.3)/((1+.3)*(1-2*.3))
    for row in a:
        if 'result' in row:assert row['result'].startswith('PASS');continue
        effective=20 if row['case'] in ('dry','total') else 20-9.80665
        expected=-effective*2*2/(2*constrained); d=np.array(row['disp']);diff=abs(min(d[1::2])-expected)
        assert row['status'].startswith('PASS|') and diff<1e-10 and max(abs(d[::2]))<1e-10
        out.append(dict(case=row['case'],settlement=min(d[1::2]),analytic=expected,error=diff))
    return out

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--basic',type=Path);ap.add_argument('--box',type=Path);ap.add_argument('--cutoff',type=Path);ap.add_argument('--fixture',type=Path,default=Path(__file__).parent/'fixtures/groundwater_cutoff.json');ap.add_argument('--out',type=Path);a=ap.parse_args(); result=[]
    if a.basic:result+=check_basic(a.basic)
    if a.box:result.append(check_box(a.box))
    if a.cutoff:result.append(check_cutoff(a.cutoff,a.fixture))
    assert result
    text=json.dumps(result,ensure_ascii=False,indent=2)
    if a.out:a.out.write_text(text,encoding='utf-8')
    print(text)
if __name__=='__main__':main()
