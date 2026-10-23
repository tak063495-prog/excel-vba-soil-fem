"""Independent monolithic Biot acceptance for the confined elastic column.

All displacement DOFs and all physical pressure GPs are compared to saved
Excel frames. Pressure unknowns exist only at corner nodes. The mechanics
uses the application's documented 2x2 Q8 integration, fluid blocks 3x3.
"""
import argparse,csv,json
from pathlib import Path
import numpy as np
from test_fixed_stress_reference import reference_local,q8_shape,pressure_shape,grad

def read_nodes(folder):
    rows=list(csv.DictReader((folder/'node.csv').open(encoding='utf-8-sig')))
    return np.array([[float(x['ux']),float(x['uy'])] for x in rows]).ravel()

def check(fixture,root):
    f=json.loads(Path(fixture).read_text(encoding='utf-8-sig')); root=Path(root)
    coords=np.array([[row[1],row[2]] for row in f['nodes']]); nn=len(coords); nd=2*nn
    corner=sorted({n-1 for e in f['elements'] for n in e[1:5]}); pm={n:i for i,n in enumerate(corner)}; np_=len(corner)
    K=np.zeros((nd,nd)); Q=np.zeros((nd,np_)); M=np.zeros((np_,np_)); H=np.zeros_like(M); gravity=np.zeros(np_)
    E,nu=f['materials'][0][1:3]; alpha,storage,kx,ky=f['water']['coupledMaterials'][0][1:5]; gamma=f['water']['gamma']
    D=E/((1+nu)*(1-2*nu))*np.array([[1-nu,nu,0],[nu,1-nu,0],[0,0,(1-2*nu)/2]])
    gauss=[(-1/np.sqrt(3),-1/np.sqrt(3)),(1/np.sqrt(3),-1/np.sqrt(3)),(1/np.sqrt(3),1/np.sqrt(3)),(-1/np.sqrt(3),1/np.sqrt(3))]
    for e in f['elements']:
        nodes=np.array(e[1:9])-1; xy=coords[nodes]; dof=np.array([[2*n,2*n+1] for n in nodes]).ravel(); ip=[pm[n] for n in nodes[:4]]
        lm,lh,_,lq,lg=reference_local(xy,'q8',alpha,storage,kx,ky,gamma,E,nu)
        M[np.ix_(ip,ip)]+=lm; H[np.ix_(ip,ip)]+=lh; Q[np.ix_(dof,ip)]+=lq;gravity[ip]+=lg
        for x,y in gauss:
            d=grad(q8_shape,x,y); J=xy.T@d.T;g=np.linalg.solve(J.T,d);B=np.zeros((3,16))
            B[0,0::2]=g[0];B[1,1::2]=g[1];B[2,0::2]=g[1];B[2,1::2]=g[0]
            K[np.ix_(dof,dof)]+=B.T@D@B*np.linalg.det(J)
    fixed=np.array([[bool(n[3]),bool(n[4])] for n in f['nodes']]).ravel(); free=np.flatnonzero(~fixed)
    ymax=coords[:,1].max(); pfixed=np.isclose(coords[corner,1],ymax); pfree=np.flatnonzero(~pfixed)
    p=gamma*(ymax-coords[corner,1]); u=read_nodes(root/'stg01_GRAVITY'); load=np.zeros(nd)
    for row in f['loads']:
        assert row[3]=='FORCE';load[2*(int(row[1])-1)+(row[2]=='Y')]+=row[4]
    frames=sorted(root.glob('time_*/C*')); results=[]
    # One run only. Subsequent history tests may create a new run; don't combine.
    run=frames[0].parent;frames=sorted(run.glob('C*'))
    previous_force=np.zeros(nd)
    for folder in frames:
        state=(folder/'fixed_stress.csv').read_text(encoding='utf-8-sig').splitlines()
        phase=state[3].split(','); stage=int(phase[1]); elapsed=float(phase[2]);time=float(state[1].split(',')[1])
        dt=0. if not results else time-results[-1]['time_s']
        force=load
        b=np.r_[force-previous_force-Q@p,M@p-dt*gravity]
        A=np.block([[K,-Q],[Q.T,M+dt*H]])
        indices=np.r_[free,nd+pfree];sol=np.linalg.solve(A[np.ix_(indices,indices)],b[indices])
        du=np.zeros(nd);pnext=np.zeros(np_);du[free]=sol[:len(free)];pnext[pfree]=sol[len(free):]
        u+=du;p=pnext;previous_force=force.copy()
        expected=read_nodes(folder);uerr=np.max(np.abs(u-expected))
        rows=(folder/'water.csv').read_text(encoding='utf-8-sig').splitlines()[4:]
        # Header is line4 (index3): element,gauss,pressure_kPa.
        perr=0.
        for line in rows:
            if not line.strip():continue
            e,g,value=line.split(',');en=f['elements'][int(e)-1];x,y=gauss[int(g)-1]
            pred=pressure_shape('q8',x,y)@p[[pm[n-1] for n in en[1:5]]]
            perr=max(perr,abs(pred-float(value)))
        residual=np.max(np.abs((A@np.r_[du,p]-b)[indices]))
        results.append(dict(frame=folder.name,time_s=time,dt_s=dt,max_u_difference_m=float(uerr),max_p_difference_kPa=float(perr),absolute_linear_residual=float(residual)))
    worstu=max(r['max_u_difference_m'] for r in results);worstp=max(r['max_p_difference_kPa'] for r in results)
    out=dict(status='PASS' if worstu<1e-8 and worstp<1e-5 else 'FAIL',max_u_difference_m=worstu,max_p_difference_kPa=worstp,states=results)
    print(json.dumps(out,indent=2));assert out['status']=='PASS'
if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('fixture');ap.add_argument('root');a=ap.parse_args();check(a.fixture,a.root)
