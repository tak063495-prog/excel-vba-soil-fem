"""Recover continuum section resultants from four saved Gauss stress tensors.
Run from repository root. Compression-positive saved stress is negated before
forming traction. The method reconstructs, rather than invents, a smooth stress
inside each element; it does not guarantee exact discrete traction equilibrium.
"""
from pathlib import Path
import gzip,json,hashlib,csv,argparse,sys
import numpy as np
ROOT=Path.cwd();sys.path.insert(0,str(ROOT/'tests'))
from check_manual_case import N
A=.5773502692
GP=np.array([[-A,-A],[A,-A],[A,A],[-A,A]])
def basis(r,s):return np.array([1,r,s,r*s])
INV=np.linalg.inv(np.array([basis(r,s) for r,s in GP]))
def natural(xy,p):
    q=np.zeros(2)
    for it in range(30):
        r,s=q;delta=N(r,s)@xy-p
        dr=N(r+1e-30j,s).imag/1e-30;ds=N(r,s+1e-30j).imag/1e-30
        if np.linalg.norm(delta)<1e-11:return q
        q-=np.linalg.solve(np.column_stack([dr@xy,ds@xy]),delta)
        if np.max(np.abs(q))>20:return None
    return None
def segments(xy,axis,value,lo,hi):
    vals=[];trans=1-axis
    for i,j,k in [(0,4,1),(1,5,2),(2,6,3),(3,7,0)]:
        # Exact intersection with a quadratic Q8 edge.
        coeff=np.array([.5*(xy[i,axis]+xy[k,axis])-xy[j,axis],.5*(xy[k,axis]-xy[i,axis]),xy[j,axis]-value])
        if abs(coeff[0])<1e-12:
            roots=[] if abs(coeff[1])<1e-12 else [-coeff[2]/coeff[1]]
        else:roots=np.roots(coeff[:])
        for t in roots:
            if abs(np.imag(t))<1e-9 and -1-1e-9<=np.real(t)<=1+1e-9:
                t=float(np.real(t));p=.5*t*(t-1)*xy[i]+(1-t*t)*xy[j]+.5*t*(t+1)*xy[k];vals.append(p[trans])
    vv=sorted(set(round(float(v),11) for v in vals));ret=[]
    for aa,bb in zip(vv,vv[1:]):
        aa=max(aa,lo);bb=min(bb,hi)
        if bb-aa<1e-9:continue
        p=np.array([0.,0.]);p[axis]=value;p[trans]=(aa+bb)/2;q=natural(xy,p)
        if q is not None and max(abs(q))<=1+1e-7:ret.append((aa,bb))
    return ret
def integrate(elements,axis,value,lo,hi,order=12):
    z,w=np.polynomial.legendre.leggauss(order);cent=(lo+hi)/2;force=np.zeros(2);moment=0.;length=0.;used=[]
    for eid,xy,stress in elements:
        if value<xy[:,axis].min()-1e-8 or value>xy[:,axis].max()+1e-8:continue
        coef=INV@stress
        for aa,bb in segments(xy,axis,value,lo,hi):
            used.append(eid);length+=bb-aa
            for t,wt in zip(z,w):
                coord=(aa+bb)/2+t*(bb-aa)/2;p=np.zeros(2);p[axis]=value;p[1-axis]=coord;q=natural(xy,p)
                assert q is not None and max(abs(q))<1+1e-7,(eid,p,q)
                sx,sy,tau=-(basis(*q)@coef)
                traction=np.array([sx,tau]) if axis==0 else np.array([tau,sy])
                arm=np.array([0,coord-cent]) if axis==0 else np.array([coord-cent,0])
                ds=wt*(bb-aa)/2;force+=traction*ds;moment+=(arm[0]*traction[1]-arm[1]*traction[0])*ds
    return dict(N=force[axis],V=force[1-axis],M=moment,length=length,elements=sorted(set(used)))
def self_checks():
    # Affine stress fields on regular Q8 patches: known N,V,M.
    errors=[]
    for axis in [0,1]:
      for split in [1,3]:
        els=[]
        for i in range(split):
            low=-.2+i*.4/split;high=low+.4/split
            rr=np.array([[-1,-1],[1,-1],[1,1],[-1,1],[0,-1],[1,0],[0,1],[-1,0]],float)
            xy=np.column_stack([1+rr[:,0]*.5,(low+high)/2+rr[:,1]*(high-low)/2])
            if axis==1:xy=xy[:,::-1]
            stress=[]
            for q in GP:
                p=N(*q)@xy;normal=30+700*p[1-axis];stress.append([-normal if axis==0 else -12,-normal if axis==1 else -12,-5])
            els.append((i,xy,np.array(stress)))
        got=integrate(els,axis,1.,-.2,.2)
        expected=[12.,2.,(1 if axis==1 else -1)*700*.4**3/12]
        error=max(abs(got[k]-v) for k,v in zip(['N','V','M'],expected));assert error<1e-8,(got,expected);errors.append(error)
    return dict(analytic_affine_patch_tests=4,max_absolute_error=max(errors),basis='Known membrane plus linear bending normal stress and constant shear; 1/3 through-thickness patches; both cut normals')
def recover(data,stage=3):
    t=data['tables'];nodes={int(r[0]):np.array(r[1:3],float) for r in t['節点データ'][1:]};sr={int(r[1]):r for r in t['結果要素'][1:] if r[0]==stage and isinstance(r[4],(int,float))}
    els=[]
    for e in t['要素データ'][1:]:
        if e[9]!=2:continue
        eid=int(e[0]);r=sr[eid];els.append((eid,np.array([nodes[int(n)] for n in e[1:9]]),np.array([[r[4+17*g+k] for k in range(3)] for g in range(4)],float)))
    assert els and data['status']=='PASS'
    stage_row=next(r for r in t['ステージ結果'][1:] if r[0]==stage)
    assert stage_row[1]=='BIRTH' and stage_row[4]==1
    return els
def balance(elements,kind,station):
    # Closed wall rectangle. Recovered interface tractions, not independently
    # known reactions: diagnostic of recovered-field equilibrium only.
    x0,x1,y0,y1=(18.6,19.,station,3.4) if kind=='stem' else (15.,station,0.,.4)
    ref=np.array([(x0+x1)/2,station]) if kind=='stem' else np.array([station,(y0+y1)/2])
    forces=[];moments=[]
    for axis,value,lo,hi,sign in [(0,x0,y0,y1,-1),(0,x1,y0,y1,1),(1,y0,x0,x1,-1),(1,y1,x0,x1,1)]:
        value-=sign*1e-7
        got=integrate(elements,axis,value,lo,hi,24)
        assert abs(got['length']-(hi-lo))<2e-6,(kind,station,axis,got['length'])
        force=np.array([got['N'],got['V']]) if axis==0 else np.array([got['V'],got['N']])
        center=np.array([value,(lo+hi)/2]) if axis==0 else np.array([(lo+hi)/2,value])
        arm=center-ref;forces.append(sign*force);moments.append(sign*(got['M']+arm[0]*force[1]-arm[1]*force[0]))
    weight=24*(x1-x0)*(y1-y0);body=np.array([0.,-weight]);arm=np.array([(x0+x1)/2,(y0+y1)/2])-ref
    residual=sum(forces)+body;mr=sum(moments)+arm[0]*body[1]
    return dict(part=kind,station_m=station,body_weight_kN_per_m=weight,force_closure_kN_per_m=residual.tolist(),moment_closure_kNm_per_m=mr,force_normalizer_kN_per_m=sum(np.linalg.norm(f) for f in forces)+weight,moment_normalizer_kNm_per_m=sum(abs(m) for m in moments)+abs(arm[0]*body[1]))

def run(out):
    out.mkdir(parents=True,exist_ok=True);checks=self_checks();records=[];sources=[];balances=[]
    sections=[('stem',1,y,18.6,19.) for y in [.6,.8,1.,1.4,1.8,2.2,2.6,3.0,3.2]]+[('base',0,x,0.,.4) for x in [15.2,15.6,16.,16.4,16.8,17.2,17.6,18.,18.2,18.4]]
    for case in ['wall_davis08','wall_davis12']:
        p=ROOT/f'docs/manuals/construction-evidence/20261020/{case}/result.json.gz';data=json.loads(gzip.decompress(p.read_bytes()));els=recover(data)
        for kind,station in [('stem',1.8),('base',18.)]:balances.append(dict(case=case,**balance(els,kind,station)))
        sources.append(dict(case=case,path=p.relative_to(ROOT).as_posix(),sha256=hashlib.sha256(p.read_bytes()).hexdigest(),wall_elements=len(els)))
        for kind,axis,station,lo,hi in sections:
            sides=[];errors=[]
            for sign in [-1,1]:
                v=station+sign*1e-7;got=integrate(els,axis,v,lo,hi,12);fine=integrate(els,axis,v,lo,hi,24)
                assert abs(got['length']-(hi-lo))<2e-7,(case,kind,station,got['length'])
                err=max(abs(got[k]-fine[k]) for k in ['N','V','M']);assert err<2e-4,(case,kind,station,err)
                sides.append(fine);errors.append(err)
            row=dict(case=case,stage=3,Fs=1.,part=kind,station_m=station,normal_axis='X' if axis==0 else 'Y',width_m=hi-lo,centroid_transverse_m=(hi+lo)/2,
                N_kN_per_m=sum(s['N'] for s in sides)/2,V_kN_per_m=sum(s['V'] for s in sides)/2,Mz_kNm_per_m=sum(s['M'] for s in sides)/2,
                M_minus=sides[0]['M'],M_plus=sides[1]['M'],M_one_sided_gap=abs(sides[1]['M']-sides[0]['M']),quadrature_max_abs_difference=max(errors),
                elements_minus=sides[0]['elements'],elements_plus=sides[1]['elements'])
            records.append(row)
    with (out/'sections.csv').open('w',encoding='utf8',newline='') as f:
        ww=csv.DictWriter(f,fieldnames=list(records[0]));ww.writeheader();ww.writerows(records)
    summary=dict(scope='Existing construction-complete stage3 Fs1, DAVIS, unchanged archived workbooks. No new Excel solve.',stress_convention='Saved compression-positive tensor C; tensile-positive T=-C. n=+Y for stem, n=+X for base. Mz=integral (r-rc) cross (T n) ds, positive counter-clockwise.',reconstruction='Bilinear in natural coordinates through four saved GP stress tensors; exact Q8 edge intersections, inverse mapping and Gauss line quadrature. Mean of fixed-station +/-1e-7m limits; gap retained.',unit_depth_m=1.,self_checks=checks,closed_subdomain_checks=balances,sources=sources,records=records)
    (out/'summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    print('PASS',len(records),'sections;',checks)
if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--out',type=Path,default=Path('tmp/wall-moment-review'));args=ap.parse_args();run(args.out)
