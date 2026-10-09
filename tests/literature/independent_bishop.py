"""Independent dry homogeneous circular-slip comparison, Bishop simplified.

This is a bounded search, not a certified global minimum or an FEM oracle.
Formula: F=sum((c*b+W*tan(phi))/m)/sum(W*sin(alpha)),
m=cos(alpha)+sin(alpha)*tan(phi)/F, with zero pore pressure.
"""
import argparse
import hashlib
import json
from pathlib import Path
import numpy as np


def ground(x, crest, run, top, toe):
    return np.where(x<=crest, top, np.where(x<crest+run, top-(top-toe)*(x-crest)/run, toe))


def evaluate(parameters, model, slices):
    entry, exit, distance = parameters.T
    y1=ground(entry,**model["ground"]); y2=ground(exit,**model["ground"])
    dx=exit-entry; dy=y2-y1; chord=np.hypot(dx,dy)
    xc=(entry+exit)/2-distance*dy/chord; yc=(y1+y2)/2+distance*dx/chord
    radius=np.hypot(distance,chord/2)
    frac=(np.arange(slices)+.5)/slices
    x=entry[:,None]+dx[:,None]*frac
    radicand=radius[:,None]**2-(x-xc[:,None])**2
    root=np.sqrt(np.maximum(radicand,1e-24))
    base=yc[:,None]-root
    heights=ground(x,**model["ground"])-base
    bottom=np.where((xc>=entry)&(xc<=exit),yc-radius,np.minimum(y1,y2))
    valid=(dx>.05)&(bottom>=-1e-9)&(yc>np.maximum(y1,y2))&np.all(heights>=-1e-9,axis=1)&np.all(radicand>0,axis=1)
    # On each straight ground segment, ground-minus-lower-arc is concave;
    # its minimum is at a segment endpoint. Check every terrain kink exactly.
    for breakpoint in [model["ground"]["crest"],model["ground"]["crest"]+model["ground"]["run"]]:
        x_break=np.clip(breakpoint,entry,exit)
        rad_break=radius**2-(x_break-xc)**2
        base_break=yc-np.sqrt(np.maximum(rad_break,1e-24))
        valid&=(rad_break>0)&(ground(x_break,**model["ground"])-base_break>=-1e-9)
    width=dx/slices
    weight=model["gamma"]*np.maximum(heights,0)*width[:,None]
    sin_alpha=(xc[:,None]-x)/radius[:,None]
    cos_alpha=root/radius[:,None]
    drive=np.sum(weight*sin_alpha,axis=1)
    valid&=drive>1e-8
    friction=np.tan(np.deg2rad(model["phi"]))
    numerator=model["c"]*width[:,None]+weight*friction
    fos=np.full(len(parameters),1.5)
    converged=np.zeros(len(parameters),dtype=bool)
    for _ in range(120):
        m=cos_alpha+sin_alpha*friction/fos[:,None]
        valid&=np.all(m>1e-8,axis=1)
        new=np.sum(numerator/np.maximum(m,1e-8),axis=1)/np.maximum(drive,1e-8)
        change=np.abs(new-fos)
        converged|=change<1e-9*np.maximum(1,np.abs(new))
        fos=.5*fos+.5*new
        if np.all(converged|~valid):break
    # Validate the actual fixed point, not only an earlier iteration's residual.
    m=cos_alpha+sin_alpha*friction/fos[:,None]
    check=np.sum(numerator/np.maximum(m,1e-8),axis=1)/np.maximum(drive,1e-8)
    valid&=(np.abs(check-fos)<1e-7*np.maximum(1,np.abs(fos)))&np.isfinite(fos)
    fos=np.where(valid,fos,np.inf)
    geometry=np.column_stack([xc,yc,radius,bottom])
    return fos,geometry


def search(model,seed):
    rng=np.random.default_rng(seed); width=model["width"]; height=model["ground"]["top"]-model["ground"]["toe"]
    pool=[]
    for _ in range(5):
        entry=rng.uniform(0,model["ground"]["crest"]+.8*model["ground"]["run"],12000)
        exit=entry+rng.uniform(.02,1,12000)*(width-entry)
        distance=height*np.exp(rng.uniform(np.log(.02),np.log(30),12000))
        p=np.column_stack([entry,exit,distance]); f,_=evaluate(p,model,120)
        finite=np.isfinite(f)
        p=p[finite];f=f[finite]
        if len(p):pool.append(p[np.argsort(f)[:60]])
    if not pool:raise RuntimeError("No valid circles")
    candidates=np.concatenate(pool)
    for level in range(13):
        f,_=evaluate(candidates,model,160)
        best=candidates[np.argsort(f)[:80]]
        perturb=rng.normal(size=(len(best),70,3))
        scale=np.array([width*.035,width*.035,height*.25])*(.66**level)
        trial=(best[:,None,:]+perturb*scale).reshape(-1,3)
        trial[:,0]=np.clip(trial[:,0],0,width-.1)
        trial[:,1]=np.clip(trial[:,1],.1,width)
        trial[:,2]=np.maximum(trial[:,2],.0001)
        candidates=np.concatenate([best,trial])
    # Freeze the highest-resolution optimum for the discretization comparison.
    f,g=evaluate(candidates,model,640); index=np.argmin(f)
    fixed_candidate=candidates[index:index+1]
    rows=[]; search_rows=[]
    for slices in [80,160,320,640]:
        fixed_f,fixed_g=evaluate(fixed_candidate,model,slices)
        rows.append(dict(slices=slices,fos=float(fixed_f[0]),entry_x=float(fixed_candidate[0,0]),
                         exit_x=float(fixed_candidate[0,1]),center_distance=float(fixed_candidate[0,2]),
                         circle_center_x=float(fixed_g[0,0]),circle_center_y=float(fixed_g[0,1]),
                         radius=float(fixed_g[0,2]),minimum_elevation=float(fixed_g[0,3])))
        f,g=evaluate(candidates,model,slices);i=np.argmin(f)
        search_rows.append(dict(slices=slices,fos=float(f[i]),entry_x=float(candidates[i,0]),
                         exit_x=float(candidates[i,1]),center_distance=float(candidates[i,2]),
                         circle_center_x=float(g[i,0]),circle_center_y=float(g[i,1]),
                         radius=float(g[i,2]),minimum_elevation=float(g[i,3])))
    return dict(seed=seed,candidate_count=len(candidates),slice_convergence=rows,
                searched_minimum_by_slices=search_rows)


def main():
    parser=argparse.ArgumentParser();parser.add_argument("--output",type=Path,required=True)
    args=parser.parse_args()
    models=[
        dict(name="griffiths_ex1",width=32,c=10,phi=20,gamma=20,ground=dict(crest=12,run=20,top=10,toe=0),published=1.38),
        dict(name="griffiths_ex3_phi0",width=60,c=50,phi=0,gamma=20,ground=dict(crest=20,run=20,top=20,toe=10),published=1.47),
        dict(name="pruska_h7_phi10",width=40,c=20,phi=10,gamma=24,ground=dict(crest=15,run=10,top=15,toe=8),published=1.22),
        dict(name="pruska_h7_phi20",width=40,c=20,phi=20,gamma=24,ground=dict(crest=15,run=10,top=15,toe=8),published=1.65),
        dict(name="pruska_h105_phi10",width=40,c=20,phi=10,gamma=24,ground=dict(crest=15,run=10,top=18.5,toe=8),published=.84)]
    report=dict(method="independent Bishop simplified, homogeneous dry circular slips",
                script_sha256=hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
                limits=["finite deterministic search; circular mechanism restriction","no FEM side restraints, no K0 or stress path","not an exact oracle or certified global minimum"],cases=[])
    args.output.parent.mkdir(parents=True,exist_ok=True)
    for model in models:
        runs=[search(model,20261009),search(model,20261010)]
        report["cases"].append(dict(model=model,runs=runs))
        args.output.write_text(json.dumps(report,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
        print(model["name"],[(r["seed"],r["slice_convergence"][-1]["fos"]) for r in runs],flush=True)


if __name__=="__main__":main()
