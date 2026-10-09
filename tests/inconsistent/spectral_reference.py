"""Independent six-facet KKT oracle; no sorted-rank branch selection."""
import argparse, itertools, json, math
from pathlib import Path
import numpy as np

def elastic(E, nu):
    G = E / (2 * (1 + nu)); L = E * nu / ((1 + nu) * (1 - 2 * nu))
    C3 = np.full((3, 3), L); np.fill_diagonal(C3, L + 2 * G)
    C4 = np.array([[L+2*G,L,0,L],[L,L+2*G,0,L],[0,0,G,0],[L,L,0,L+2*G]])
    return C3, C4

PAIRS = [(i, j) for i in range(3) for j in range(3) if i != j]
def oracle(case):
    C3, C4 = elastic(case['E'], case['nu'])
    previous = np.array(case['prev']); strain = np.r_[case['strain'], 0.]
    trial = previous + C4 @ strain
    values, vectors = np.linalg.eigh([[trial[0],trial[2]],[trial[2],trial[1]]])
    principal = np.r_[values, trial[3]]  # Physical plane axes + z; not sorted ranks.
    N = np.zeros((6,3)); M = N.copy()
    s, sp = np.sin(np.deg2rad([case['phi'],case['psi']]))
    strength = 2 * case['c'] * math.cos(math.radians(case['phi']))
    for a,(i,j) in enumerate(PAIRS):
        N[a,i]=1-s; N[a,j]=-1-s; M[a,i]=1-sp; M[a,j]=-1-sp
    F = N @ principal - strength
    tol = 1e-11 * (1 + np.max(np.abs(principal)) + strength)
    candidates = []
    if np.max(F) <= 1e-10 * (1 + np.max(np.abs(principal)) + strength):
        return dict(ok=True, stress=trial.tolist(), ep=[0.]*4, elastic=True)
    for size in (1,2,3):
        for active in itertools.combinations(range(6),size):
            indices=list(active); ns=N[indices]; ms=M[indices]
            A=ns @ C3 @ ms.T
            if np.linalg.matrix_rank(A, tol=1e-10*np.max(np.abs(A))) < size: continue
            gamma=np.linalg.solve(A,F[indices])
            if np.min(gamma) < -1e-13: continue
            result=principal - C3 @ ms.T @ gamma
            residual=N @ result - strength
            if np.max(residual) > tol or np.max(np.abs(residual[indices])) > tol: continue
            S=vectors @ np.diag(result[:2]) @ vectors.T
            stress=np.array([S[0,0],S[1,1],S[0,1],result[2]])
            ep=np.linalg.solve(C4, trial-stress)
            candidates.append((stress,ep))
    if not candidates: return dict(ok=False)
    reference=candidates[0]
    for stress, ep in candidates:
        assert np.max(np.abs(stress-reference[0])) < tol*10, 'nonunique oracle stress'
    return dict(ok=True,stress=reference[0].tolist(),ep=reference[1].tolist(),elastic=False)

def rotation(case, angle):
    a=math.radians(angle); R=np.array([[math.cos(a),-math.sin(a)],[math.sin(a),math.cos(a)]])
    S=R @ np.array([[case['prev'][0],case['prev'][2]],[case['prev'][2],case['prev'][1]]]) @ R.T
    e=R @ np.array([[case['strain'][0],case['strain'][2]/2],[case['strain'][2]/2,case['strain'][1]]]) @ R.T
    return dict(case, name=case['name']+f'_r{angle}',prev=[S[0,0],S[1,1],S[0,1],case['prev'][3]],strain=[e[0,0],e[1,1],2*e[0,1]],rotation=angle,rotation_base=case['name']+'_r0')

def fixtures():
    cases=[]; E=1000;nu=.3;c=1;phi=30
    C3,_=elastic(E,nu); s=math.sin(math.radians(phi)); low=((1-s)*12-2*c*math.cos(math.radians(phi)))/(1+s)
    for psi in (0,30):
        sp=math.sin(math.radians(psi)); m13=np.array([1-sp,0,-1-sp])
        for branch,corrected,second in [('face',[12,(12+low)/2,low],None),('edge12',[12,12,low],np.array([0,1-sp,-1-sp])),('edge23',[12,low,low],np.array([1-sp,-1-sp,0]))]:
            plastic=.001*m13
            if second is not None: plastic+=.0007*second
            trial=np.array(corrected)+C3@plastic
            for zrank,perm in enumerate(([0,1,2],[0,2,1],[1,2,0])):
                physical=trial[list(perm)]
                base=dict(name=f'{branch}_psi{psi}_z{zrank}',E=E,nu=nu,c=c,phi=phi,psi=psi,prev=[physical[0],physical[1],0.,physical[2]],strain=[0.,0.,0.])
                cases.extend(rotation(base,theta) for theta in (0,30,60))
    base=dict(E=E,nu=nu,c=c,phi=phi,psi=0,prev=[0.,0.,0.,0.])
    for name,eps in [('elastic',[1e-4,0,0]),('mixed',[.01,-.009,.004]),('shear',[0,0,.02]),('compression',[0,.01,0])]:
        case=dict(base,name=name,strain=eps);cases.extend(rotation(case,t) for t in (0,30,60))
    apex=-c/math.tan(math.radians(phi))
    cases.extend([dict(base,name='apex_isochoric',prev=[apex+5,apex-5,0.,apex],strain=[0.,0.,0.],skip_tangent=True),dict(base,name='tension_psi0_rejected',prev=[-10.,-10.,0.,-10.],strain=[0.,0.,0.],skip_tangent=True),dict(base,name='apex_associated',psi=30,prev=[-10.,-10.,0.,-10.],strain=[0.,0.,0.])])
    # Interior versus boundary of the apex flow cone: zero versus nonzero derivative.
    for label,epsp in [('apex_cone_interior',[-.004,-.004,-.004]),('apex_cone_boundary',[.005,0.,-.015])]:
        trial=C3@np.array(epsp)+apex
        item=dict(base,name=label,psi=30,prev=[trial[0],trial[1],0.,trial[2]],strain=[0.,0.,0.],expected_ready=(label=='apex_cone_interior'))
        cases.extend(rotation(item,t) for t in (0,30,60))
    return cases

def analyze(raw):
    rows=[];actual={}
    for record in raw:
        case=record['case'];v=record['v']; ref=oracle(case);ok=bool(v[0]);assert ok==ref['ok'],(case['name'],'status',ok,ref)
        row=dict(name=case['name'],ok=ok)
        if ok:
            assert bool(v[1]) and v[2]==1,(case['name'],'unexpected substeps')
            stress=np.array(v[3:7]);ep=np.array(v[7:11]);T=np.array(v[11:20]).reshape(3,3)
            if 'expected_ready' in case:
                assert bool(v[20])==case['expected_ready'],(case['name'],'strict apex cone derivative selection')
                if case['expected_ready']:assert np.max(abs(T))==0.,(case['name'],'nonzero interior apex derivative')
            gap=float(np.max(np.abs(stress-ref['stress'])));assert gap<1e-7,(case['name'],'stress',gap)
            C3,C4=elastic(case['E'],case['nu']); reconstruction=np.array(case['prev'])+C4@(np.r_[case['strain'],0.]-ep)
            rec=float(np.max(np.abs(reconstruction-stress)));assert rec<1e-7,(case['name'],'reconstruction',rec)
            epgap=float(np.max(np.abs(ep-ref['ep'])));assert epgap<1e-9,(case['name'],'flow',epgap)
            if case['psi']==0: assert abs(ep[0]+ep[1]+ep[3])<1e-10,(case['name'],'plastic volume')
            row.update(stress_gap=gap,reconstruction_gap=rec,plastic_gap=epgap)
            actual[case['name']]=stress
            if not case.get('skip_tangent'):
                J=np.zeros((3,3));h=1e-8
                for j in range(3):
                    plus=dict(case,strain=list(case['strain']));minus=dict(case,strain=list(case['strain']));plus['strain'][j]+=h;minus['strain'][j]-=h
                    rp,rm=oracle(plus),oracle(minus);assert rp['ok'] and rm['ok'],(case['name'],'FD outside domain')
                    J[:,j]=(np.array(rp['stress'][:3])-np.array(rm['stress'][:3]))/(2*h)
                terr=float(np.linalg.norm(T-J)/max(1.,np.linalg.norm(J)));assert terr<1e-5,(case['name'],'tangent',terr)
                row['tangent_relative_gap']=terr
        rows.append(row)
    for record in raw:
        case=record['case'];key=case.get('rotation_base')
        if key and case['name'] in actual:
            angle=math.radians(case['rotation']);R=np.array([[math.cos(angle),-math.sin(angle)],[math.sin(angle),math.cos(angle)]])
            b=actual[key];rot=R@np.array([[b[0],b[2]],[b[2],b[1]]])@R.T
            expected=[rot[0,0],rot[1,1],rot[0,1],b[3]];assert np.max(np.abs(actual[case['name']]-expected))<1e-7,(case['name'],'frame rotation')
    return dict(status='PASS',cases=len(rows),checks=rows)

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--generate',type=Path);p.add_argument('--analyze',type=Path);p.add_argument('--output',type=Path);args=p.parse_args()
    if args.generate:args.generate.write_bytes(json.dumps(fixtures(),ensure_ascii=False,indent=2).encode())
    if args.analyze:
        result=analyze(json.loads(args.analyze.read_text(encoding='utf-8-sig')));args.output.write_bytes(json.dumps(result,indent=2).encode());print(result['status'],result['cases'])
