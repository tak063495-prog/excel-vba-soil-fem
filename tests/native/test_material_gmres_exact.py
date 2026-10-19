"""O2/O3 material-DLL x87 GMRES: all buffers exact; independent residual on success."""
import argparse,ctypes as C,hashlib,json,math
from pathlib import Path
from test_gmres_tuned import fixture
from test_element_exact import load

def solve(d,side,restart,dense,variant):
    g,a=fixture(side,min(restart,side*side),max_iter=1 if variant=='limit' else 200,dense=dense)
    if variant=='zero_rhs':
        for i in range(len(a[5])):a[5][i]=0.
    elif variant=='nan_rhs':a[5][0]=float('nan')
    elif variant=='zero_operator':
        for i in range(len(a[3])):a[3][i]=0.
    elif variant=='pivot':a[4][a[2][0]]=0.
    before=bytes(a[6]);rc=d.sfm_gmres(g,1e-11);res=None
    if rc==1:
        rp,ci,dp,vals,ilu,rhs,x=a[:7]
        nb=math.sqrt(sum(v*v for v in rhs)) or 1.
        res=math.sqrt(sum((rhs[r]-sum(vals[p]*x[ci[p]] for p in range(rp[r],rp[r+1])))**2 for r in range(g[0])))/nb
        assert math.isfinite(res) and res<=1.001e-11,res
    else:assert bytes(a[6])==before,'Failed GMRES changed solution'
    return (rc,[bytes(x) for x in a]),res

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--reference',required=True);p.add_argument('--candidate',required=True);p.add_argument('--out',required=True);a=p.parse_args()
    ds=[load(v) for v in (a.reference,a.candidate)]
    for d in ds:d.sfm_gmres.argtypes=[C.POINTER(C.c_int64),C.c_double];d.sfm_gmres.restype=C.c_int32
    words=(0x027f,0x037f) if all(hasattr(d,'sfm_probe_cw') for d in ds) else (None,)
    rows=[];saved=None
    try:
        for cw in words:
            if cw is not None:
                old=ds[0].sfm_probe_cw(cw)
                if saved is None:saved=old
            for dense in (False,True):
                for side in (1,2,3,5,8):
                    for restart in (1,2,3,5):
                        for variant in ('normal','limit','zero_rhs','nan_rhs','zero_operator','pivot'):
                            r,res=solve(ds[0],side,restart,dense,variant);c,cres=solve(ds[1],side,restart,dense,variant)
                            assert r==c,(cw,dense,side,restart,variant,'GMRES buffer bits')
                            rows.append(dict(cw=cw,dense=dense,side=side,restart=restart,variant=variant,rc=r[0],true_residual=res))
    finally:
        if saved is not None:ds[0].sfm_probe_cw(saved)
    assert any(r['rc']==1 for r in rows) and any(r['rc']==0 for r in rows)
    result=dict(status='PASS',cases=len(rows),all_buffer_bytes_equal=True,fp_control_restored=True,
        x87_control_words=words,reference_sha256=hashlib.sha256(Path(a.reference).read_bytes()).hexdigest(),
        candidate_sha256=hashlib.sha256(Path(a.candidate).read_bytes()).hexdigest(),results=rows)
    Path(a.out).write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8');print('PASS',len(rows),'material GMRES cases')
