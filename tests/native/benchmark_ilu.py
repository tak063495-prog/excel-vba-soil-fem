"""Run the disposable ILU microbenchmark DLL.

Build separately with VS x64 tools, for example:
``cl /nologo /O2 /fp:strict /MT /LD /W4 /WX /std:c11 /Fe:ilu_probe.dll ilu_probe.c /link /INCREMENTAL:NO``
"""
from __future__ import annotations
import argparse, ctypes, json, math, statistics
from pathlib import Path

ROOT=Path(__file__).resolve().parents[2]

def band(n, bw):
    rp=[0]; ci=[]; av=[]; dp=[]
    for r in range(n):
        cols=list(range(max(0,r-bw),min(n,r+bw+1)))
        if r not in cols: cols.append(r); cols.sort()
        dp.append(len(ci)+cols.index(r))
        for c in cols: ci.append(c); av.append((3.0 if c==r else (0.01 if c<r else 0.02)))
        rp.append(len(ci))
    return rp,ci,dp,av

def run(dll,n,bw,reps):
    rp,ci,dp,av=band(n,bw); inp=[1.0+0.001*i for i in range(n)]
    I32=ctypes.c_int32; D=ctypes.c_double
    p=ctypes.POINTER(I32); q=ctypes.POINTER(D)
    fn=dll.ilu_probe_run; fn.argtypes=[I32,p,p,p,q,q,I32,q,q,q,q]; fn.restype=I32
    A=(I32*len(rp))(*rp); C=(I32*len(ci))(*ci); DP=(I32*n)(*dp); V=(D*len(av))(*av); X=(D*n)(*inp)
    O=[(D*n)() for _ in range(3)]; T=(D*5)()
    if fn(n,A,C,DP,V,X,reps,*O,T)!=1: raise RuntimeError('probe failed')
    b=list(O[0]); result={'n':n,'bandwidth':bw,'reps':reps,'seconds':list(T)[:3], 'inverse_setup_seconds': T[3]}
    result['speedup_tuned']=result['seconds'][0]/result['seconds'][1]-1
    result['speedup_inverse']=result['seconds'][0]/result['seconds'][2]-1
    result['max_delta_tuned']=max(abs(a-bb) for a,bb in zip(b,O[1])); result['max_delta_inverse']=max(abs(a-bb) for a,bb in zip(b,O[2]))
    for name, vec in [('tuned', O[1]), ('inverse', O[2])]:
        if not all(math.isfinite(x) for x in vec): raise AssertionError(f'{name} nonfinite')
        y=[0.0]*n
        for r in range(n):
            y[r]=av[dp[r]]*vec[r]
            for p0 in range(dp[r]+1,rp[r+1]): y[r]+=av[p0]*vec[ci[p0]]
        z=[0.0]*n
        for r in range(n):
            z[r]=y[r]
            for p0 in range(rp[r],dp[r]): z[r]+=av[p0]*y[ci[p0]]
        rr=math.sqrt(sum((z[i]-inp[i])**2 for i in range(n)))/max(1.0,math.sqrt(sum(x*x for x in inp)))
        result[f'{name}_relative_residual']=rr
        if rr > 1e-12: raise AssertionError(f'{name} residual too high: {rr}')
    return result

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--dll',required=True); ap.add_argument('--reps',type=int,default=1000); ap.add_argument('--json',required=True); a=ap.parse_args()
    dll=ctypes.WinDLL(str(Path(a.dll).resolve())); rows=[]
    for n,bw in ((169,7),(676,20)):
        # The DLL runs variants in a fixed 0/1/2 cycle; repeat whole cycles so
        # each sample contains all variants under the same process conditions.
        samples=[run(dll,n,bw,a.reps) for _ in range(5)]
        row=dict(samples[0]); row['samples_seconds']=[s['seconds'] for s in samples]
        row['seconds']=[statistics.median(s['seconds'][v] for s in samples) for v in range(3)]
        row['variance_seconds']=[statistics.pvariance(s['seconds'][v] for s in samples) for v in range(3)]
        row['candidate2_vs_variant1_ratio']=row['seconds'][2]/row['seconds'][1]
        row['speedup_tuned']=row['seconds'][0]/row['seconds'][1]-1
        row['speedup_inverse']=row['seconds'][0]/row['seconds'][2]-1
        row['max_delta_tuned']=max(s['max_delta_tuned'] for s in samples)
        row['max_delta_inverse']=max(s['max_delta_inverse'] for s in samples)
        row['relative_residual_tuned']=max(s['tuned_relative_residual'] for s in samples)
        row['relative_residual_inverse']=max(s['inverse_relative_residual'] for s in samples)
        rows.append(row)
    Path(a.json).write_text(json.dumps({'cases':rows},indent=2)+'\n',encoding='utf-8')
    print(json.dumps({'cases':rows},indent=2))
if __name__=='__main__': main()
