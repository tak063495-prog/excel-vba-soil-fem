#!/usr/bin/env python3
"""ctypes regression and small benchmark for soilfem_native precision modes."""
import argparse, ctypes as C, hashlib, json, math, os, statistics, time
from concurrent.futures import ThreadPoolExecutor

def lib(path):
    d = C.WinDLL(path)
    i = C.c_int32; p = C.POINTER
    d.sfn_band_factor.argtypes=[i,i,i,p(C.c_double),p(C.c_double),p(C.c_double),C.c_double,p(i)]; d.sfn_band_factor.restype=i
    d.sfn_band_solve.argtypes=[i,i,p(C.c_double),p(C.c_double),p(C.c_double),C.c_double,p(i)]; d.sfn_band_solve.restype=i
    d.sfn_nsband_factor.argtypes=[i,i,i,p(C.c_double),p(C.c_double),p(i),p(C.c_double),p(i),p(i),p(C.c_double)]; d.sfn_nsband_factor.restype=i
    d.sfn_nsband_forward.argtypes=[i,i,i,p(i),p(C.c_double),p(C.c_double)]; d.sfn_nsband_forward.restype=i
    d.sfn_nsband_back.argtypes=[i,i,i,p(C.c_double),p(C.c_double),p(C.c_double),i,p(i)]; d.sfn_nsband_back.restype=i
    for n in ('sfn_set_precision','sfn_get_precision','sfn_cpu_features'):
        if hasattr(d,n): getattr(d,n).argtypes=[C.c_int32] if n=='sfn_set_precision' else []; getattr(d,n).restype=C.c_int32
    return d

def arr(xs): return (C.c_double*len(xs))(*xs)
def residual(A, x, b):
    return max(abs(sum(row[j]*x[j] for j in range(len(x)))-rowb) for row,rowb in zip(A,b))

def symmetric(d, n=31, bw=4, return_state=False, ld_pad=0):
    band=[0.0]*(n*(bw+1)); A=[[0.0]*n for _ in range(n)]
    for i in range(n):
        A[i][i]=3.0 + 0.01*i
        for j in range(max(0,i-bw),i): A[i][j]=A[j][i]=0.07/(1+abs(i-j))
    for i in range(n):
        for j in range(bw+1): band[i+j*n]=A[i][i+j] if i+j<n else 0.0
    b=[math.sin(i*0.37)+0.2 for i in range(n)]; packed=arr([0.0]*(n*(bw+1))); akj=arr([0.0]*(bw+1)); rhs=arr(b); x=arr([0.0]*n); fail=C.c_int32(); ld=n+ld_pad; raw=[0.0]*(ld*(bw+1))
    for i in range(n):
        for j in range(bw+1): raw[i+j*ld]=band[i+j*n]
    rc=d.sfn_band_factor(n,bw,ld,arr(raw),packed,akj,1e-30,C.byref(fail)); assert rc==0, (rc,fail.value)
    rc=d.sfn_band_solve(n,bw,packed,rhs,x,1e-30,C.byref(fail)); assert rc==0
    err=residual(A,list(x),b)
    return (err, list(x), bytes(packed), bytes(rhs), bytes(x)) if return_state else (err, list(x))

def nonsym(d, n=29, b=3, return_state=False):
    cols=3*b+1; a=[0.0]*(n*cols); A=[[0.0]*n for _ in range(n)]
    for i in range(n):
        A[i][i]=4.0+0.01*i
        for j in range(max(0,i-b),min(n,i+b+1)):
            if j!=i: A[i][j]=0.03*(1+i+j)/(1+abs(i-j))
            a[i+(j-i+b)*n]=A[i][j]
    A[0][0]=0.001; A[1][0]=1.0; a[0+b*n]=A[0][0]; a[1+(b-1)*n]=A[1][0]
    rhs0=[0.4+math.cos(i*0.21) for i in range(n)]; aa=arr(a); rhs=arr(rhs0); piv=(C.c_int32*n)(); low=arr([0.0]*(n*(b+1))); fk=C.c_int32(); reason=C.c_int32(); largest=C.c_double()
    assert d.sfn_nsband_factor(n,b,n,aa,rhs,piv,low,C.byref(fk),C.byref(reason),C.byref(largest))==0
    sol=arr([0.0]*n); assert d.sfn_nsband_back(n,b,n,aa,rhs,sol,1,C.byref(fk))==0
    err=residual(A,list(sol),rhs0)
    fresh=arr(rhs0); assert d.sfn_nsband_forward(n,b,n,piv,low,fresh)==0
    sol2=arr([0.0]*n); assert d.sfn_nsband_back(n,b,n,aa,fresh,sol2,1,C.byref(fk))==0
    assert any(piv[k] != k for k in range(n))
    return (err, max(abs(x-y) for x,y in zip(sol,sol2)), bytes(aa), bytes(rhs), bytes(sol), bytes(low), bytes(piv)) if return_state else (err, max(abs(x-y) for x,y in zip(sol,sol2)))

def benchmark(d, n=600, bw=64, reps=5):
    band=[0.0]*(n*(bw+1))
    for i in range(n):
        band[i]=5.0
        for j in range(1,min(bw+1,n-i)): band[i+j*n]=0.001/(1+j)
    times=[]; packed=arr([0.0]*(n*(bw+1))); work=arr([0.0]*(bw+1)); src=arr(band)
    for k in range(reps+1):
        fk=C.c_int32(); t=time.perf_counter(); rc=d.sfn_band_factor(n,bw,n,src,packed,work,1e-30,C.byref(fk)); dt=(time.perf_counter()-t)*1000
        if rc: raise RuntimeError(rc)
        if k: times.append(dt)
    return statistics.median(times)

def benchmark_ns(d, n=600, b=64, reps=5):
    raw=[0.0]*(n*(3*b+1))
    for i in range(n):
        for j in range(max(0,i-b),min(n,i+b+1)):
            raw[i+(j-i+b)*n]=5.0 if i==j else 0.001/(1+abs(i-j))
    times=[]
    for k in range(reps+1):
        a=arr(raw); rhs=arr([1.0]*n); piv=(C.c_int32*n)(); low=arr([0.0]*(n*(b+1)))
        fk=C.c_int32(); reason=C.c_int32(); largest=C.c_double()
        t=time.perf_counter()
        rc=d.sfn_nsband_factor(n,b,n,a,rhs,piv,low,C.byref(fk),C.byref(reason),C.byref(largest))
        elapsed=(time.perf_counter()-t)*1000
        assert rc==0
        if k: times.append(elapsed)
    return statistics.median(times)

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--dll',required=True); ap.add_argument('--baseline'); ap.add_argument('--out'); a=ap.parse_args()
    d=lib(a.dll); modes={}
    assert all(hasattr(d,name) for name in ('sfn_set_precision','sfn_get_precision','sfn_cpu_features')), 'Updated precision API required'
    assert d.sfn_get_precision()==0, 'New thread defaults to strict'
    assert d.sfn_set_precision(1)==0
    def other_thread():
        assert d.sfn_get_precision()==0
        assert d.sfn_set_precision(0)==0
    with ThreadPoolExecutor(max_workers=1) as pool:
        pool.submit(other_thread).result()
    assert d.sfn_get_precision()==1, 'Thread selection leaked'
    d.sfn_set_precision(0)
    if hasattr(d,'sfn_set_precision'):
        before=d.sfn_get_precision(); assert d.sfn_set_precision(-1)==100 and d.sfn_set_precision(3)==100 and d.sfn_get_precision()==before
    for mode in (0,1,2):
        if hasattr(d,'sfn_set_precision') and d.sfn_set_precision(mode)!=0: continue
        tail=max(symmetric(d,n,min(4,n-1))[0] for n in (1,2,3,4,5,7,8,9,29,31)); assert tail<=1e-11
        assert symmetric(d,17,3,ld_pad=3)[0] <= 1e-11
        modes[str(mode)]={'sym_residual':symmetric(d)[0], 'ns_residual':nonsym(d)[0], 'ns_replay_delta':nonsym(d)[1], 'tail_max':tail, 'benchmark_factor_ms':benchmark(d), 'benchmark_ns_factor_ms':benchmark_ns(d)}
    # Exercise short lengths/tails and explicit rejection paths.
    bad=C.c_int32(-1); assert d.sfn_band_factor(0,0,0,None,None,None,1e-30,C.byref(bad))==100
    singular=arr([0.0]*4); assert d.sfn_band_factor(2,1,2,singular,arr([0.0]*4),arr([0.0]*2),1e-30,C.byref(bad))!=0
    feat=int(d.sfn_cpu_features()) if hasattr(d,'sfn_cpu_features') else 0
    if hasattr(d,'sfn_set_precision'):
        d.sfn_set_precision(2); assert d.sfn_get_precision()==(2 if feat&2 else 1)
    result={'dll':os.path.abspath(a.dll),'dll_sha256':hashlib.sha256(open(a.dll,'rb').read()).hexdigest(),'cpu_features':feat,'benchmark_n':600,'benchmark_band':64,'benchmark_reps':5,'modes':modes}
    if a.baseline:
        bd=lib(a.baseline); result['baseline']=os.path.abspath(a.baseline)
        if hasattr(d,'sfn_set_precision'): assert d.sfn_set_precision(0)==0
        if hasattr(bd,'sfn_set_precision'): assert bd.sfn_set_precision(0)==0
        cur=symmetric(d,return_state=True); ref=symmetric(bd,return_state=True); result['strict_solution_max_delta']=max(abs(x-y) for x,y in zip(cur[1],ref[1])); assert cur[1:]==ref[1:]
        assert nonsym(d,return_state=True)[2:]==nonsym(bd,return_state=True)[2:]
        result['strict_factor_rhs_solution_bytes_equal']=True
        result['baseline_factor_ms']=benchmark(bd)
        result['baseline_ns_factor_ms']=benchmark_ns(bd)
    text=json.dumps(result,indent=2); print(text)
    if a.out: open(a.out,'w',encoding='utf-8').write(text+'\n')
    assert all(v['sym_residual']<1e-11 and v['ns_residual']<1e-11 and v['ns_replay_delta']<1e-11 for v in modes.values())
if __name__=='__main__': main()
