"""Independent sparse-grid GMRES checks; no Excel or third-party libraries."""
import argparse
import ctypes as C
import json
import math

def array(kind,values): return (kind*len(values))(*values)

def grid(side):
    rows=[]; rp=[0]; ci=[]; vals=[]; diag=[]
    for r in range(side*side):
        y,x=divmod(r,side); row={r:5.0+0.01*r}
        if x: row[r-1]=-1.1
        if x+1<side: row[r+1]=-0.9
        if y: row[r-side]=-0.8
        if y+1<side: row[r+side]=-1.2
        rows.append(row)
        for col,value in sorted(row.items()):
            if col==r: diag.append(len(ci))
            ci.append(col); vals.append(value)
        rp.append(len(ci))
    ilu=vals[:]
    for r,row in enumerate(rows):
        positions={ci[p]:p for p in range(rp[r],rp[r+1])}
        for p in range(rp[r],diag[r]):
            col=ci[p]; ilu[p]/=ilu[diag[col]]
            for q in range(diag[col]+1,rp[col+1]):
                if ci[q] in positions: ilu[positions[ci[q]]]-=ilu[p]*ilu[q]
    known=[math.sin(i*0.31)+0.1 for i in range(side*side)]
    rhs=[sum(value*known[col] for col,value in row.items()) for row in rows]
    return rows,rp,ci,vals,diag,ilu,known,rhs

def call(dll,side,restart=3,max_iter=200,tol=1e-11,mutate=None,zero=False):
    rows,rp,ci,vals,diag,ilu,known,rhs=grid(side); n=len(rhs)
    if zero: known=[0.0]*n; rhs=[0.0]*n
    need=n*(2*restart+5)+(restart+1)*restart+4*restart+2; cap=max_iter+8
    rp=array(C.c_int32,rp); ci=array(C.c_int32,ci); diag=array(C.c_int32,diag)
    vals=array(C.c_double,vals); ilu=array(C.c_double,ilu); rhs=array(C.c_double,rhs)
    out=array(C.c_double,[987.25]*n); work=array(C.c_double,[0.0]*(need+1))
    events=array(C.c_double,[0.0]*(cap*3)); info=array(C.c_double,[-777.0]*4)
    g=array(C.c_int64,[n,C.addressof(rp),C.addressof(ci),C.addressof(vals),C.addressof(ilu),C.addressof(rhs),C.addressof(out),C.addressof(work),restart,max_iter,C.addressof(events),cap,C.addressof(info),C.addressof(diag),len(work),len(vals)])
    if mutate: mutate(g,rp,ci,diag,rhs)
    before=bytes(out); info_before=bytes(info)
    rc=dll.sfn_gmres(g,tol)
    if rc==1:
        norm=math.sqrt(sum(value*value for value in rhs)) or 1.0
        residual=math.sqrt(sum((sum(value*out[col] for col,value in row.items())-rhs[i])**2 for i,row in enumerate(rows)))/norm
        assert residual<=tol*1.001,(rc,residual,list(info))
        assert max(abs(a-b) for a,b in zip(out,known))<max(1e-9,10*tol)
        assert 0<=info[1]<=max_iter and info[3]==0
        return dict(relative_residual=residual,iterations=int(info[1]),events=int(info[2]))
    assert bytes(out)==before, 'Rejected/failed solve wrote solution'
    if rc==-1: assert bytes(info)==info_before,'Invalid preflight changed diagnostics'
    return dict(rc=rc,iterations=int(info[1]),events=int(info[2]))

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--dll',required=True); ap.add_argument('--out'); args=ap.parse_args()
    d=C.WinDLL(args.dll)
    d.sfn_set_precision.argtypes=[C.c_int32]; d.sfn_set_precision.restype=C.c_int32
    d.sfn_gmres.argtypes=[C.POINTER(C.c_int64),C.c_double]; d.sfn_gmres.restype=C.c_int32
    records=[]
    for mode in (1,2):
        assert d.sfn_set_precision(mode)==0
        for side in (1,2,3,5,8):
            records.append(dict(mode=mode,side=side,**call(d,side,restart=min(3,side*side))))
        assert call(d,8,restart=2,max_iter=1)['rc']==0
        assert call(d,3,zero=True)['iterations']==0
        for tolerance in (1e-8,1e-10,1e-12):
            records.append(dict(mode=mode,side=8,tolerance=tolerance,**call(d,8,restart=2,tol=tolerance)))
        def nonfinite(g,rp,ci,dp,rhs): rhs[0]=float('nan')
        assert call(d,3,mutate=nonfinite)['rc']==0
        def singular_preconditioner(g,rp,ci,dp,rhs):
            C.c_double.from_address(g[4]+8*dp[0]).value=0.0
        assert call(d,3,mutate=singular_preconditioner)['rc']==0
        def zero_operator(g,rp,ci,dp,rhs):
            for i in range(g[15]): C.c_double.from_address(g[3]+8*i).value=0.0
        assert call(d,3,mutate=zero_operator)['rc']==0
        for name,value in (('m',0),('work',0),('nnz',1),('huge_n',2147483647),('event_cap',1)):
            index={'m':8,'work':14,'nnz':15,'huge_n':0,'event_cap':11}[name]
            def bad(g,rp,ci,dp,rhs,index=index,value=value): g[index]=value
            assert call(d,3,mutate=bad)['rc']==-1
        def bad_diagonal(g,rp,ci,dp,rhs): dp[0]=rp[1]
        assert call(d,3,mutate=bad_diagonal)['rc']==-1
    d.sfn_set_precision(0)
    assert call(d,3)['rc']==-1, 'Strict must delegate to x87 reference'
    text=json.dumps(dict(status='PASS',solutions=records,checked_failure_transactionality=True),indent=2)
    print(text)
    if args.out:
        with open(args.out,'w',encoding='utf-8') as stream: stream.write(text+'\n')

if __name__=='__main__': main()
