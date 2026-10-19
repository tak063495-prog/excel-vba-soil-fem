"""Independent true-residual checks for binary64 SIMD GMRES (no Excel)."""
import argparse, ctypes as C, pathlib, unittest, os, sys, math, json
from test_gmres_simd import grid
DLL_PATH=pathlib.Path(__file__).parents[2]/'native'/'bin'/'soilfem_native_x64.dll'
RESULTS=[]

def arr(t, xs): return (t*len(xs))(*xs)

def dense_band(n,band=7):
    rows=[];rp=[0];ci=[];vals=[];diag=[]
    for r in range(n):
        row={c:((r*13+c*7)%17-8)*.015 for c in range(max(0,r-band),min(n,r+band+1)) if c!=r}
        row[r]=3.+sum(abs(v) for v in row.values());rows.append(row)
        for c,v in sorted(row.items()):
            if c==r: diag.append(len(ci))
            ci.append(c);vals.append(v)
        rp.append(len(ci))
    ilu=vals[:]
    for r in range(n):
        positions={ci[p]:p for p in range(rp[r],rp[r+1])}
        for p in range(rp[r],diag[r]):
            c=ci[p];ilu[p]/=ilu[diag[c]]
            for q in range(diag[c]+1,rp[c+1]):
                if ci[q] in positions:ilu[positions[ci[q]]]-=ilu[p]*ilu[q]
    known=[math.sin(i*.31)+.1 for i in range(n)]
    rhs=[sum(vals[p]*known[ci[p]] for p in range(rp[r],rp[r+1])) for r in range(n)]
    return rows,rp,ci,vals,diag,ilu,known,rhs

def fixture(side, restart, max_iter=200, tol=1e-11,dense=False):
    rows,rp,ci,vals,diag,ilu,known,rhs=dense_band(side*side) if dense else grid(side); n=len(rhs)
    need=n*(2*restart+5)+(restart+1)*restart+4*restart+2
    a=[arr(C.c_int32,rp),arr(C.c_int32,ci),arr(C.c_int32,diag),arr(C.c_double,vals),arr(C.c_double,ilu),arr(C.c_double,rhs),arr(C.c_double,[987.25]*n),arr(C.c_double,[0.]*(need+1)),arr(C.c_double,[0.]*(max_iter+8)*3),arr(C.c_double,[-777.]*4)]
    g=(C.c_int64*16)(n,C.addressof(a[0]),C.addressof(a[1]),C.addressof(a[3]),C.addressof(a[4]),C.addressof(a[5]),C.addressof(a[6]),C.addressof(a[7]),restart,max_iter,C.addressof(a[8]),max_iter+8,C.addressof(a[9]),C.addressof(a[2]),len(a[7]),len(vals))
    return g,a

def snapshot(g,a): return [bytes(x) for x in a], bytes(g)

class TunedGMRES(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        dll=DLL_PATH
        if not dll.exists() or os.name!='nt': raise unittest.SkipTest('native DLL unavailable')
        cls.d=C.WinDLL(str(dll));
        for name in ('sfn_gmres','sfn_gmres_tuned'):
            f=getattr(cls.d,name); f.argtypes=[C.POINTER(C.c_int64),C.c_double]; f.restype=C.c_int32
        cls.d.sfn_set_precision.argtypes=[C.c_int32]; cls.d.sfn_set_precision.restype=C.c_int32

    def ab(self, side, restart, mode, mutate=None, max_iter=200,dense=False):
        self.assertEqual(self.d.sfn_set_precision(mode),0)
        gb,ab=fixture(side,restart,max_iter,dense=dense); gt,at=fixture(side,restart,max_iter,dense=dense)
        if mutate: mutate(gb,ab); mutate(gt,at)
        before_b=snapshot(gb,ab); before_t=snapshot(gt,at)
        rb=self.d.sfn_gmres(gb,1e-11); rt=self.d.sfn_gmres_tuned(gt,1e-11)
        self.assertEqual(rb,rt)
        residuals=[];delta=0.
        for g,a,before in ((gb,ab,before_b),(gt,at,before_t)):
            self.assertEqual(bytes(g),before[1])
            self.assertEqual([bytes(x) for x in a[:6]],before[0][:6])
            if rb==-1: self.assertEqual(snapshot(g,a),before)
            if rb==0: self.assertEqual(bytes(a[6]),before[0][6])
            if rb==1:
                rp,ci,dp,vals,ilu,rhs,x=a[:7]
                ss=sum((rhs[r]-sum(vals[p]*x[ci[p]] for p in range(rp[r],rp[r+1])))**2 for r in range(g[0]))
                norm=math.sqrt(sum(v*v for v in rhs)) or 1.
                residual=math.sqrt(ss)/norm;residuals.append(residual)
                self.assertLessEqual(residual,1e-11*1.001)
                if not mutate:
                    known=(dense_band(side*side) if dense else grid(side))[6]
                    self.assertLess(max(abs(x[i]-known[i]) for i in range(len(x))),1e-8)
        if rb==1:
            delta=max(abs(x-y) for x,y in zip(ab[6],at[6]))
            self.assertLess(delta,2e-8)
        RESULTS.append(dict(side=side,restart=restart,mode=mode,dense_band=dense,max_iterations=max_iter,baseline_rc=rb,tuned_rc=rt,
            max_abs_solution_delta=delta,true_residuals=residuals))

    def test_modes_sides_restarts_and_failures(self):
        for mode in (1,2):
            for side in (1,2,3,5,8):
                for restart in (1,2,3,5): self.ab(side,restart,mode)
                self.ab(side,min(3,side*side),mode,max_iter=1)
            self.ab(3,2,mode,lambda g,a: a[5].__setitem__(0,float('nan')))
            self.ab(3,2,mode,lambda g,a: a[3].__setitem__(0,0.0))
            self.ab(3,2,mode,lambda g,a: g.__setitem__(8,0))
            self.ab(3,2,mode,lambda g,a: g.__setitem__(14,0))
            self.ab(3,2,mode,lambda g,a: g.__setitem__(15,1))
            self.ab(3,2,mode,lambda g,a: a[2].__setitem__(0,a[0][1]))
            self.ab(3,2,mode,lambda g,a: [a[5].__setitem__(i,0.) for i in range(len(a[5]))])
            self.ab(3,2,mode,lambda g,a: [a[3].__setitem__(i,0.) for i in range(len(a[3]))])
            self.ab(3,2,mode,lambda g,a: a[4].__setitem__(a[2][0],0.))
        self.ab(3,2,0)

    def test_wide_lower_upper_triangular_rows(self):
        rows,rp,ci,vals,dp,ilu,x,rhs=dense_band(25)
        self.assertGreaterEqual(dp[12]-rp[12],4)
        self.assertGreaterEqual(rp[13]-dp[12]-1,4)
        for mode in (1,2):
            for side in (3,5,8):
                for restart in (1,3,5):self.ab(side,restart,mode,dense=True)

if __name__=='__main__':
    ap=argparse.ArgumentParser(); ap.add_argument('--dll'); ap.add_argument('--out'); args,rest=ap.parse_known_args()
    if args.dll: DLL_PATH=pathlib.Path(args.dll).resolve()
    result=unittest.main(argv=[sys.argv[0]]+rest,exit=False)
    if args.out:
        pathlib.Path(args.out).write_text(json.dumps(dict(status='PASS' if result.result.wasSuccessful() else 'FAIL',cases=RESULTS),indent=2)+'\n',encoding='utf-8')
    sys.exit(0 if result.result.wasSuccessful() else 1)
