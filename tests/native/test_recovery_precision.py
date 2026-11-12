"""Check optional SIMD LM against dense normal equations and true residuals."""
from pathlib import Path
import argparse,ctypes as C,json,time
import numpy as np
from scipy import sparse
ap=argparse.ArgumentParser();ap.add_argument('--dll',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args()
runtime=C.WinDLL('C:/Windows/System32/msvcr100.dll');lib=C.WinDLL(str(a.dll.resolve()));assert lib.sfm_init()==4
for name in ('sfm_regularized','sfn_regularized'):
 f=getattr(lib,name);f.argtypes=[C.POINTER(C.c_int64),C.c_double,C.c_double];f.restype=C.c_int32
rng=np.random.default_rng(20261113);records=[]
def run(A,bc,F,mode,damping,which='sfn_regularized'):
 csr=sparse.csr_matrix(A);n=A.shape[0];rp=csr.indptr.astype(np.int32);ci=csr.indices.astype(np.int32);v=csr.data;bw=int(max(abs(ci-np.repeat(np.arange(n),np.diff(rp))),default=0))*2
 D=np.full(n,314.);W=np.zeros(n*(bw+1));V=np.zeros(3*n);o=np.zeros(2)
 R=(C.c_int64*11)(n,rp.ctypes.data,ci.ctypes.data,v.ctypes.data,bc.ctypes.data,F.ctypes.data,D.ctypes.data,W.ctypes.data,bw,V.ctypes.data,o.ctypes.data)
 assert lib.sfn_set_precision(mode)==0
 t=time.monotonic();rc=getattr(lib,which)(R,damping,1e-8)
 return rc,D,o,time.monotonic()-t
for n,bw in ((12,2),(37,3),(111,7),(257,13)):
 for symmetric in (False,True):
  A=sparse.diags([rng.normal(0,.1,n-abs(k)) if k else rng.uniform(1.,2.,n) for k in range(-bw,bw+1)],range(-bw,bw+1)).toarray()
  if symmetric:A=(A+A.T)*.5
  bc=np.zeros(n,dtype=np.int32);bc[::9]=1;F=rng.normal(size=n);free=bc==0
  J=A[np.ix_(free,free)];scale=np.sqrt((J*J).sum(axis=0));B=J/scale
  for damping in (1e-8,1e-4,.25):
   H=B.T@B+damping*np.eye(J.shape[0]);g=B.T@F[free];exact=np.linalg.solve(H,g)/scale
   for mode in (1,2):
    rc,D,o,sec=run(A,bc,F,mode,damping);assert rc==1,(n,mode,rc,o)
    q=np.linalg.norm(H@(D[free]*scale)-g)/np.linalg.norm(g);err=np.linalg.norm(D[free]-exact)/max(1.,np.linalg.norm(exact))
    assert q<=1e-8 and err<1e-8 and np.all(D[~free]==0.),(q,err)
    records.append({'n':n,'bw':bw,'symmetric':symmetric,'damping':damping,'precision_mode':mode,'true_normal_residual':q,'direction_error':err,'seconds':sec})
bc=np.zeros(12,dtype=np.int32);F=np.ones(12);I=np.eye(12)
rc,D,o,sec=run(I,bc,F,0,1e-8);assert rc==-1 and np.all(D==314.)
for bad in ('nonfinite','all_fixed','zero_rhs'):
 b=bc.copy();f=F.copy()
 if bad=='nonfinite':f[3]=np.nan
 if bad=='all_fixed':b[:]=1
 if bad=='zero_rhs':f[:]=0
 rc,D,o,sec=run(I,b,f,2,1e-8);assert rc==0 and np.all(D==314.),(bad,rc)
a.out.parent.mkdir(parents=True,exist_ok=True);a.out.write_text(json.dumps({'status':'PASS','cases':records,'strict_declines_unchanged':True,'rejection_controls':True},indent=2),encoding='utf-8');print('PASS',len(records),'dense reference cases + STRICT/rejection controls')
