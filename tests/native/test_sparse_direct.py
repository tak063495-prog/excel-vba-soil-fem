"""Independent dense, boundary, rejection and factor-ownership checks for sparse DLL."""
from pathlib import Path
import argparse, ctypes, json
import numpy as np

ap=argparse.ArgumentParser();ap.add_argument('--dll',type=Path,required=True);ap.add_argument('--report',type=Path,required=True);a=ap.parse_args()
dll=ctypes.WinDLL(str(a.dll.resolve()))
dptr=np.ctypeslib.ndpointer(dtype=np.float64,flags='C_CONTIGUOUS')
iptr=np.ctypeslib.ndpointer(dtype=np.int32,flags='C_CONTIGUOUS')
bandptr=np.ctypeslib.ndpointer(dtype=np.float64,flags='F_CONTIGUOUS')
dll.sfs_factor_band.argtypes=[ctypes.c_int,ctypes.c_int,ctypes.c_int,bandptr,iptr,ctypes.c_double,ctypes.POINTER(ctypes.c_int64),dptr]
dll.sfs_solve.argtypes=[ctypes.c_int64,ctypes.c_int,dptr,dptr,ctypes.c_double,dptr,dptr]
dll.sfs_release.argtypes=[ctypes.c_int64]
assert dll.sfs_abi()==0x10000

def factor(matrix,locked,limit=1024**3):
 n=len(matrix);band=np.zeros((n,n),order='F')
 for j in range(n):band[:n-j,j]=matrix[np.arange(n-j),np.arange(j,n)]
 token=ctypes.c_int64();metrics=np.zeros(8)
 rc=dll.sfs_factor_band(n,n-1,n,band,locked,limit,ctypes.byref(token),metrics)
 return rc,token.value,metrics

rng=np.random.default_rng(1306);cases=0;max_error=0.;max_residual=0.
for n in (1,3,10,30):
 for trial in range(8):
  g=rng.normal(size=(n,n));scale=np.geomspace(.1,100,n);matrix=(g.T@g+np.eye(n))*np.outer(scale,scale)
  locked=np.zeros(n,dtype=np.int32)
  if n>1 and trial%2:locked[::3]=1
  bc=rng.normal(size=n)*.01;rhs=rng.normal(size=n)
  free=np.flatnonzero(locked==0);fixed=np.flatnonzero(locked!=0)
  expected=bc.copy()
  if len(free):expected[free]=np.linalg.solve(matrix[np.ix_(free,free)],rhs[free]-matrix[np.ix_(free,fixed)]@bc[fixed])
  rc,token,m=factor(matrix,locked);assert rc==0,(n,trial,rc,m)
  for reuse in range(3):
   # In-place prescribed/output is the Excel call contract.
   result=bc.copy();rc=dll.sfs_solve(token,n,rhs,result,1e-10,result,m);assert rc==0,(n,trial,rc,m)
   error=np.linalg.norm(result-expected)/max(1.,np.linalg.norm(expected));assert error<1e-8,error
   max_error=max(max_error,float(error));max_residual=max(max_residual,float(m[6]));cases+=1
  assert dll.sfs_release(token)==0
  result=np.full(n,73.);assert dll.sfs_solve(token,n,rhs,bc,1e-10,result,m)==4 and np.all(result==73)

rejected=0
for matrix,rc_expected in ((np.array([[1.,1.],[1.,1.]]),3),(np.array([[1.,2.],[2.,1.]]),3),(np.diag([-1.,1.]),2),(np.diag([np.nan,1.]),2),(np.diag([np.inf,1.]),2)):
 rc,token,m=factor(matrix,np.zeros(2,dtype=np.int32));assert rc==rc_expected and token==0,(matrix,rc);rejected+=1
assert factor(np.eye(2),np.zeros(2,dtype=np.int32),limit=1)[0]==7;rejected+=1
rc,t1,m=factor(np.eye(2),np.zeros(2,dtype=np.int32));assert rc==0
rc,t2,_=factor(2*np.eye(2),np.zeros(2,dtype=np.int32));assert rc==0 and t1!=t2
for token,value in ((t1,1.),(t2,.5)):
 x=np.zeros(2);assert dll.sfs_solve(token,2,np.ones(2),np.zeros(2),1e-10,x,m)==0 and np.all(x==value)
for bad in (np.nan,np.inf,-np.inf):
 for which in ('rhs','prescribed'):
  rhs=np.ones(2);bc=np.zeros(2);x=np.full(2,73.)
  (rhs if which=='rhs' else bc)[0]=bad
  assert dll.sfs_solve(t1,2,rhs,bc,1e-10,x,m)==2 and np.all(x==73);rejected+=1
dll.sfs_release(t1)
x=np.zeros(2);assert dll.sfs_solve(t2,2,np.ones(2),np.zeros(2),1e-10,x,m)==0 and np.all(x==.5)
dll.sfs_release(t2)
record={'status':'PASS','dense_and_reused_rhs_cases':cases,'rejected_invalid_inputs':rejected,'independent_dense_max_relative_error':max_error,'true_linear_max_relative_residual':max_residual,'independent_factor_ownership':'PASS','in_place_boundary_output':'PASS'}
a.report.parent.mkdir(parents=True,exist_ok=True);a.report.write_text(json.dumps(record,indent=2,allow_nan=False),encoding='utf-8');print(json.dumps(record),flush=True)
