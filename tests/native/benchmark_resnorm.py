"""Build probe with cl /LD /O2 /fp:strict /W4 /WX, then run two CSR shapes."""
import argparse,ctypes as C,hashlib,json,pathlib,statistics,math
def main():
 p=argparse.ArgumentParser();p.add_argument('--dll',required=True);p.add_argument('--out',required=True);p.add_argument('--reps',type=int,default=1000);a=p.parse_args();d=C.WinDLL(str(pathlib.Path(a.dll).resolve()));d.resnorm_check.argtypes=[C.c_int,C.c_int];d.resnorm_check.restype=C.c_double;d.resnorm_bench.argtypes=[C.c_int,C.c_int,C.c_int,C.c_int];d.resnorm_bench.restype=C.c_double
 out={'dll':str(a.dll),'dll_sha256':hashlib.sha256(pathlib.Path(a.dll).read_bytes()).hexdigest(),'source_sha256':hashlib.sha256(pathlib.Path(__file__).with_name('resnorm_probe.c').read_bytes()).hexdigest(),'build_flags':['/O2','/fp:strict','/W4','/WX','/LD'],'reps':a.reps,'shapes':{}}
 for n,bw in ((169,8),(676,20)):
  assert d.resnorm_check(n,bw)==1.0; t=[[],[]]
  for rep in range(5):
   for mode in ((0,1) if rep%2==0 else (1,0)):
    value=d.resnorm_bench(mode,n,bw,a.reps);assert math.isfinite(value) and value>0;t[mode].append(value)
  old,new=statistics.median(t[0]),statistics.median(t[1]);out['shapes'][str(n)]={'bw':bw,'old_samples_s':t[0],'fused_samples_s':t[1],'old_median_s':old,'fused_median_s':new,'reduction_pct':(old-new)/old*100}
 pathlib.Path(a.out).parent.mkdir(parents=True,exist_ok=True);pathlib.Path(a.out).write_text(json.dumps(out,indent=2)+'\n');print(json.dumps(out,indent=2))
if __name__=='__main__':main()
