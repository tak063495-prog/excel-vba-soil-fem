"""Compare opaque-workspace layouts while retaining exact x87 arithmetic."""
from pathlib import Path
import argparse,ctypes as C,importlib.util,json,os,subprocess,time
import numpy as np
from scipy import sparse
root=Path(__file__).resolve().parents[2]
ap=argparse.ArgumentParser();ap.add_argument('--gcc',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--reference-revision',default='26310435289f999898077de37c309520a483aabf');a=ap.parse_args();a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=True)
spec=importlib.util.spec_from_file_location('point_probe',Path(__file__).with_name('test_material_points.py'));mp=importlib.util.module_from_spec(spec);spec.loader.exec_module(mp)
probe='''
__declspec(dllexport) int32_t __stdcall sfm_probe_regularized(const int64_t *R,double damping,double tol,unsigned short cw) {
 unsigned short saved;int32_t rc;g_ready=1;
 __asm__ volatile("fnstcw %0" : "=m"(saved));__asm__ volatile("fldcw %0" :: "m"(cw));
 rc=sfm_regularized(R,damping,tol);
 __asm__ volatile("fldcw %0" :: "m"(saved));return rc;
}
'''
flags=['-O3','-std=c11','-Wall','-Wextra','-Werror','-Wno-unused-parameter','-Wno-misleading-indentation','-mfpmath=387','-fexcess-precision=standard','-ffp-contract=off','-fno-fast-math','-fno-builtin','-fno-stack-protector','-mno-stack-arg-probe','-nostdlib','-nostartfiles','-shared','-Wl,--entry=DllMain','-Wl,--kill-at','-s']
runtime=C.WinDLL('C:/Windows/System32/msvcr100.dll')
env=os.environ.copy();env['PATH']=str(a.gcc.resolve().parent)+os.pathsep+env['PATH'];dlls=[]
for name,source in [('reference',subprocess.check_output(['git','show',a.reference_revision+':native/soilfem_mat.c'],cwd=root).decode('utf-8-sig')),('candidate',(root/'native/soilfem_mat.c').read_text(encoding='utf-8-sig'))]:
    src=a.out/(name+'.c');src.write_text(source+mp.PROBE+probe,encoding='utf-8',newline='\n');dll=a.out/(name+'.dll')
    subprocess.run([str(a.gcc.resolve()),*flags,str(src),'-o',str(dll),'-lkernel32','-lgcc'],env=env,check=True)
    lib=C.WinDLL(str(dll));lib.sfm_init.restype=C.c_int32;assert lib.sfm_init()==4
    lib.sfm_probe_regularized.argtypes=[C.POINTER(C.c_int64),C.c_double,C.c_double,C.c_ushort];lib.sfm_probe_regularized.restype=C.c_int32;dlls.append(lib)
rng=np.random.default_rng(20261113);records=[]
def run(lib,A,bc,rhs,damping,cw):
    A=A.tocsr();A.sort_indices();n=A.shape[0];bw=int(max(abs(A.indices-np.repeat(np.arange(n),np.diff(A.indptr)))))
    rp=A.indptr.astype(np.int32);ci=A.indices.astype(np.int32);values=A.data.copy();d=np.full(n,3.14);workspace=np.zeros(n*(2*bw+1));vec=np.zeros(3*n);result=np.zeros(2)
    p=(C.c_int64*11)(n,rp.ctypes.data,ci.ctypes.data,values.ctypes.data,bc.ctypes.data,rhs.ctypes.data,d.ctypes.data,workspace.ctypes.data,2*bw,vec.ctypes.data,result.ctypes.data)
    start=time.perf_counter();rc=lib.sfm_probe_regularized(p,damping,1e-8,cw)
    return rc,d.tobytes(),result.tobytes(),time.perf_counter()-start
for n,bw in [(5,1),(12,2),(37,3),(111,7),(257,13),(2000,64)]:
    A=sparse.diags([rng.normal(0,.1,n-abs(k)) if k else rng.uniform(1.,2.,n) for k in range(-bw,bw+1)],range(-bw,bw+1),shape=(n,n),format='csr')
    bc=np.zeros(n,dtype=np.int32);bc[::9]=1;rhs=rng.normal(size=n)
    for cw in [0x027f,0x037f]:
      for damping in [1e-8,1e-4,.25]:
        outputs=[run(lib,A,bc,rhs,damping,cw) for lib in dlls]
        assert outputs[0][:3]==outputs[1][:3],(n,bw,cw,damping,outputs[0][0],outputs[1][0])
        assert outputs[1][0]==1,(n,bw,cw,damping)
        records.append({'n':n,'bw':bw,'cw':hex(cw),'damping':damping,'exact_output_bytes_equal':True,'old_seconds':outputs[0][3],'packed_seconds':outputs[1][3]})
for bad in ('nonfinite','all_fixed','zero_rhs'):
    n=12;A=sparse.eye(n,format='csr');bc=np.zeros(n,dtype=np.int32);rhs=np.ones(n)
    if bad=='nonfinite':rhs[3]=np.nan
    if bad=='all_fixed':bc[:]=1
    if bad=='zero_rhs':rhs[:]=0
    outputs=[run(lib,A,bc,rhs,1e-8,0x027f) for lib in dlls]
    assert outputs[0][:3]==outputs[1][:3] and outputs[0][0]==0,bad
report={'status':'PASS','exact_workspace_output_equivalence':True,'cases':records,'rejected_controls':['nonfinite','all_fixed','zero_rhs']}
(a.out/'results.json').write_text(json.dumps(report,indent=2,allow_nan=False),encoding='utf-8');print('PASS',len(records),'exact-output cases plus 3 rejection controls')
