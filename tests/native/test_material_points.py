"""Generate disposable point probes, or compare old/new x87 material updates."""
import argparse,ctypes as C,json,math,random,sys
from pathlib import Path

PROBE=r'''
/* test-only ABI; omitted from production DLL */
__declspec(dllexport) int32_t __stdcall sfm_probe_size(void) { return sizeof(mp_out); }
__declspec(dllexport) void __stdcall sfm_probe_point(const double *prev,const double *de,const double *mat,
    unsigned short cw,unsigned char *output,int32_t *result) {
  mp_mat m;mp_out o;int fast=-1;unsigned short saved;
  memcpy(&m,mat,sizeof m);memset(&o,0,sizeof o);
  __asm__ volatile("fnstcw %0" : "=m"(saved));
  __asm__ volatile("fldcw %0" :: "m"(cw));
  fp_clear();result[0]=mp_update(prev,de,&m,&o,&fast);result[1]=fast;result[2]=fp_bad();
  memcpy(output,&o,sizeof o);
  __asm__ volatile("fldcw %0" :: "m"(saved));
}
'''

def cases(count):
    rng=random.Random(20261020)
    for i in range(count):
        young=10**rng.uniform(3,8);nu=rng.uniform(.05,.48)
        phi=rng.uniform(0,.8);psi=rng.uniform(-.2,phi);coh=10**rng.uniform(-5,5)
        d00=young*(1-nu)/((1+nu)*(1-2*nu));d01=young*nu/((1+nu)*(1-2*nu));d22=young/(2*(1+nu))
        mat=[d00,d01,d22,math.sin(phi),math.cos(phi),math.sin(psi),young,nu,coh]
        prev=[rng.uniform(-2,2)*coh for _ in range(4)];de=[rng.uniform(-.02,.02) for _ in range(3)]
        if i%7==0:de=[0.,-0.,0.]
        if i%11==0:prev=[coh]*4;prev[2]=0.
        if i%13==0:mat[5]=0.
        if i%17==0:mat[3]=0.;mat[4]=1.
        if i%23==0:de=[v*1e-14 for v in de]
        if i%101==0:prev[0]=float('nan')
        if i%103==0:de[2]=float('inf')
        if i%107==0:prev[0]=1e200
        yield prev,de,mat

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--prepare');ap.add_argument('--source',default='native/soilfem_mat.c')
    ap.add_argument('--reference');ap.add_argument('--candidate');ap.add_argument('--count',type=int,default=30000);ap.add_argument('--out')
    a=ap.parse_args()
    if a.prepare:
        p=Path(a.prepare);assert not p.exists();p.parent.mkdir(parents=True,exist_ok=True)
        p.write_text(Path(a.source).read_text(encoding='utf-8-sig')+PROBE,encoding='utf-8',newline='\n');return
    assert a.reference and a.candidate and a.out and a.count>0
    runtime=C.WinDLL(str(Path('C:/Windows/System32/msvcr100.dll')))
    ds=[C.WinDLL(str(Path(p).resolve())) for p in (a.reference,a.candidate)]
    assert ds[0]._handle!=ds[1]._handle
    sizes=[]
    for d in ds:
        d.sfm_init.restype=C.c_int32;assert d.sfm_init()==4
        d.sfm_probe_size.restype=C.c_int32;sizes.append(d.sfm_probe_size())
        d.sfm_probe_point.argtypes=[C.POINTER(C.c_double)]*3+[C.c_ushort,C.POINTER(C.c_ubyte),C.POINTER(C.c_int32)]
        d.sfm_probe_point.restype=None
    assert sizes[0]==sizes[1]
    counts={};maximum=0.
    for cw in (0x027f,0x037f):
        handled=elastic=plastic=declined=bad=0
        for prev,de,mat in cases(a.count):
            args=[(C.c_double*len(v))(*v) for v in (prev,de,mat)]
            outputs=[]
            for d in ds:
                out=(C.c_ubyte*sizes[0])();result=(C.c_int32*3)()
                d.sfm_probe_point(*args,cw,out,result);outputs.append((bytes(out),tuple(result)))
            assert outputs[0][1]==outputs[1][1],(cw,prev,de,mat,outputs)
            rc,fast,flags=outputs[0][1];bad+=flags
            if rc:
                handled+=1;elastic+=fast==1;plastic+=fast==0
                assert outputs[0][0]==outputs[1][0],('point output bits',cw,prev,de,mat)
            else:declined+=1
        counts[hex(cw)]=dict(handled=handled,elastic=elastic,plastic=plastic,declined=declined,fp_bad=bad)
        assert elastic>100 and plastic>100 and declined>100
    result=dict(status='PASS',cases=a.count*2,x87_control_words=counts,all_handled_output_bytes_equal=True,
                decline_and_fp_flags_equal=True,point_output_bytes=sizes[0])
    Path(a.out).write_text(json.dumps(result,indent=2)+'\n',encoding='utf-8');print(json.dumps(result,indent=2))

if __name__=='__main__':main()
