"""Direct sfm_eval/sfm_tangent byte comparison for material element kernels."""
import argparse, ctypes as C, hashlib, json, pathlib, random

CW_PROBE=r'''
/* Disposable test-only ABI; never ship this export. */
__declspec(dllexport) unsigned short __stdcall sfm_probe_cw(unsigned short word) {
 unsigned short old; __asm__ volatile("fnstcw %0" : "=m"(old));
 __asm__ volatile("fldcw %0" :: "m"(word)); return old;
}
'''

O_STRIDE, O_U, O_MATNO, O_ELNODE, O_BMAT, O_SMAT, O_SPMAT, O_STMAT, O_MSTMAT, O_PLSTRAIN, O_PLMULT, O_YIELDF, O_YIELDED, O_PRSTRESS, O_LASTSUB, O_STREAK, O_DPMAT, O_KMAT, O_DJ, O_ELMWEIGHT, O_HGCOUNT, O_HGSCALE, O_HGMODE, O_ISJOINT, O_SPVALID, O_TDIRTY = range(26)

def load(p):
    C.WinDLL('C:/Windows/System32/msvcr100.dll')
    d=C.WinDLL(str(pathlib.Path(p).resolve())); d.sfm_init.restype=C.c_int32
    if d.sfm_init()!=4: raise RuntimeError('sfm_init failed')
    for n in ('sfm_eval','sfm_tangent'):
        f=getattr(d,n); f.restype=C.c_int32
    d.sfm_eval.argtypes=[C.POINTER(C.c_int64),C.c_int32]
    d.sfm_tangent.argtypes=[C.POINTER(C.c_int64),C.POINTER(C.c_int16)]
    if hasattr(d,'sfm_probe_cw'):
        d.sfm_probe_cw.argtypes=[C.c_ushort];d.sfm_probe_cw.restype=C.c_ushort
    return d

def fixture(flags, seed, variant):
    rng=random.Random(seed); stride=20000; buf=(C.c_ubyte*(stride*2))()
    off=[0]*26; pos=256
    sizes={O_U:16*8,O_MATNO:4,O_ELNODE:16*4,O_BMAT:3*16*4*8,O_SMAT:3*16*4*8,O_SPMAT:3*16*4*8,O_STMAT:17*4*8,O_MSTMAT:17*4*8,O_PLSTRAIN:4*4*8,O_PLMULT:4*8,O_YIELDF:4*8,O_YIELDED:4*2,O_PRSTRESS:3*4*8,O_LASTSUB:4*4,O_STREAK:4*4,O_DPMAT:3*3*4*8,O_KMAT:256*8,O_DJ:4*8,O_ELMWEIGHT:8,O_HGCOUNT:4,O_HGSCALE:8,O_HGMODE:4*16*8,O_ISJOINT:2,O_SPVALID:4*2,O_TDIRTY:2}
    for k in range(1,26): off[k]=pos; pos += sizes.get(k,8)+32
    off[O_STRIDE]=stride
    def ptr(t, at): return C.cast(C.byref(buf,at),C.POINTER(t))
    en=ptr(C.c_int32,off[O_ELNODE]); [en.__setitem__(i,i) for i in range(16)]
    B=ptr(C.c_double,off[O_BMAT]); S=ptr(C.c_double,off[O_SMAT]); Sp=ptr(C.c_double,off[O_SPMAT]); dj=ptr(C.c_double,off[O_DJ]);
    for i in range(3*16*4): B[i]=rng.uniform(-.2,.2); S[i]=rng.uniform(-100,100); Sp[i]=rng.uniform(-100,100)
    for im in range(4): dj[im]=1.0; ptr(C.c_int16,off[O_SPVALID])[im]=-1 if im%2 else 0
    for i in range(3): S[i+3*(i+16*0)]=1.0; S[i+3*(i+16*1)]=1.0; S[i+3*(i+16*2)]=1.0; S[i+3*(i+16*3)]=1.0
    ptr(C.c_int32,off[O_MATNO])[0]=0; ptr(C.c_int16,off[O_TDIRTY])[0]=1
    mt=(C.c_double*16)(1,1e4,3e3,4e3,.4794255,.8775826,.0998334,1e4,.3,100,1,1,0,0,0,0)
    td=(C.c_double*32)(*[rng.uniform(-.01,.01) for _ in range(32)]); inc=(C.c_double*32)(*[rng.uniform(-.001,.001) for _ in range(32)]); inf=(C.c_double*32)(*[0.0]*32)
    # Initialize the second element too, sharing the first element's DOFs.
    C.memmove(C.byref(buf,stride),C.byref(buf),stride)
    for e in (0,stride):
        for im in range(4):
            ptr(C.c_int32,e+off[O_LASTSUB])[im]=1
            for j,v in zip((0,1,2,16),(50.,40.,1.,0.)):
                ptr(C.c_double,e+off[O_MSTMAT])[j+17*im]=v
    active=(C.c_int16*2)(-1,-1)
    if variant=='zero':
        for j in range(32):td[j]=inc[j]
    elif variant=='joint':ptr(C.c_int16,stride+off[O_ISJOINT])[0]=-1
    elif variant=='material':ptr(C.c_int32,stride+off[O_MATNO])[0]=1
    elif variant=='substep':ptr(C.c_int32,stride+off[O_LASTSUB])[2]=2
    elif variant=='nan':ptr(C.c_double,stride+off[O_MSTMAT])[0]=float('nan')
    elif variant=='overflow':
        for j in range(192):ptr(C.c_double,stride+off[O_BMAT])[j]=1e308
    elif variant=='dof':ptr(C.c_int32,stride+off[O_ELNODE])[5]=32
    elif variant=='inactive':active[1]=0
    elif variant=='skipdof':
        for e in (0,stride):ptr(C.c_int32,e+off[O_ELNODE])[5]=-1
    elif variant=='hg':
        for e in (0,stride):
            ptr(C.c_int32,e+off[O_HGCOUNT])[0]=2
            ptr(C.c_double,e+off[O_HGSCALE])[0]=.25
            for j in range(64):ptr(C.c_double,e+off[O_HGMODE])[j]=rng.uniform(-1,1)
    elif variant=='clean':
        for e in (0,stride):ptr(C.c_int16,e+off[O_TDIRTY])[0]=0
    elif variant=='plastic':
        for j in range(32):td[j]*=100
    out=(C.c_int64*16)();out[8]=-1
    app=(C.c_int16*8)(0,-1,0,-1,0,-1,0,-1); gn=(C.c_double*32)(*[1.0]*32)
    offarr=(C.c_int64*26)(*off); P=(C.c_int64*16)(); P[0]=C.addressof(buf); P[1]=C.addressof(offarr); P[2]=2; P[3]=C.addressof(td); P[4]=C.addressof(inc); P[5]=C.addressof(inf); P[6]=0; P[9]=C.addressof(mt); P[10]=1; P[11]=C.addressof(app); P[12]=flags; P[13]=C.addressof(out); P[14]=C.addressof(gn); P[15]=31
    # Keep offsets alive through the calls.
    P[6]=C.addressof(active);P[7]=0;P[8]=1
    P._keep=(buf,offarr,td,inc,inf,mt,out,app,gn,active)
    return P,buf,out,inf,app

def snap(d,flags,seed,variant,tangent_only):
    P,b,o,inf,app=fixture(flags,seed,variant); done=(C.c_int16*2)(-9,-9)
    r1=-99 if tangent_only else d.sfm_eval(P,0)
    before=(bytes(b),bytes(inf),bytes(app),bytes(o))
    r2=d.sfm_tangent(P,done)
    return r1,r2,before,bytes(b),bytes(inf),bytes(app),bytes(o),bytes(done)

def main():
    a=argparse.ArgumentParser();a.add_argument('--prepare');a.add_argument('--source');a.add_argument('--reference');a.add_argument('--candidate');a.add_argument('--count',type=int,default=100);a.add_argument('--out');x=a.parse_args()
    if x.prepare:
        p=pathlib.Path(x.prepare);assert not p.exists();p.parent.mkdir(parents=True,exist_ok=True)
        p.write_text(pathlib.Path(x.source).read_text(encoding='utf-8-sig')+CW_PROBE,encoding='utf-8',newline='\n');return
    r,c=load(x.reference),load(x.candidate);assert r._handle!=c._handle
    words=(0x027f,0x037f) if hasattr(r,'sfm_probe_cw') and hasattr(c,'sfm_probe_cw') else (None,)
    variants=('normal','zero','joint','material','substep','nan','overflow','dof','inactive','skipdof','hg','clean','plastic')
    rows=0;statuses={};old=None
    try:
        for cw in words:
            if cw is not None:
                saved=r.sfm_probe_cw(cw)
                if old is None:old=saved
            for i in range(x.count):
                flags=(0,1,2,3,4,8,16,18,19,20,24)[i%11]
                for variant in variants:
                    for tangent_only in (False,True):
                        A=snap(r,flags,20261020+i*31,variant,tangent_only); B=snap(c,flags,20261020+i*31,variant,tangent_only)
                        assert A==B,(cw,i,flags,variant,tangent_only,'element/force/app/out/done bytes')
                        statuses[str(A[:2])]=statuses.get(str(A[:2]),0)+1;rows+=1
    finally:
        if old is not None:r.sfm_probe_cw(old)
    assert any(k.startswith('(2,') for k in statuses) and any(k.startswith('(1,') for k in statuses)
    report={'status':'PASS','reference':str(x.reference),'candidate':str(x.candidate),'reference_sha256':hashlib.sha256(pathlib.Path(x.reference).read_bytes()).hexdigest(),'candidate_sha256':hashlib.sha256(pathlib.Path(x.candidate).read_bytes()).hexdigest(),'cases':rows,'matches':rows,'x87_control_words':words,'statuses':statuses,'complete_output_bytes_equal':True}
    pathlib.Path(x.out).parent.mkdir(parents=True,exist_ok=True); pathlib.Path(x.out).write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8'); print(json.dumps(report,indent=2))
if __name__=='__main__': main()
