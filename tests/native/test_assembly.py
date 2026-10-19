"""Independent binary64/nonzero ordering, layout and rejection checks."""
import ctypes as C
import os
from pathlib import Path
import random
import unittest


class Fixture:
    def __init__(self, n=2, style='sparse', seed=17):
        self.n=n; self.stride=2112; self.offset=32; self.nnz=17
        rng=random.Random(seed)
        self.blob=(C.c_ubyte*(n*self.stride))()
        self.ep=(C.c_int32*(n*256))()
        self.expected=[]; self.starts=[0]
        for k in range(n):
            for i in range(16):
                for j in range(16):
                    v=rng.choice([0.,-0.,5e-324,-1e-300,1e300,-3.25])
                    if style=='zero': v=-0.
                    if style=='dense': v=float(1+k*256+i*16+j)
                    C.c_double.from_buffer(self.blob,k*self.stride+self.offset+8*(i+16*j)).value=v
                    p=(k*7+i*3+j)%self.nnz
                    self.ep[k+n*(i+16*j)]=p
                    if v!=0.: self.expected.append((i,j,p,k*256+i*16+j))
            self.starts.append(len(self.expected))
        self.cap=len(self.expected)+5
        self.start=(C.c_int32*(n+1))(*([-71]*(n+1)))
        self.outputs=[(C.c_int32*self.cap)(*([-91]*self.cap)) for _ in range(4)]
        self.out=C.c_int32(-81)
        self.p=(C.c_int64*13)(C.addressof(self.blob),self.stride,n,self.offset,C.addressof(self.ep),self.nnz,
            C.addressof(self.start),self.cap,*[C.addressof(x) for x in self.outputs],C.addressof(self.out))

    def output_bytes(self):
        return [bytes(self.start),*[bytes(x) for x in self.outputs],bytes(self.out)]


class AssemblyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if os.name!='nt': raise unittest.SkipTest('Windows x64 DLL')
        path=Path(__file__).resolve().parents[2]/'native/bin/soilfem_native_x64.dll'
        cls.dll=C.WinDLL(str(path))
        cls.fn=cls.dll.sfn_csr_scatter_list
        cls.fn.argtypes=[C.POINTER(C.c_int64)]; cls.fn.restype=C.c_int32

    def test_order_all_values_and_input_immutability(self):
        for n in (1,2,7,20):
            for style in ('zero','sparse','dense'):
                for seed in (17,291):
                    with self.subTest(n=n,style=style,seed=seed):
                        f=Fixture(n,style,seed)
                        before=(bytes(f.blob),bytes(f.ep),bytes(f.p))
                        # Count-only permits absent row/col/pos/valueIndex pointers.
                        f.p[7]=-1
                        for q in range(8,12): f.p[q]=0
                        self.assertEqual(self.fn(f.p),0)
                        self.assertEqual(list(f.start),f.starts)
                        self.assertEqual(f.out.value,len(f.expected))
                        for x in f.outputs: self.assertEqual(list(x),[-91]*f.cap)
                        f.p[7]=f.cap
                        for q,x in enumerate(f.outputs,8): f.p[q]=C.addressof(x)
                        self.assertEqual(self.fn(f.p),0)
                        self.assertEqual(list(f.start),f.starts)
                        self.assertEqual(list(zip(*[list(x)[:f.out.value] for x in f.outputs])),f.expected)
                        for x in f.outputs: self.assertEqual(list(x)[f.out.value:],[-91]*5)
                        self.assertEqual((bytes(f.blob),bytes(f.ep),bytes(f.p)),before)

    def test_decline_before_any_output_write(self):
        changes=[(0,0),(1,0),(1,2047),(1,2**63-1),(2,0),(2,8388608),(3,-1),
            (3,65),(4,0),(5,0),(5,2**31),(6,0),(7,-2),(7,0),(7,2**31),
            (8,0),(9,0),(10,0),(11,0),(12,0),(0,2**63-100),
            (4,2**63-100),(6,2**63-2),(8,2**63-2),(12,2**63-2)]
        self.assertEqual(self.fn(None),-1)
        for index,value in changes:
            with self.subTest(index=index,value=value):
                f=Fixture(style='dense'); f.p[index]=value
                before=f.output_bytes(); inputs=(bytes(f.blob),bytes(f.ep),bytes(f.p))
                self.assertEqual(self.fn(f.p),-1)
                self.assertEqual(f.output_bytes(),before)
                self.assertEqual((bytes(f.blob),bytes(f.ep),bytes(f.p)),inputs)

    def test_nonfinite_and_invalid_positions_both_passes(self):
        for cap in (-1,600):
            for v in (float('nan'),float('inf'),-float('inf')):
                f=Fixture(style='dense'); f.p[7]=cap
                C.c_double.from_buffer(f.blob,f.stride+f.offset+2040).value=v
                before=f.output_bytes()
                self.assertEqual(self.fn(f.p),-1); self.assertEqual(f.output_bytes(),before)
            for p in (-1,17):
                f=Fixture(style='dense'); f.p[7]=cap; f.ep[511]=p
                before=f.output_bytes()
                self.assertEqual(self.fn(f.p),-1); self.assertEqual(f.output_bytes(),before)


if __name__=='__main__': unittest.main()
