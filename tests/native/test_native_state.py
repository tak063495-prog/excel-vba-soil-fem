"""Exact byte and transaction tests; does not start or touch Excel."""
import argparse
import ctypes as C
import json
import random
import struct

SIZES=(544,128,32,32,8,96,8,1536,288)
STRIDE=4096

def fixture(n):
    rng=random.Random(706+n)
    dest=C.create_string_buffer(b'\xa5'*(n*STRIDE+64))
    offsets=[]; offset=32
    for size in SIZES:
        offsets.append(offset); offset+=size+16
    assert offset<STRIDE
    source=[C.create_string_buffer(rng.randbytes((n*STRIDE if i==0 else n*size)+64)) for i,size in enumerate(SIZES)]
    dp=(C.c_int64*9)(*(C.addressof(dest)+offset for offset in offsets))
    sp=(C.c_int64*9)(*(C.addressof(buf)+32 for buf in source))
    yielded=(C.c_int16*(n*4)).from_address(sp[4])
    for i in range(n*4): yielded[i]=-1 if i%3 else 0
    # Specific double payloads include negative zero and a quiet NaN.
    C.memmove(sp[1],struct.pack('<QQ',0x8000000000000000,0x7ff8000000001234),16)
    plastic=(C.c_int16*(n*4))(*([-1,0,-1,0]*n))
    count=C.c_int32(sum(v!=0 for v in plastic))
    return dest,source,dp,sp,plastic,count,offsets

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--dll',required=True); args=ap.parse_args()
    dll=C.WinDLL(args.dll); f=dll.sfn_restore_committed
    f.argtypes=[C.c_int32,C.c_int64,C.POINTER(C.c_int64),C.POINTER(C.c_int64),C.POINTER(C.c_int16),C.c_int32,C.c_int32,C.POINTER(C.c_int16),C.POINTER(C.c_int32)]
    f.restype=C.c_int32
    passed=0
    for n in (1,2,7):
        for active_mode in ('all','partial','none'):
            dest,source,dp,sp,plastic,count,offsets=fixture(n)
            expected=bytearray(dest.raw); src_before=[buf.raw for buf in source]
            yielded=(C.c_int16*(n*4)).from_address(sp[4])
            lo,hi=(1,2) if active_mode=='partial' else (0,n-1)
            active=None if active_mode=='none' else (C.c_int16*(hi-lo+1))(*([-1 if i%2 else 0 for i in range(hi-lo+1)]))
            want=[]
            for k in range(n):
                on=active is None or k<lo or k>hi or bool(active[k-lo])
                want.extend(-1 if on and yielded[4*k+i] else 0 for i in range(4))
                for i,size in enumerate(SIZES):
                    payload=C.string_at(sp[i]+k*(STRIDE if i==0 else size),size)
                    start=k*STRIDE+offsets[i]; expected[start:start+size]=payload
            delta=sum(v!=0 for v in want)-sum(v!=0 for v in plastic)
            wanted_count=count.value+delta
            assert f(n,STRIDE if n>1 else 0,dp,sp,active,lo,hi,plastic,C.byref(count))==0
            assert dest.raw==bytes(expected), 'Field bits or guard bytes changed'
            assert list(plastic)==want and count.value==wanted_count
            assert [buf.raw for buf in source]==src_before, 'Source modified'
            passed+=1
    for invalid in ('zero_n','negative_stride','short_stride','null_src','null_dst','null_field','count_overflow','count_underflow','stride_overflow','bad_active_range'):
        n=2; dest,source,dp,sp,plastic,count,offsets=fixture(n)
        stride=STRIDE; active=None; lo=0; hi=n-1; src_arg=sp; dst_arg=dp
        if invalid=='zero_n': n=0
        if invalid=='negative_stride': stride=-1
        if invalid=='short_stride': stride=100
        if invalid=='stride_overflow': n=3; stride=2**63-1
        if invalid=='null_src': src_arg=None
        if invalid=='null_dst': dst_arg=None
        if invalid=='null_field': sp[3]=0
        if invalid=='bad_active_range': active=(C.c_int16*1)(0); lo=2; hi=1
        if invalid in ('count_overflow','count_underflow'):
            yielded=(C.c_int16*8).from_address(sp[4])
            for i in range(8):
                yielded[i]=-1 if invalid=='count_overflow' else 0
                plastic[i]=0 if invalid=='count_overflow' else -1
            count.value=2**31-1 if invalid=='count_overflow' else -2**31
        before=(dest.raw,bytes(plastic),count.value,[buf.raw for buf in source])
        assert f(n,stride,dst_arg,src_arg,active,lo,hi,plastic,C.byref(count))==100, invalid
        assert before==(dest.raw,bytes(plastic),count.value,[buf.raw for buf in source]), ('Mutation on rejection',invalid)
        passed+=1
    print(json.dumps({'status':'PASS','cases':passed,'exact_field_bits_and_full_buffer_guards':True,'transactional_rejection':True}))

if __name__=='__main__': main()
