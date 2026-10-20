"""Standalone ABI smoke test for the one-DLL build (does not start Excel)."""
import argparse, ctypes, hashlib, json
from pathlib import Path

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--dll', required=True); ap.add_argument('--old-mat'); ap.add_argument('--out', required=True); a=ap.parse_args()
    # sfm_init intentionally requires the same MSVCR100 runtime as VBA.
    ctypes.WinDLL(r'C:\Windows\System32\msvcr100.dll')
    new=ctypes.WinDLL(str(Path(a.dll).resolve()))
    new.sfm_init.restype=ctypes.c_int32; new.sfn_abi.restype=ctypes.c_int32
    record={'dll':str(Path(a.dll).resolve()),'dll_sha256':hashlib.sha256(Path(a.dll).read_bytes()).hexdigest(),
            'sfm_init':new.sfm_init(),'sfn_abi':new.sfn_abi()}
    if record['sfm_init'] != 4 or record['sfn_abi'] != 0x10001: raise SystemExit('combined runtime ABI failed')
    if a.old_mat:
        old=ctypes.WinDLL(str(Path(a.old_mat).resolve())); old.sfm_init.restype=ctypes.c_int32
        record['old_mat_sha256']=hashlib.sha256(Path(a.old_mat).read_bytes()).hexdigest()
        record['old_sfm_init']=old.sfm_init()
        if record['old_sfm_init'] != record['sfm_init']: raise SystemExit('old/new sfm_init mismatch')
    Path(a.out).write_text(json.dumps(record,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(record,indent=2))
if __name__=='__main__': main()
