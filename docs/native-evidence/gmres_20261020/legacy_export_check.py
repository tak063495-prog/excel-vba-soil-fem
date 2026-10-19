import sys,ctypes as C,json,pathlib
sys.path.insert(0,str(pathlib.Path('tests/native').resolve()))
from test_gmres_tuned import fixture
ds=[C.WinDLL(str(pathlib.Path(p).resolve())) for p in ('workbook/inco_assembly_20261020/soilfem_native_x64.dll','native/bin/soilfem_native_x64.dll')]
for d in ds:
 d.sfn_gmres.argtypes=[C.POINTER(C.c_int64),C.c_double];d.sfn_gmres.restype=C.c_int32
 d.sfn_set_precision.argtypes=[C.c_int32];d.sfn_set_precision.restype=C.c_int32
records=[]
for mode in (0,1,2):
 for dense in (False,True):
  for side in (1,2,3,5,8):
   for restart in (1,2,3,5):
    data=[]
    for d in ds:
     assert d.sfn_set_precision(mode)==0
     g,a=fixture(side,restart,dense=dense);before=bytes(g)
     rc=d.sfn_gmres(g,1e-11)
     assert before==bytes(g)
     data.append((rc,[bytes(v) for v in a]))
    assert data[0]==data[1],(mode,dense,side,restart)
    records.append(dict(mode=mode,dense=dense,side=side,restart=restart,rc=data[0][0]))
pathlib.Path('tests/results/gmres/legacy_gmres_check.json').write_text(json.dumps(dict(status='PASS',count=len(records),all_array_bytes_equal=True,cases=records),indent=2)+'\n')
print('PASS',len(records),'old-export cases, all solution/workspace/event bytes match previous DLL')
