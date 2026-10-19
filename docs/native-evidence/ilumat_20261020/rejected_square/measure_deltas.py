import sys,pathlib,ctypes as C,json,struct,math
sys.path.insert(0,str(pathlib.Path('tests/native').resolve()))
from test_material_points import cases
runtime=C.WinDLL('C:/Windows/System32/msvcr100.dll')
ds=[C.WinDLL(str(pathlib.Path(p).resolve())) for p in ('tests/tmp/mat_points_before_build/soilfem_mat_x64.dll','tests/tmp/mat_points_square_build/soilfem_mat_x64.dll')]
for d in ds:
 d.sfm_init.restype=C.c_int32;assert d.sfm_init()==4
 d.sfm_probe_size.restype=C.c_int32
 d.sfm_probe_point.argtypes=[C.POINTER(C.c_double)]*3+[C.c_ushort,C.POINTER(C.c_ubyte),C.POINTER(C.c_int32)];d.sfm_probe_point.restype=None
records=[]
for cw in (0x027f,0x037f):
 changed=decisions=flags=0;maximum=[0.]*23;relative=[0.]*23
 for prev,de,mat in cases(100000):
  args=[(C.c_double*len(v))(*v) for v in (prev,de,mat)];outputs=[]
  for d in ds:
   out=(C.c_ubyte*d.sfm_probe_size())();r=(C.c_int32*3)();d.sfm_probe_point(*args,cw,out,r);outputs.append((bytes(out),tuple(r)))
  a,b=outputs;decisions+=a[1][:2]!=b[1][:2];flags+=a[1][2]!=b[1][2]
  if a[1][0] and b[1][0]:
   changed+=a[0]!=b[0];x=struct.unpack('23d',a[0][:184]);y=struct.unpack('23d',b[0][:184])
   for i,(v,w) in enumerate(zip(x,y)):
    assert math.isfinite(v) and math.isfinite(w)
    maximum[i]=max(maximum[i],abs(v-w));relative[i]=max(relative[i],abs(v-w)/(1+abs(v)))
 records.append(dict(cw=hex(cw),points=100000,changed_output=changed,decision_differences=decisions,fp_flag_differences=flags,max_abs_by_field=maximum,max_scaled_by_field=relative))
p=pathlib.Path('tests/results/ilu_material/point_square_diagnostic.json');p.write_text(json.dumps(dict(cases=records),indent=2));print(p.read_text())
