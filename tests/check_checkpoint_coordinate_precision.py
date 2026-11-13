"""Test shallow-water Q8/T6 checkpoints, legacy coordinate rounding and corruption."""
from pathlib import Path
import argparse,csv,json,shutil,subprocess,sys
import win32com.client

ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True)
ap.add_argument('--out',type=Path,required=True);ap.add_argument('--runner',type=Path)
a=ap.parse_args();root=Path(__file__).resolve().parents[1];a.out=a.out.resolve()
a.out.mkdir(parents=True,exist_ok=False)
data=json.loads((root/'tests/fixtures/mixed_t6/mixed_water.json').read_text(encoding='utf-8'))
offset=.123456789045
for node in data['nodes']:node[1]+=offset;node[2]+=offset
level=offset+(1-.5773502692)/2+.0005
data['name']='shallow_water_precision'
data['water']['profiles']=[['W',-1+offset,level,''],['W',3+offset,level,'']]
data.setdefault('settings',{})['EXPORT_MODE']='STAGE'
fixture=a.out/'fixture.json';fixture.write_text(json.dumps(data),encoding='utf-8')
runner=(a.runner or root/'tests/run_groundwater_fixture.ps1').resolve()
dest=a.out/'analysis';log=a.out/'analysis.log'
with log.open('w',encoding='utf-8') as stream:
 result=subprocess.run(['pwsh','-NoProfile','-File',str(runner),'-Book',str(a.book.resolve()),'-Fixture',str(fixture),'-Out',str(dest),'-Native','AUTO','-Resume'],cwd=root,stdout=stream,stderr=subprocess.STDOUT)
if result.returncode:print(log.read_text(encoding='utf-8')[-2500:]);sys.exit(1)
scientific=dest/f"{data['name']}_out"/'stg01_GRAVITY'
node=list(csv.reader((scientific/'node.csv').open()))
assert all('E' in value.upper() for row in node[1:] for value in row[1:5])
expected={int(n[0]):n[1:3] for n in data['nodes']}
error=max(abs(float(row[j])-expected[int(row[0])][j-1]) for row in node[1:] for j in (1,2))
assert error<1e-13,error
legacy=a.out/'legacy_fixed10';shutil.copytree(scientific,legacy)
for row in node[1:]:
 for j in (1,2):row[j]=format(float(row[j]),'.10f')
 for j in (3,4):row[j]=format(float(row[j]),'.12f')
with (legacy/'node.csv').open('w',newline='',encoding='utf-8') as stream:csv.writer(stream).writerows(node)
legacy_error=max(abs(float(row[j])-expected[int(row[0])][j-1]) for row in node[1:] for j in (1,2))
assert legacy_error>4e-11,legacy_error
scientific_rounded=a.out/'scientific_rounded_coordinates';shutil.copytree(legacy,scientific_rounded)
for row in node[1:]:
 for j in (1,2):row[j]=format(float(row[j]),'.16e')
with (scientific_rounded/'node.csv').open('w',newline='',encoding='utf-8') as stream:csv.writer(stream).writerows(node)
corrupt=a.out/'legacy_corrupt_pressure';shutil.copytree(legacy,corrupt)
water=list(csv.reader((corrupt/'water.csv').open()))
row=next(r for r in water[4:] if float(r[2])>0)
row[2]=format(float(row[2])+1e-7,'.17e')
with (corrupt/'water.csv').open('w',newline='',encoding='utf-8') as stream:csv.writer(stream).writerows(water)
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;w=None
try:
 w=app.Workbooks.Open(str(dest/'fresh_results.xlsm'),UpdateLinks=0)
 for ws in w.Worksheets:
  if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
 mod=w.VBProject.VBComponents.Add(1);mod.Name='CoordinatePrecisionProbe';mod.CodeModule.AddFromString('''
Public Function CoordinatePrecisionView(ByVal folder As String) As String
 On Error GoTo Failed
 SuppressUserMessages=True
 P3LoadTimeFrame folder
 CoordinatePrecisionView="PASS":Exit Function
Failed:
 CoordinatePrecisionView="ERROR|" & Err.Description
End Function
''')
 cm=w.VBProject.VBComponents('FEMWater').CodeModule;cm.AddFromString('''
Public Function CoordinatePrecisionEndpointProbe() As Variant
 Dim saved As Variant,points As Variant,profile As String,px As Double,py As Double,oldOutside As Boolean
 Dim interior As Double,endpoint As Double
 profile=mMatProfile(Elem(0).MatNo):saved=mProfileData(profile):points=saved
 oldOutside=mOutsideDry
 interior=GWLegacyPressureRoundoff(0,0)
 FEMGaussPhysicalXY 0,0,px,py
 points(0,0)=px:mProfileData(profile)=points:mOutsideDry=True
 endpoint=GWLegacyPressureRoundoff(0,0)
 mProfileData(profile)=saved:mOutsideDry=oldOutside
 CoordinatePrecisionEndpointProbe=Array(interior,endpoint)
End Function
''')
 prefix="'"+w.Name+"'!";views={}
 for name,folder in (('scientific',scientific),('legacy_fixed10',legacy),('legacy_corrupt_pressure',corrupt),('scientific_rounded_coordinates',scientific_rounded)):
  value=app.Run(prefix+'CoordinatePrecisionView',str(folder));views[name]=value
  if name in ('legacy_corrupt_pressure','scientific_rounded_coordinates'):assert value.startswith('ERROR|') and 'チェックポイント' in value,value
  else:assert value=='PASS',(name,value)
  if name=='scientific':
   interior,endpoint=app.Run(prefix+'CoordinatePrecisionEndpointProbe')
   assert interior>0 and endpoint==0,(interior,endpoint)
 record={'status':'PASS','fresh_and_resume':'PASS','new_coordinate_error_m':error,'legacy_coordinate_error_m':legacy_error,'views':views,'corrupt_pressure_delta_kPa':1e-7,'interior_roundoff_bound_kPa':interior,'dry_endpoint_roundoff_bound_kPa':endpoint}
 (a.out/'results.json').write_text(json.dumps(record,indent=2,ensure_ascii=False),encoding='utf-8')
 print(json.dumps(record),flush=True)
finally:
 if w is not None:w.Close(False)
 app.Quit()
