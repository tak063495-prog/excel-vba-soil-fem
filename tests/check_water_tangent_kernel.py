"""Audit optional native tangent integration with physical water and VBA material updates."""
from pathlib import Path
import argparse,json,shutil,win32com.client
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args()
root=Path(__file__).resolve().parents[1];a.out=a.out.resolve();assert not a.out.exists();a.out.mkdir(parents=True)
dst=a.out/'water_tangent_audit.xlsm';shutil.copy2(a.book,dst)
shutil.copy2(a.book.resolve().parent/'soilfem_native_x64.dll',a.out/'soilfem_native_x64.dll')
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;w=None
try:
 w=app.Workbooks.Open(str(dst),UpdateLinks=0)
 for ws in w.Worksheets:
  if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
 for source in (root/'src/vba').glob('*.bas'):
  try:c=w.VBProject.VBComponents(source.stem)
  except Exception:c=w.VBProject.VBComponents.Add(1);c.Name=source.stem
  cm=c.CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(source.read_text(encoding='utf-8-sig'))
 cm=w.VBProject.VBComponents('FEMEngine').CodeModule;cm.AddFromString('''
Public Function WaterTangentAudit() As Variant
 Dim saved() As Double,stress() As Double,weight() As Double,sw() As Double
 Dim e As Long,i As Long,j As Long,g As Long,k As Long,calls As Long,evalCalls As Long
 Dim maxK As Double,maxStress As Double,maxWeight As Double,maxSW As Double
 SuppressUserMessages=True:P0_RunAnalysis
 If ResultStatus<>"PASS" Or Not GWActive Then Err.Raise 1001,,"Physical-water PASS fixture required"
 evalCalls=FEMMatEvalCalls
 ReDim saved(0 To NumberOfElement*256-1):ReDim stress(0 To NumberOfElement*68-1):ReDim weight(0 To NumberOfElement-1)
 sw=P3SelfWeightForce
 P3MarkAllTangentDirty
 If Not P3RebuildTangentFromSpmat() Then Err.Raise 1002,,"Native tangent failed"
 calls=FEMMatTangentCalls
 For e=0 To NumberOfElement-1
  weight(e)=Elem(e).ElmWeight
  For i=0 To 15:For j=0 To 15:saved(e*256+i*16+j)=Elem(e).kmat(i,j):Next j:Next i
  For g=0 To 3:For i=0 To 16:stress(e*68+g*17+i)=Elem(e).Stmat(i,g):Next i:Next g
 Next e
 FEMNativeSetMode "OFF"
 P3MarkAllTangentDirty
 If Not P3RebuildTangentFromSpmat() Then Err.Raise 1003,,"VBA tangent failed"
 For e=0 To NumberOfElement-1
  If Abs(weight(e)-Elem(e).ElmWeight)>maxWeight Then maxWeight=Abs(weight(e)-Elem(e).ElmWeight)
  For i=0 To 15:For j=0 To 15
   If Abs(saved(e*256+i*16+j)-Elem(e).kmat(i,j))>maxK Then maxK=Abs(saved(e*256+i*16+j)-Elem(e).kmat(i,j))
  Next j:Next i
  For g=0 To 3:For i=0 To 16
   If Abs(stress(e*68+g*17+i)-Elem(e).Stmat(i,g))>maxStress Then maxStress=Abs(stress(e*68+g*17+i)-Elem(e).Stmat(i,g))
  Next i:Next g
 Next e
 For i=0 To lastDof:If Abs(sw(i)-P3SelfWeightForce(i))>maxSW Then maxSW=Abs(sw(i)-P3SelfWeightForce(i))
 Next i
 WaterTangentAudit=Array(ResultStatus,calls,evalCalls,maxK,maxStress,maxWeight,maxSW)
End Function
''')
 prefix="'"+w.Name+"'!"
 for key,value in [('NATIVE_DLL','AUTO'),('NATIVE_WATER_TANGENT',1.),('NATIVE_SPARSE_DIRECT',0.),('EXPORT_LOAD','')]:app.Run(prefix+('FEMWriteTextSetting' if isinstance(value,str) else 'FEMWriteNumericSetting'),key,value)
 cc=app.VBE.CommandBars.FindControl(1,578)
 if cc.Enabled:cc.Execute()
 assert not app.VBE.CommandBars.FindControl(1,578).Enabled,'Probe compile failed'
 v=list(app.Run(prefix+'WaterTangentAudit'));assert v[0]=='PASS' and v[1]>0 and v[2]==0 and max(v[3:])<1e-9,v
 record=dict(zip(('analysis','native_tangent_calls','native_material_evaluations','max_element_K_error','max_stress_change','max_element_weight_error','max_self_weight_force_change'),v));record['status']='PASS'
 (a.out/'results.json').write_text(json.dumps(record,indent=2,allow_nan=False),encoding='utf-8');print(json.dumps(record),flush=True)
finally:
 if w is not None:w.Close(False)
 app.Quit()
