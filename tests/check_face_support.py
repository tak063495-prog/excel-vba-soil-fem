"""Audit compression-only support, consistent global tangent and force balance."""
from pathlib import Path
import argparse,json,shutil,win32com.client
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args();a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=True)
root=Path(__file__).resolve().parents[1];copy=a.out/'face_audit.xlsm';assert not copy.exists();shutil.copy2(a.book,copy);shutil.copy2(a.book.resolve().parent/'soilfem_native_x64.dll',a.out/'soilfem_native_x64.dll')
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;app.AutomationSecurity=1;wb=None
try:
 wb=app.Workbooks.Open(str(copy),UpdateLinks=0);app.Calculation=-4135
 for ws in wb.Worksheets:
  if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
 for source in (root/'src/vba').glob('*.bas'):
  try:component=wb.VBProject.VBComponents(source.stem)
  except Exception:component=wb.VBProject.VBComponents.Add(1);component.Name=source.stem
  cm=component.CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(source.read_text(encoding='utf-8-sig'))
 try:wb.VBProject.VBComponents.Remove(wb.VBProject.VBComponents('FaceAuditDriver'))
 except Exception:pass
 cm=wb.VBProject.VBComponents('FEMEngine').CodeModule;cm.AddFromString('''
Public Function FaceMatrixProbe() As Variant
 Dim base() As Double,f0() As Double,f1() As Double,d() As Double,kd() As Double
 Dim i As Long,j As Long,p As Long,h As Double,err2 As Double,ref2 As Double
 base=TDisp:ReDim f0(lastDof):ReDim f1(lastDof):ReDim d(lastDof):ReDim kd(lastDof)
 If Not P3EvaluateTrialState(base,f0) Then Err.Raise 1001,,"base force failed"
 If Not P3RebuildTangentFromSpmat() Then Err.Raise 1002,,"tangent failed"
 For i=0 To lastDof:d(i)=Sin(CDbl(i)+1#):Next i
 If P6UseCSR Then
  For i=0 To lastDof
   For p=P6CSRRowPtr(i) To P6CSRRowPtr(i+1)-1:kd(i)=kd(i)+P6CSROriginalValues(p)*d(P6CSRColumnIndex(p)):Next p
  Next i
 Else
  For i=0 To lastDof
   For j=0 To lastDof
    If Abs(j-i)<=BandWidth Then kd(i)=kd(i)+OriginalMat(Application.Min(i,j),Abs(j-i))*d(j)
   Next j
  Next i
 End If
 h=0.00000001
 For i=0 To lastDof:TDisp(i)=base(i)+h*d(i):Next i
 If Not P3EvaluateTrialState(base,f1) Then Err.Raise 1003,,"perturbed force failed"
 For i=0 To lastDof
  err2=err2+((f1(i)-f0(i))/h-kd(i))^2:ref2=ref2+kd(i)^2
 Next i
 TDisp=base
 FaceMatrixProbe=Array(Sqr(err2/ref2),P6UseCSR,BandWidth)
End Function
''')
 cm=wb.VBProject.VBComponents('FEMSolver').CodeModule;cm.AddFromString('''
Public Function FaceFallbackMatrixProbe() As Double
 Dim before() As Double,i As Long,j As Long,matrixError As Double
 If P6UseCSR Then Exit Function
 before=OriginalMat
 If Not P6RebuildBandFromElements() Then Err.Raise 1004,,"Band fallback rebuild failed"
 For i=0 To lastDof
  For j=0 To BandWidth
   If Abs(before(i,j)-OriginalMat(i,j))>matrixError Then matrixError=Abs(before(i,j)-OriginalMat(i,j))
  Next j
 Next i
 FaceFallbackMatrixProbe=matrixError
End Function
''')
 cm=wb.VBProject.VBComponents('FEMFaceSupport').CodeModule;cm.AddFromString('''
Public Function FaceInvariantProbe() As Variant
 Dim saved() As Double,f() As Double,i As Long,j As Long,rigidMax As Double,openMax As Double,netX As Double,moment As Double,slaveForce As Double
 saved=TDisp:ReDim f(lastDof)
 For i=0 To lastDof Step 2:TDisp(i)=0.1+0.002*XXX(i+1):TDisp(i+1)=-0.05-0.002*XXX(i):Next i
 FaceSupportAddForce f
 For i=0 To lastDof:If Abs(f(i))>rigidMax Then rigidMax=Abs(f(i))
 Next i
 For i=0 To lastDof:TDisp(i)=0#:f(i)=0#:Next i
 For i=0 To mCount-1:TDisp(mDof(i,0))=mSense(i)*0.001:Next i
 FaceSupportAddForce f
 For i=0 To lastDof:If Abs(f(i))>openMax Then openMax=Abs(f(i))
 Next i
 For i=0 To lastDof:TDisp(i)=0#:f(i)=0#:Next i
 TDisp(mDof(0,0))=-mSense(0)*0.001
 FaceSupportAddForce f:slaveForce=f(mDof(0,0))*mSense(0)
 For i=0 To lastDof Step 2:netX=netX+f(i):moment=moment-XXX(i+1)*f(i):Next i
 TDisp=saved
 FaceInvariantProbe=Array(rigidMax,openMax,netX,moment,slaveForce)
End Function
''')
 cc=app.VBE.CommandBars.FindControl(1,578)
 if cc.Enabled:cc.Execute()
 if app.VBE.CommandBars.FindControl(1,578).Enabled:
  pane=app.VBE.ActiveCodePane;selection=pane.GetSelection();print('COMPILE_LOCATION',pane.CodeModule.Parent.Name,selection,pane.CodeModule.Lines(selection[0],1),flush=True)
 assert not app.VBE.CommandBars.FindControl(1,578).Enabled,'Probe compilation failed'
 prefix="'"+wb.Name+"'!";app.Run(prefix+'FEMWriteTextSetting','EXPORT_LOAD','');app.Run(prefix+'FEMWriteTextSetting','EXPORT_MODE','OFF')
 mod=wb.VBProject.VBComponents.Add(1);mod.Name='FaceAuditDriver';mod.CodeModule.AddFromString('''
Public Function FaceAuditRun() As Variant
 SuppressUserMessages=True:P0_RunAnalysis
 FaceAuditRun=Array(ResultStatus,AnalysisMessage,TotalSelfWeight,P6ReactionSumX,P6ReactionSumY,RelativeResidualFree)
End Function
''')
 r=list(app.Run(prefix+'FaceAuditRun'));assert r[0]=='PASS' and abs(r[2]-36)<1e-8 and abs(r[3]-1)<1e-8 and abs(r[4]-36)<1e-8 and r[5]<=1e-5,r
 matrix=list(app.Run(prefix+'FaceMatrixProbe'));fallback=float(app.Run(prefix+'FaceFallbackMatrixProbe'));invariants=list(app.Run(prefix+'FaceInvariantProbe'))
 assert fallback<1e-8,fallback
 assert matrix[0]<1e-6 and max(abs(x) for x in invariants[:4])<1e-8 and invariants[4]<0,(matrix,invariants)
 record={'status':'PASS','analysis':r,'directional_jacobian_error':matrix[0],'csr':matrix[1],'bandwidth':matrix[2],'rigid_motion_force':invariants[0],'opening_force':invariants[1],'net_support_force':invariants[2],'net_support_moment':invariants[3],'compression_internal_force':invariants[4],'band_fallback_matrix_max_error':fallback}
 (a.out/'results.json').write_text(json.dumps(record,indent=2),encoding='utf-8');print(json.dumps(record),flush=True)
finally:
 if wb is not None:wb.Close(False)
 app.Quit()
