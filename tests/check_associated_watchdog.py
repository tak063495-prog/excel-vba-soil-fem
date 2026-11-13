"""Exercise associated look-ahead acceptance and rollback on a real elastic FE model."""
from pathlib import Path
import argparse,json,shutil,win32com.client
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args()
root=Path(__file__).resolve().parents[1];a.out=a.out.resolve();assert not a.out.exists();a.out.mkdir(parents=True)
dst=a.out/'associated_watchdog_audit.xlsm';shutil.copy2(a.book,dst);shutil.copy2(a.book.resolve().parent/'soilfem_native_x64.dll',a.out/'soilfem_native_x64.dll')
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
Public Function AssociatedWatchdogAudit() As Variant
 Dim base() As Double,internal() As Double,target() As Double,r() As Double,kd() As Double,direction() As Double
 Dim i As Long,j As Long,lo As Long,hi As Long,it As Long,accepted As Boolean,rejected As Boolean,inc As Long
 Dim aa As Double,bb As Double,rr As Double,beta As Double,original As Double,candidateRatio As Double,endpointWork As Double,afterRatio As Double,correction As Double,commitError As Double,rollbackError As Double
 SuppressUserMessages=True:P0_RunAnalysis
 If ResultStatus<>"PASS" Or P6UseCSR Then Err.Raise 1001,,"Symmetric face_support/static fixture required"
 If Not P3AssociatedPotentialEligible() Then Err.Raise 1002,,"Associated scope required"
 base=TDisp:ReDim internal(lastDof):ReDim target(lastDof):ReDim r(lastDof):ReDim kd(lastDof):ReDim direction(lastDof)
 If Not P3EvaluateTrialState(base,internal) Then Err.Raise 1003
 P3MarkAllTangentDirty
 If Not P3RebuildTangentFromSpmat() Then Err.Raise 1004
 lo=-1:hi=-1
 For i=0 To lastDof
  If NodeCond(i)=0 Then
   If lo<0 Then lo=i:hi=i
   If OriginalMat(i,0)<OriginalMat(lo,0) Then lo=i
   If OriginalMat(i,0)>OriginalMat(hi,0) Then hi=i
  End If
 Next i
 r(lo)=0.001:r(hi)=0.001*Sqr(OriginalMat(lo,0)/OriginalMat(hi,0))
 For i=0 To lastDof
  target(i)=internal(i)+r(i)
  If NodeCond(i)=0 Then
   For j=0 To lastDof
    If NodeCond(j)=0 And Abs(i-j)<=BandWidth Then kd(i)=kd(i)+OriginalMat(Application.Min(i,j),Abs(i-j))*r(j)
   Next j
   aa=aa+r(i)*kd(i):bb=bb+kd(i)^2:rr=rr+r(i)^2
  End If
 Next i
 If rr*bb/(aa*aa)<=4# Then Err.Raise 1005,,"Fixture lacks the intended norm/potential ridge"
 beta=0.5*rr/aa:original=Sqr(rr)
 For i=0 To lastDof:direction(i)=beta*r(i):TDisp(i)=base(i)+direction(i):UDisp(i)=TDisp(i):Next i
 If Not P3EvaluateTrialState(base,internal) Then Err.Raise 1006
 P3UpdateResidualMetrics target,internal
 candidateRatio=ResidualNormFree/original:endpointWork=P3ResidualDirectionWork(target,internal,direction)
 P3RestoreCommittedMaterialState False
 TDisp=base:UDisp=base:Disp=direction
 inc=P3SuccessfulIncrementCount
 accepted=P3AccelApplyCorrection(base,target,internal,correction,original,it)
 afterRatio=ResidualNormFree/original
 For i=0 To lastDof:If Abs(P3CommittedDisp(i)-base(i))>commitError Then commitError=Abs(P3CommittedDisp(i)-base(i))
 Next i
 P3RestoreCommittedMaterialState False
 TDisp=base:UDisp=base
 For i=0 To lastDof:Disp(i)=-direction(i):Next i
 it=0:rejected=Not P3TryNewtonWatchdog(base,target,internal,correction,it)
 For i=0 To lastDof:If Abs(TDisp(i)-base(i))>rollbackError Then rollbackError=Abs(TDisp(i)-base(i))
 Next i
 AssociatedWatchdogAudit=Array(accepted,candidateRatio,endpointWork,afterRatio,commitError,rejected,rollbackError,P3SuccessfulIncrementCount-inc)
End Function
''')
 prefix="'"+w.Name+"'!"
 for key,value in [('FLOW_POLICY','DAVIS'),('NATIVE_SPARSE_DIRECT',0.),('ASSOCIATED_ENERGY_RECOVERY',1.),('ASSOCIATED_ENERGY_FIRST',1.),('EXPORT_LOAD','')]:app.Run(prefix+('FEMWriteTextSetting' if isinstance(value,str) else 'FEMWriteNumericSetting'),key,value)
 cc=app.VBE.CommandBars.FindControl(1,578)
 if cc.Enabled:cc.Execute()
 assert not app.VBE.CommandBars.FindControl(1,578).Enabled,'Probe compile failed'
 v=list(app.Run(prefix+'AssociatedWatchdogAudit'));assert v[0] and v[1]>1 and v[2]>0 and v[3]<1e-4 and v[4]==0 and v[5] and v[6]==0 and v[7]==0,v
 r=dict(zip(('accepted','intermediate_residual_ratio','positive_endpoint_work','accepted_residual_ratio','committed_displacement_change','negative_work_rejected','rollback_displacement_error','committed_increment_change'),v));r['status']='PASS'
 (a.out/'results.json').write_text(json.dumps(r,indent=2,allow_nan=False),encoding='utf-8');print(json.dumps(r),flush=True)
finally:
 if w is not None:w.Close(False)
 app.Quit()
