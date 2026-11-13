"""Compare identical associated material states under both flow-policy labels."""
from pathlib import Path
import argparse,json,shutil,win32com.client
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True)
ap.add_argument('--out',type=Path,required=True);ap.add_argument('--native',choices=['AUTO','OFF'],default='AUTO')
ap.add_argument('--expect-label-dependence',action='store_true');a=ap.parse_args()
root=Path(__file__).resolve().parents[1];a.out.mkdir(parents=True,exist_ok=False)
dest=a.out.resolve()/'hourglass_policy.xlsm';shutil.copy2(a.book,dest)
shutil.copy2(a.book.resolve().parent/'soilfem_native_x64.dll',dest.parent/'soilfem_native_x64.dll')
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;w=None
try:
 w=app.Workbooks.Open(str(dest),UpdateLinks=0)
 for ws in w.Worksheets:
  if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
 for source in (root/'src/vba').glob('*.bas'):
  cm=w.VBProject.VBComponents(source.stem).CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(source.read_text(encoding='utf-8-sig'))
 cm=w.VBProject.VBComponents('FEMEngine').CodeModule;cm.AddFromString('''
Public Function HourglassPolicyAudit() As Variant
 On Error GoTo Failed
 Dim base() As Double,fd() As Double,fi() As Double,kd() As Double,sd() As Double
 Dim i As Long,j As Long,e As Long,g As Long,p As Long,n As Long,mode As Long,ks As Long
 Dim forceError As Double,matrixError As Double,stressError As Double,staticForceError As Double
 Dim committedError As Double,hgNorm As Double,projection As Double,hgForce() As Double
 SuppressUserMessages=True:P0_RunAnalysis
 If ResultStatus<>"PASS" Or GWActive Then Err.Raise 1001,,"Dry successful associated fixture required"
 For i=0 To NumberOfMaterial-1
  If Abs(Material(i).fai-Material(i).psai)>0.000000000001 Then Err.Raise 1002,,"Actual associated parameters required"
 Next i
 base=TDisp:ReDim fd(lastDof):ReDim fi(lastDof):ReDim hgForce(lastDof)
 ReDim kd(NumberOfElement*256-1):ReDim sd(NumberOfElement*68-1)
 For i=0 To lastDof
  If NodeCond(i)=0 Then TDisp(i)=base(i)+0.000001*Sin(CDbl(i)+1#)
  UDisp(i)=TDisp(i)
 Next i
 P3IncoElasticCorrection=False
 For ks=0 To 1
  P3SrmEnabled=(ks=1):P3PolicyCacheReady=True:P3FlowPolicyMode=1
  P3RestoreCommittedMaterialState False
  If Not P3EvaluateTrialState(base,fd) Then Err.Raise 1003,,"DAVIS evaluation"
  P3MarkAllTangentDirty
  If Not P3RebuildTangentFromSpmat() Then Err.Raise 1004,,"DAVIS tangent"
  p=0:n=0
  For e=0 To NumberOfElement-1
   For i=0 To 15:For j=0 To 15:kd(p)=Elem(e).kmat(i,j):p=p+1:Next j:Next i
   For g=0 To 3:For i=0 To 16:sd(n)=Elem(e).Stmat(i,g):n=n+1:Next i:Next g
  Next e
  P3FlowPolicyMode=0
  P3RestoreCommittedMaterialState False
  If Not P3EvaluateTrialState(base,fi) Then Err.Raise 1005,,"INCONSISTENT evaluation"
  P3MarkAllTangentDirty
  If Not P3RebuildTangentFromSpmat() Then Err.Raise 1006,,"INCONSISTENT tangent"
  p=0:n=0
  For e=0 To NumberOfElement-1
   For i=0 To 15:For j=0 To 15
    If ks=1 And Abs(kd(p)-Elem(e).kmat(i,j))>matrixError Then matrixError=Abs(kd(p)-Elem(e).kmat(i,j))
    p=p+1
   Next j:Next i
   For g=0 To 3:For i=0 To 16
    If Abs(sd(n)-Elem(e).Stmat(i,g))>stressError Then stressError=Abs(sd(n)-Elem(e).Stmat(i,g))
    n=n+1
   Next i:Next g
  Next e
  For i=0 To lastDof
   If ks=1 Then
    If Abs(fd(i)-fi(i))>forceError Then forceError=Abs(fd(i)-fi(i))
   Else
    If Abs(fd(i)-fi(i))>staticForceError Then staticForceError=Abs(fd(i)-fi(i))
   End If
  Next i
 Next ks
 For e=0 To NumberOfElement-1
  For mode=0 To Elem(e).HourglassModeCount-1
   projection=0#
   For j=0 To 15
    i=Elem(e).elnode(j)
    If i>=0 Then projection=projection+Elem(e).HourglassMode(mode,j)*TDisp(i)
   Next j
   For j=0 To 15
    i=Elem(e).elnode(j)
    If i>=0 Then hgForce(i)=hgForce(i)+Elem(e).HourglassScale*projection*Elem(e).HourglassMode(mode,j)
   Next j
  Next mode
 Next e
 For i=0 To lastDof
  If Abs(P3CommittedDisp(i)-base(i))>committedError Then committedError=Abs(P3CommittedDisp(i)-base(i))
  If Abs(hgForce(i))>hgNorm Then hgNorm=Abs(hgForce(i))
 Next i
 HourglassPolicyAudit=Array("PASS",forceError,matrixError,stressError,staticForceError,committedError,hgNorm,FEMMatStatusText()):Exit Function
Failed:
 HourglassPolicyAudit=Array("ERROR|" & Err.Description)
End Function
''')
 prefix="'"+w.Name+"'!"
 app.Run(prefix+'FEMWriteTextSetting','NATIVE_DLL',a.native)
 app.Run(prefix+'FEMWriteNumericSetting','NATIVE_SPARSE_DIRECT',0.)
 app.Run(prefix+'FEMWriteTextSetting','EXPORT_LOAD','')
 control=app.VBE.CommandBars.FindControl(1,578)
 if control.Enabled:control.Execute()
 assert not app.VBE.CommandBars.FindControl(1,578).Enabled
 v=list(app.Run(prefix+'HourglassPolicyAudit'));assert len(v)==8 and v[0]=='PASS',v
 assert v[3]<1e-10 and v[4]<1e-10 and v[5]==0 and v[6]>1e-8,v
 if a.expect_label_dependence:assert v[1]>1e-8 and v[2]>1e-4,v
 else:assert v[1]<1e-10 and v[2]<1e-8,v
 record=dict(zip(('execution','srm_force_label_difference','srm_tangent_label_difference','stress_label_difference','ordinary_static_force_difference','committed_change','nonzero_hourglass_control_force','material_backend'),v))
 record.update(status='LABEL_DEPENDENCE_REPRODUCED' if a.expect_label_dependence else 'PASS',native=a.native)
 (a.out/'results.json').write_bytes(json.dumps(record,indent=2).encode('utf-8'));print(json.dumps(record),flush=True)
finally:
 if w is not None:w.Close(False)
 app.Quit()
