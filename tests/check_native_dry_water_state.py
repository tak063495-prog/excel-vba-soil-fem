"""Compare identical accepted/trial states, without solving or committing a stage."""
from pathlib import Path
import argparse, json, shutil
import win32com.client

ap=argparse.ArgumentParser()
ap.add_argument('--book',type=Path,required=True)
ap.add_argument('--checkpoint',type=Path,required=True)
ap.add_argument('--out',type=Path,required=True)
a=ap.parse_args();root=Path(__file__).resolve().parents[1]
a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=False)
dest=a.out/'hybrid_state_audit.xlsm';shutil.copy2(a.book,dest)
for dll in ('soilfem_native_x64.dll','soilfem_sparse_x64.dll'):
    if (a.book.resolve().parent/dll).exists():shutil.copy2(a.book.resolve().parent/dll,a.out/dll)
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;w=None
try:
    w=app.Workbooks.Open(str(dest),UpdateLinks=0)
    for ws in w.Worksheets:
        if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
    for source in (root/'src/vba').glob('*.bas'):
        cm=w.VBProject.VBComponents(source.stem).CodeModule
        cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(source.read_text(encoding='utf-8-sig'))
    cm=w.VBProject.VBComponents('FEMEngine').CodeModule
    cm.AddFromString('''
Public Function HybridStateAudit(ByVal folder As String, ByVal magnitude As Double) As Variant
 On Error GoTo Failed
 Dim base() As Double, forceA() As Double, forceB() As Double, st() As Double, sp() As Double, dp() As Double
 Dim e As Long,g As Long,i As Long,j As Long,p As Long,q As Long,k As Long
 Dim forceError As Double,stressError As Double,tangentError As Double,commitError As Double,t0 As Double,timeA As Double,timeB As Double
 SuppressUserMessages=True:P3LoadTimeFrame folder
 base=P3CommittedDisp:ReDim forceA(lastDof):ReDim forceB(lastDof)
 ReDim st(NumberOfElement*68-1):ReDim sp(NumberOfElement*192-1):ReDim dp(NumberOfElement*36-1)
 For i=0 To lastDof
  If NodeCond(i)=0 Then TDisp(i)=base(i)+magnitude*Sin(CDbl(i)+1#)
  UDisp(i)=TDisp(i)
 Next i
 FEMWriteNumericSetting "NATIVE_DRY_WATER",0#
 P3RestoreCommittedMaterialState False
 t0=Timer
 If Not P3EvaluateTrialState(base,forceA) Then Err.Raise 1001,,"VBA trial failed"
 timeA=Timer-t0:p=0:q=0:k=0
 For e=0 To NumberOfElement-1:For g=0 To 3
  For i=0 To 16:st(p)=Elem(e).Stmat(i,g):p=p+1:Next i
  For j=0 To 15:For i=0 To 2:sp(q)=Elem(e).Spmat(i,j,g):q=q+1:Next i:Next j
  For j=0 To 2:For i=0 To 2:dp(k)=Elem(e).Dpmat(i,j,g):k=k+1:Next i:Next j
 Next g:Next e
 FEMWriteNumericSetting "NATIVE_DRY_WATER",1#
 P3RestoreCommittedMaterialState False:FEMMatResetCounts
 t0=Timer
 If Not P3EvaluateTrialState(base,forceB) Then Err.Raise 1002,,"Hybrid trial failed"
 timeB=Timer-t0:p=0:q=0:k=0
 For e=0 To NumberOfElement-1:For g=0 To 3
  For i=0 To 16
   If i<>6 And i<>9 Then
    If Abs(st(p)-Elem(e).Stmat(i,g))>stressError Then stressError=Abs(st(p)-Elem(e).Stmat(i,g))
   End If
   p=p+1
  Next i
  For j=0 To 15:For i=0 To 2
   If Abs(sp(q)-Elem(e).Spmat(i,j,g))>tangentError Then tangentError=Abs(sp(q)-Elem(e).Spmat(i,j,g))
   q=q+1
  Next i:Next j
  For j=0 To 2:For i=0 To 2
   If Abs(dp(k)-Elem(e).Dpmat(i,j,g))>tangentError Then tangentError=Abs(dp(k)-Elem(e).Dpmat(i,j,g))
   k=k+1
  Next i:Next j
 Next g:Next e
 For i=0 To lastDof
  If Abs(forceA(i)-forceB(i))>forceError Then forceError=Abs(forceA(i)-forceB(i))
  If Abs(base(i)-P3CommittedDisp(i))>commitError Then commitError=Abs(base(i)-P3CommittedDisp(i))
 Next i
 HybridStateAudit=Array("PASS",forceError,stressError,tangentError,commitError,FEMMatDryWaterElements,timeA,timeB):Exit Function
Failed:
 HybridStateAudit=Array("ERROR|" & Err.Description):Err.Clear
End Function
''')
    cc=app.VBE.CommandBars.FindControl(1,578)
    if cc.Enabled:cc.Execute()
    assert not app.VBE.CommandBars.FindControl(1,578).Enabled
    reports=[]
    for magnitude in (0.,1e-7):
        value=list(app.Run("'"+w.Name+"'!HybridStateAudit",str(a.checkpoint.resolve()),magnitude))
        assert value[0]=='PASS' and len(value)==8,value
        report=dict(zip(('status','max_force_difference','max_state_difference','max_tangent_difference','committed_displacement_change','native_dry_elements','vba_seconds','hybrid_seconds'),value))
        report['magnitude']=magnitude;reports.append(report)
        (a.out/'results.json').write_text(json.dumps(reports,indent=2),encoding='utf-8')
        print(json.dumps(report),flush=True)
        assert value[1]<1e-8 and value[2]<1e-8 and value[3]<1e-6 and value[4]==0 and value[5]>0,value
finally:
    if w is not None:w.Close(False)
    app.Quit()
