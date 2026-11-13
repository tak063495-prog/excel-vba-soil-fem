"""Reject a singular uncommitted look-ahead operator without poisoning its base."""
from pathlib import Path
import argparse, ast, json, re, shutil
import win32com.client

ap=argparse.ArgumentParser()
ap.add_argument('--book',type=Path,required=True)
ap.add_argument('--out',type=Path,required=True)
ap.add_argument('--failure',choices=['singular','restore_singular','runtime','capacity','input','cancel'],default='singular')
ap.add_argument('--native',choices=['AUTO','OFF'],default='AUTO')
ap.add_argument('--expect-old-fatal',action='store_true')
a=ap.parse_args();root=Path(__file__).resolve().parents[1]
a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=False)
dst=a.out/'linear_trial_reject.xlsm';shutil.copy2(a.book,dst)
shutil.copy2(a.book.resolve().parent/'soilfem_native_x64.dll',a.out/'soilfem_native_x64.dll')
# Reuse the existing real FE ridge fixture builder, without executing its CLI.
tree=ast.parse((root/'tests/check_associated_watchdog.py').read_text(encoding='utf-8-sig'))
drivers=[n.value for n in ast.walk(tree) if isinstance(n,ast.Constant) and isinstance(n.value,str) and 'Public Function AssociatedWatchdogAudit' in n.value]
assert len(drivers)==1
driver=drivers[0].replace('Private AssociatedBudgetTestDamping As Boolean','Private AssociatedBudgetTestDamping As Boolean\nPrivate LFBase() As Double,LFTarget() As Double,LFDirection() As Double\nPrivate LFMode As String,LFUsed As Long')
driver=driver.replace(' inc=P3SuccessfulIncrementCount',' LFBase=base:LFTarget=target:LFDirection=direction\n inc=P3SuccessfulIncrementCount')
driver+='''
Private Function LinearRejectBandForTest() As Boolean
 If LFMode<>"" And LFUsed=0 Then
  LFUsed=1
  Select Case LFMode
   Case "singular","restore_singular":SetAnalysisFailure RESULT_GLOBAL_SINGULAR,"Injected uncommitted trial pivot",vbObjectError+3304,-1,-1,CurrentIncrement,CurrentIteration
   Case "runtime":SetAnalysisFailure RESULT_RUNTIME_ERROR,"Injected runtime error",5,-1,-1,CurrentIncrement,CurrentIteration
   Case "capacity":SetAnalysisFailure RESULT_CAPACITY_ERROR,"Injected capacity error",7,-1,-1,CurrentIncrement,CurrentIteration
   Case "input":SetAnalysisFailure RESULT_INPUT_ERROR,"Injected input error",vbObjectError+3001,-1,-1,CurrentIncrement,CurrentIteration
   Case "cancel":P3UserCancel=True:SetAnalysisFailure RESULT_NONCONVERGED,"Injected cancel",18,-1,-1,CurrentIncrement,CurrentIteration
  End Select
  Exit Function
 End If
 LinearRejectBandForTest=BandSolver()
End Function
Private Function LinearRejectRestoreForTest() As Boolean
 If LFMode="restore_singular" And LFUsed>0 Then
  SetAnalysisFailure RESULT_GLOBAL_SINGULAR,"Injected base restoration failure",vbObjectError+3304,-1,-1,CurrentIncrement,CurrentIteration
  Exit Function
 End If
 LinearRejectRestoreForTest=P3RebuildTangentFromSpmat()
End Function
Public Function AssociatedLinearRejectAudit(ByVal mode As String) As Variant
 On Error GoTo Failed
 Dim setup As Variant,baseForce() As Double,internal() As Double,baseK() As Double
 Dim i As Long,j As Long,p As Long,e As Long,it As Long,inc As Long
 Dim norm As Double,correction As Double,rollback As Double,forceError As Double,kError As Double,commitError As Double
 Dim rejected As Boolean,status As String,errorNumber As Long,fatal As Boolean,accepted As Boolean,ratio As Double,note As String
 LFMode="":setup=AssociatedWatchdogAudit(False)
 If VarType(setup(0))=vbString Then Err.Raise 1001,,CStr(setup(0))
 P3ClearTransientFailure:P3UserCancel=False:P3SearchFatal=False
 P3SrmTrialRunning=True:P3SrmTrialFailNote=""
 P3RestoreCommittedMaterialState False:TDisp=LFBase:UDisp=LFBase:Disp=LFDirection
 ReDim internal(lastDof):ReDim baseK(NumberOfElement*256-1)
 If Not P3EvaluateTrialState(LFBase,internal) Then Err.Raise 1002
 P3UpdateResidualMetrics LFTarget,internal
 P3MarkAllTangentDirty:If Not P3RebuildTangentFromSpmat() Then Err.Raise 1003
 baseForce=internal:norm=ResidualNormFree:inc=P3SuccessfulIncrementCount:p=0
 For e=0 To NumberOfElement-1:For j=0 To 15:For i=0 To 15
  baseK(p)=Elem(e).kmat(i,j):p=p+1
 Next i:Next j:Next e
 LFMode=mode:LFUsed=0:it=0
 rejected=Not P3TryAssociatedPotentialWatchdog(LFBase,LFTarget,internal,correction,it)
 status=ResultStatus:errorNumber=AnalysisErrorNumber:fatal=P3IsFatalFailure():note=P3SrmTrialFailNote:p=0
 For i=0 To lastDof
  If Abs(TDisp(i)-LFBase(i))>rollback Then rollback=Abs(TDisp(i)-LFBase(i))
  If Abs(internal(i)-baseForce(i))>forceError Then forceError=Abs(internal(i)-baseForce(i))
  If Abs(P3CommittedDisp(i)-LFBase(i))>commitError Then commitError=Abs(P3CommittedDisp(i)-LFBase(i))
 Next i
 For e=0 To NumberOfElement-1:For j=0 To 15:For i=0 To 15
  If Abs(Elem(e).kmat(i,j)-baseK(p))>kError Then kError=Abs(Elem(e).kmat(i,j)-baseK(p))
  p=p+1
 Next i:Next j:Next e
 LFMode="":P3UserCancel=False:P3SearchFatal=False:P3ClearTransientFailure
 P3SrmTrialFailNote=""
 P3RestoreCommittedMaterialState False:TDisp=LFBase:UDisp=LFBase:Disp=LFDirection
 If Not P3EvaluateTrialState(LFBase,internal) Then Err.Raise 1004
 P3UpdateResidualMetrics LFTarget,internal
 P3MarkAllTangentDirty:If Not P3RebuildTangentFromSpmat() Then Err.Raise 1005
 LFMode=mode:LFUsed=0:it=0
 accepted=P3AccelApplyCorrection(LFBase,LFTarget,internal,correction,norm,it)
 ratio=ResidualNormFree/norm
 AssociatedLinearRejectAudit=Array(rejected,status,errorNumber,fatal,rollback,forceError,kError,commitError,accepted,ratio,P3IsFatalFailure(),ResultStatus,P3SuccessfulIncrementCount-inc,LFUsed,note,P3SrmTrialFailNote):Exit Function
Failed:
 AssociatedLinearRejectAudit=Array("ERROR|" & Err.Description):Err.Clear
End Function
'''
driver=driver.replace('P3MarkAllTangentDirty:', 'Call P3MarkAllTangentDirty:').replace('P3ClearTransientFailure:', 'Call P3ClearTransientFailure:')
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;w=None
try:
 w=app.Workbooks.Open(str(dst),UpdateLinks=0)
 for ws in w.Worksheets:
  if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
 for source in (root/'src/vba').glob('*.bas'):
  cm=w.VBProject.VBComponents(source.stem).CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(source.read_text(encoding='utf-8-sig'))
 cm=w.VBProject.VBComponents('FEMEngine').CodeModule;code=cm.Lines(1,cm.CountOfLines)
 start=re.search(r'^Private Function P3TryAssociatedPotentialWatchdog\b',code,re.M|re.I).start()
 end=start+re.search(r'^End Function',code[start:],re.M|re.I).end()
 fragment,count=re.subn(r'\bBandSolver\s*\(\s*\)','LinearRejectBandForTest()',code[start:end],flags=re.I);assert count==1
 before,restore=fragment.split('RestoreTrial:',1)
 restore,count=re.subn(r'\bP3RebuildTangentFromSpmat\s*\(\s*\)','LinearRejectRestoreForTest()',restore,count=1,flags=re.I);assert count==1
 declarations,procedures=driver.split('Private Function WatchdogDampedBandForTest',1)
 code=code[:start]+before+'RestoreTrial:'+restore+code[end:]
 code=code.replace('Option Explicit','Option Explicit\r\n'+declarations,1)+'\r\nPrivate Function WatchdogDampedBandForTest'+procedures
 cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(code)
 prefix="'"+w.Name+"'!"
 for key,value in [('FLOW_POLICY','DAVIS'),('NATIVE_DLL',a.native),('NATIVE_SPARSE_DIRECT',0.),('ASSOCIATED_ENERGY_RECOVERY',1.),('ASSOCIATED_ENERGY_FIRST',1.),('EXPORT_LOAD','')]:app.Run(prefix+('FEMWriteTextSetting' if isinstance(value,str) else 'FEMWriteNumericSetting'),key,value)
 cc=app.VBE.CommandBars.FindControl(1,578)
 if cc.Enabled:cc.Execute()
 if app.VBE.CommandBars.FindControl(1,578).Enabled:
  pane=app.VBE.ActiveCodePane
  print('COMPILE_LOCATION',pane.CodeModule.Parent.Name,pane.GetSelection(),flush=True)
  print(pane.CodeModule.Lines(max(1,pane.GetSelection()[0]-2),5),flush=True)
  raise RuntimeError('Probe compilation failed')
 v=list(app.Run(prefix+'AssociatedLinearRejectAudit',a.failure));assert len(v)==16,v
 record=dict(zip(('rejected','status_after_trial','error_after_trial','fatal_after_trial','rollback_displacement_error','rollback_force_error','rollback_tangent_error','committed_displacement_change','fallback_accepted','fallback_residual_ratio','fallback_fatal','fallback_status','committed_increment_change','injections','note_after_trial','fallback_note'),v))
 record.update(failure=a.failure,native=a.native,expect_old_fatal=a.expect_old_fatal)
 (a.out/'results.json').write_text(json.dumps(record,indent=2),encoding='utf-8');print(json.dumps(record),flush=True)
 assert v[0] and v[4]==0 and v[7]==0 and v[12]==0 and v[13]==1,v
 if a.failure=='singular' and not a.expect_old_fatal:
  assert not v[3] and v[2]==0 and v[5]<1e-12 and v[6]<1e-8 and v[8] and v[9]<1 and not v[10] and v[14]==v[15]=='',v
 else:
  expected={'singular':('GLOBAL_SINGULAR',-2147218200),'restore_singular':('GLOBAL_SINGULAR',-2147218200),'runtime':('RUNTIME_ERROR',5),'capacity':('CAPACITY_ERROR',7),'input':('INPUT_ERROR',-2147218503),'cancel':('NONCONVERGED',18)}[a.failure]
  assert v[3] and not v[8] and v[10] and v[1]==v[11]==expected[0] and v[2]==expected[1] and v[14] and v[15],v
finally:
 if w is not None:w.Close(False)
 app.Quit()
