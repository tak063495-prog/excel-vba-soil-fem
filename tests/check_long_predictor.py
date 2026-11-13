"""Exercise long V1 history on an elastic face_support/static fixture workbook.

The main FE correction is deliberately damped; physical force and material
responses remain real. Use a workbook initialized from that public fixture.
"""
from pathlib import Path
import argparse,json,shutil
import numpy as np
import win32com.client

ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True)
ap.add_argument('--failure',choices=['none','singular','runtime','capacity','input','cancel'],default='none')
ap.add_argument('--expect-old-cutback','--expect-old-singular',dest='expect_old_singular',action='store_true');a=ap.parse_args()
root=Path(__file__).resolve().parents[1];a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=False)
shutil.copy2(a.book.parent/'soilfem_native_x64.dll',a.out/'soilfem_native_x64.dll')
driver='''
Private HistoryFailureMode As String,HistoryFailureUsed As Boolean,HistoryGuardChecks As Long
Private Sub LongHistoryCheckedRemember(ByRef base() As Double,ByVal delta As Double,ByVal iterations As Long,ByVal retries As Long)
 Dim oldFS As Boolean,oldConsol As Boolean,oldMixed As Boolean,oldRetry As Boolean,oldJoint As Boolean
 AccelRememberIncrement base,delta,iterations,retries
 If Not AccelHistoryReady Or iterations<=5 Then Exit Sub
 oldFS=FSActive:oldConsol=P6ConsolActive:oldMixed=P6MixedUP:oldRetry=AccelBaselineRetry:oldJoint=P3HasJointElements
 FSActive=True:AccelRememberIncrement base,delta,iterations,retries
 If AccelHistoryReady Then Err.Raise 1201,,"Coupled history retained"
 FSActive=oldFS:HistoryGuardChecks=HistoryGuardChecks+1
 P6ConsolActive=True:AccelRememberIncrement base,delta,iterations,retries
 If AccelHistoryReady Then Err.Raise 1202,,"Transient history retained"
 P6ConsolActive=oldConsol:HistoryGuardChecks=HistoryGuardChecks+1
 P6MixedUP=True:AccelRememberIncrement base,delta,iterations,retries
 If AccelHistoryReady Then Err.Raise 1203,,"Mixed history retained"
 P6MixedUP=oldMixed:HistoryGuardChecks=HistoryGuardChecks+1
 AccelRememberIncrement base,delta,iterations,1
 If AccelHistoryReady Then Err.Raise 1204,,"Retried history retained"
 HistoryGuardChecks=HistoryGuardChecks+1
 AccelBaselineRetry=True:AccelRememberIncrement base,delta,iterations,retries
 If AccelHistoryReady Then Err.Raise 1205,,"Baseline retry retained"
 AccelBaselineRetry=oldRetry:HistoryGuardChecks=HistoryGuardChecks+1
 AccelRememberIncrement base,0#,iterations,retries
 If AccelHistoryReady Then Err.Raise 1206,,"Zero-step history retained"
 HistoryGuardChecks=HistoryGuardChecks+1
 P3HasJointElements=True:AccelRememberIncrement base,delta,iterations,retries
 If AccelHistoryReady Then Err.Raise 1207,,"Joint history retained"
 P3HasJointElements=oldJoint:HistoryGuardChecks=HistoryGuardChecks+1
 AccelRememberIncrement base,delta,iterations,retries
 If Not AccelHistoryReady Then Err.Raise 1208,,"Valid accepted history lost"
End Sub
Private Function HistoryDampedBand(ByVal predicted As Boolean) As Boolean
 Dim i As Long
 If predicted And P3SrmTrialRunning And HistoryFailureMode<>"none" And Not HistoryFailureUsed Then
  HistoryFailureUsed=True
  Select Case HistoryFailureMode
   Case "singular":SetAnalysisFailure RESULT_GLOBAL_SINGULAR,"Injected predicted-state pivot",vbObjectError+3304,-1,-1,CurrentIncrement,CurrentIteration
   Case "runtime":SetAnalysisFailure RESULT_RUNTIME_ERROR,"Injected runtime",5,-1,-1,CurrentIncrement,CurrentIteration
   Case "capacity":SetAnalysisFailure RESULT_CAPACITY_ERROR,"Injected capacity",7,-1,-1,CurrentIncrement,CurrentIteration
   Case "input":SetAnalysisFailure RESULT_INPUT_ERROR,"Injected input",vbObjectError+3001,-1,-1,CurrentIncrement,CurrentIteration
   Case "cancel":SetAnalysisFailure RESULT_NONCONVERGED,"Injected cancel",18,-1,-1,CurrentIncrement,CurrentIteration
  End Select
  Exit Function
 End If
 If Not BandSolver() Then Exit Function
 ' Deliberate damping creates genuinely accepted >5-iteration increments.
 ' Stress, force, material, prescribed displacements and tolerances are intact.
 For i=0 To lastDof:Disp(i)=0.5*Disp(i):Next i
 HistoryDampedBand=True
End Function
Public Function LongHistoryAudit(ByVal mode As String) As Variant
 On Error GoTo Failed
 HistoryFailureMode=mode:HistoryFailureUsed=False:HistoryGuardChecks=0
 SuppressUserMessages=True
 P0_RunAnalysis
 LongHistoryAudit=Array(ResultStatus,FSS,RelativeResidualFree,P6PerfNewtonCount,AccelPredictCount,AccelFallbackCount,P6PerfCutbackCount,AnalysisErrorNumber,HistoryFailureUsed,P3SrmTrialFailNote,P3UserCancel,P3SearchFatal,HistoryGuardChecks)
 Exit Function
Failed:LongHistoryAudit=Array("ERROR|" & CStr(Err.Number) & "|" & Err.Description)
End Function
Public Function LongHistoryFields() As Variant
 Dim v() As Double,i As Long,e As Long,g As Long,k As Variant,p As Long
 ReDim v(lastDof+NumberOfElement*16):p=0
 For i=0 To lastDof:v(p)=TDisp(i):p=p+1:Next i
 For e=0 To NumberOfElement-1:For g=0 To 3:For Each k In Array(0,1,2,16)
  v(p)=Elem(e).Stmat(CLng(k),g):p=p+1
 Next k:Next g:Next e
 LongHistoryFields=v
End Function
'''
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;w=None;records=[];reference=None
if a.failure=='none':cases=[(f'{native}_long{long}',native,long,'none') for native in ('AUTO','OFF') for long in (0,1)]
elif a.failure=='singular':cases=[('clean_reference','AUTO',1,'none'),('singular','AUTO',1,'singular')]
else:cases=[(a.failure,'AUTO',1,a.failure)]
try:
 for name,native,long,failure in cases:
  dest=a.out/(name+'.xlsm');shutil.copy2(a.book,dest);w=app.Workbooks.Open(str(dest),UpdateLinks=0)
  for ws in w.Worksheets:
   if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
  for source in (root/'src/vba').glob('*.bas'):
   text=source.read_text(encoding='utf-8-sig')
   if source.stem=='FEMEngine':
    if a.expect_old_singular:
     start=text.index("        ' This operator belongs to an uncommitted predicted increment.")
     end=text.index("        ' Acceleration failures do not establish SRM FAIL.",start)
     text=text[:start]+text[end:]
    marker='        If Not BandSolver() Then\n';assert text.count(marker)==1
    text=text.replace(marker,'        If Not HistoryDampedBand(predictorUsed) Then\n',1)
    marker='        AccelRememberIncrement incrementBase,';assert text.count(marker)==1
    text=text.replace(marker,'        LongHistoryCheckedRemember incrementBase,',1)
   cm=w.VBProject.VBComponents(source.stem).CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(text)
  w.VBProject.VBComponents('FEMEngine').CodeModule.AddFromString(driver)
  ws=w.Worksheets('ステージ');ws.Range('A2:F1000').ClearContents();ws.Range('A2:F3').Value2=((1,'GRAVITY','',1,1,'gravity'),(2,'SRM','',1,1,'fixed Fs1 replay'))
  w.Worksheets('水条件').Range('B3').Value2='NONE'
  prefix="'"+w.Name+"'!"
  settings={'FLOW_POLICY':'DAVIS','NATIVE_DLL':native,'NATIVE_SPARSE_DIRECT':0,'SRM_FIXED_FS':1.,'STATIC_INITIAL_STEPS':10.,'STATIC_KEEP_ACCEPTED_STEP':1.,'ACCEL_V1_PREDICTOR':1.,'ACCEL_PREDICT_LONG_HISTORY':float(long),'ACCEL_PREDICT_BETA':1. if a.failure=='none' else .5,'ACCEL_V2A_REUSE':0.,'ACCEL_V2B_COST':0.,'ACCEL_V4_ANDERSON':0.,'ACCEL_V5_GMRES_LU':0.,'ACCEL_ADAPTIVE':0.,'ASSOCIATED_ENERGY_RECOVERY':0.,'ASSOCIATED_ENERGY_FIRST':0.,'EXPORT_LOAD':'','EXPORT_MODE':'OFF'}
  for key,val in settings.items():app.Run(prefix+('FEMWriteTextSetting' if isinstance(val,str) else 'FEMWriteNumericSetting'),key,val)
  cc=app.VBE.CommandBars.FindControl(1,578)
  if cc.Enabled:cc.Execute()
  if app.VBE.CommandBars.FindControl(1,578).Enabled:
   pane=app.VBE.ActiveCodePane;raise RuntimeError((pane.CodeModule.Parent.Name,pane.GetSelection()))
  values=list(app.Run(prefix+'LongHistoryAudit',failure));assert len(values)==13,values
  record=dict(zip(('status','fs','relative_residual','newton_work','predictions','baseline_retries','cutbacks','error','failure_injected','failure_memo','cancel','search_fatal','history_rejection_controls'),values));record.update(case=name,native=native,long_history=long)
  if long:assert record['history_rejection_controls']>=7,record
  if failure=='none' or failure=='singular':
   assert record['status']=='PASS' and record['fs']==1. and record['relative_residual']<=1e-5 and record['error']==0 and not record['cancel'] and not record['search_fatal'],record
   fields=np.asarray(app.Run(prefix+'LongHistoryFields'),dtype=float);assert np.isfinite(fields).all()
   if reference is None:reference=fields
   record['relative_completed_field_difference']=float(np.linalg.norm(fields-reference)/max(np.linalg.norm(reference),1e-30))
   assert record['relative_completed_field_difference']<1e-4,record
   settings_rows=w.Worksheets('設定').Range('A5:E500').Value2
   row=next(i+5 for i,r in enumerate(settings_rows) if r[4]=='ACCEL_PREDICT_LONG_HISTORY')
   record['ui_choices']=w.Worksheets('設定').Cells(row,3).Validation.Formula1
   assert record['ui_choices']=='0,1',record
   if failure=='singular':
    assert record['failure_injected'],record
    if a.expect_old_singular:assert record['baseline_retries']==0 and record['cutbacks']==1 and 'GLOBAL_SINGULAR' in record['failure_memo'],record
    else:assert record['baseline_retries']==1 and record['cutbacks']==0 and record['failure_memo']=='',record
  else:
   expected={'singular':('GLOBAL_SINGULAR',-2147218200),'runtime':('RUNTIME_ERROR',5),'capacity':('CAPACITY_ERROR',7),'input':('INPUT_ERROR',-2147218503),'cancel':('NONCONVERGED',18)}[a.failure]
   assert (record['status'],record['error'])==expected and record['failure_injected'] and record['fs']==0. and record['baseline_retries']==0,record
   if a.failure=='cancel':assert record['cancel'] and record['search_fatal'],record
  records.append(record);(a.out/'results.json').write_text(json.dumps(records,indent=2,allow_nan=False),encoding='utf-8');print('PASS',json.dumps(record),flush=True)
  w.Close(False);w=None
 if a.failure=='none':
  for native in ('AUTO','OFF'):
   off,on=[r for r in records if r['native']==native]
   assert on['predictions']>off['predictions'] and on['newton_work']<.5*off['newton_work'],(off,on)
finally:
 if w is not None:w.Close(False)
 app.Quit()
