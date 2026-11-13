"""Verify static-prefix reuse and coarse scans through real completed searches."""
from pathlib import Path
import argparse,json,shutil
import numpy as np
import win32com.client

ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--only-case')
a=ap.parse_args();root=Path(__file__).resolve().parents[1];a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=False)
shutil.copy2(a.book.resolve().parent/'soilfem_native_x64.dll',a.out/'soilfem_native_x64.dll')
sources={p.stem:p.read_text(encoding='utf-8-sig') for p in (root/'src/vba').glob('*.bas')}
driver='''
Public Function SearchOptionRun() As Variant
 On Error GoTo Failed
 SuppressUserMessages=True:P0_RunAnalysis
 SearchOptionRun=Array(ResultStatus,FSS,P3SrmTrialCount,P6SrmFs1ReuseCount,RelativeResidualFree,AnalysisErrorNumber,SearchReuseGuardChecks):Exit Function
Failed:SearchOptionRun=Array("ERROR|" & Err.Description)
End Function
Public Function SearchOptionFields() As Variant
 Dim v() As Double,i As Long,e As Long,g As Long,p As Long,k As Variant
 ReDim v(lastDof+NumberOfElement*16):p=0
 For i=0 To lastDof:v(p)=TDisp(i):p=p+1:Next i
 For e=0 To NumberOfElement-1:For g=0 To 3:For Each k In Array(0,1,2,16)
  v(p)=Elem(e).Stmat(CLng(k),g):p=p+1
 Next k:Next g:Next e
 SearchOptionFields=v
End Function
'''
guards='''
Public SearchReuseGuardChecks As Long
Private Function SearchReuseCheckedForTest() As Boolean
 Dim eligible As Boolean,id As Long,kind As String,factor As Double,residual As Double,coupled As Boolean
 eligible=P3CanReuseStaticPrefixForFs1()
 If eligible Then
  id=P3LastCompletedStage:P3LastCompletedStage=-1
  If P3CanReuseStaticPrefixForFs1() Then Err.Raise 1001,,"Partial prefix was reused"
  P3LastCompletedStage=id:SearchReuseGuardChecks=SearchReuseGuardChecks+1
  kind=P3StageKind(1):P3StageKind(1)="MATSET"
  If P3CanReuseStaticPrefixForFs1() Then Err.Raise 1002,,"Construction prefix was reused"
  P3StageKind(1)=kind:SearchReuseGuardChecks=SearchReuseGuardChecks+1
  factor=P3CurrentStrengthFactor:P3CurrentStrengthFactor=1.1
  If P3CanReuseStaticPrefixForFs1() Then Err.Raise 1003,,"Nonunit strength was reused"
  P3CurrentStrengthFactor=factor:SearchReuseGuardChecks=SearchReuseGuardChecks+1
  residual=RelativeResidualFree:RelativeResidualFree=2#*P3_ENGINEERING_RESIDUAL
  If P3CanReuseStaticPrefixForFs1() Then Err.Raise 1004,,"Unconverged prefix was reused"
  RelativeResidualFree=residual:SearchReuseGuardChecks=SearchReuseGuardChecks+1
  coupled=FSActive:FSActive=True
  If P3CanReuseStaticPrefixForFs1() Then Err.Raise 1005,,"Coupled prefix was reused"
  FSActive=coupled:SearchReuseGuardChecks=SearchReuseGuardChecks+1
 End If
 SearchReuseCheckedForTest=eligible
End Function
'''
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;w=None;records=[];reference=None
cases=[(f'reuse{reuse}_step{step}',reuse,step,True) for reuse in (0,1) for step in (.1,.5)]
cases += [(f'invalid_{i}',1,value,False) for i,value in enumerate((0.,-.1,1.1,1e100))]
# Existing numeric-setting reconstruction replaces nonnumeric cells with their
# registered default before solving. Verify that behavior explicitly.
cases += [('normalized_text',1,'NaN',True)]
if a.only_case:cases=[case for case in cases if case[0]==a.only_case];assert cases
try:
 for name,reuse,step,valid in cases:
  dest=a.out/(name+'.xlsm');shutil.copy2(a.book,dest);w=app.Workbooks.Open(str(dest),UpdateLinks=0)
  for ws in w.Worksheets:
   if ws.ProtectContents:ws.Unprotect('')
  for module,text in sources.items():
   cm=w.VBProject.VBComponents(module).CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(text)
  cm=w.VBProject.VBComponents('FEMEngine').CodeModule;cm.AddFromString(guards)
  code=cm.Lines(1,cm.CountOfLines);needle='reuseStaticPrefix = P3CanReuseStaticPrefixForFs1()';assert code.count(needle)==1
  code=code.replace(needle,'reuseStaticPrefix = SearchReuseCheckedForTest()');cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(code)
  nodes=w.Worksheets('節点データ');last=nodes.Cells(nodes.Rows.Count,1).End(-4162).Row;rows=nodes.Range(f'A2:I{last}').Value2
  top=max(float(r[2]) for r in rows)
  candidates=[(float(r[1]),int(r[0])) for r in rows if float(r[2])==top and float(r[4] or 0)==0]
  assert candidates
  node=max(candidates)[1]
  ws=w.Worksheets('載荷');ws.Range('A2:K1000').ClearContents();ws.Range('A2:J2').Value2=((1,node,'Y','FORCE',-.001,None,None,None,None,'small independent load'),)
  ws=w.Worksheets('ステージ');ws.Range('A2:F1000').ClearContents();ws.Range('A2:F4').Value2=((1,'GRAVITY','',1,1,'gravity'),(2,'LOAD','1',1,1,'nonzero applied load'),(3,'SRM','',1,1,'full search'))
  mod=w.VBProject.VBComponents.Add(1);mod.Name='SearchOptionsAudit';mod.CodeModule.AddFromString(driver)
  prefix="'"+w.Name+"'!"
  settings={'FLOW_POLICY':'DAVIS','NATIVE_DLL':'AUTO','NATIVE_SPARSE_DIRECT':0.,'SRM_FIXED_FS':0.,'SRM_FMAX':2.2,'SRM_SCAN_STEP':step,'SRM_REUSE_STATIC_FS1':float(reuse),'EXPORT_LOAD':'','EXPORT_MODE':'OFF'}
  for key,value in settings.items():app.Run(prefix+('FEMWriteTextSetting' if isinstance(value,str) else 'FEMWriteNumericSetting'),key,value)
  cc=app.VBE.CommandBars.FindControl(1,578)
  if cc.Enabled:cc.Execute()
  assert not app.VBE.CommandBars.FindControl(1,578).Enabled
  v=list(app.Run(prefix+'SearchOptionRun'));assert len(v)==7,v
  record=dict(zip(('status','fs','trials','fs1_reuses','relative_residual','error','reuse_rejection_controls'),v));record.update(case=name,requested_reuse=reuse,requested_step=step)
  if valid:
   effective_step=float(app.Run(prefix+'P6ReadSetting','SRM_SCAN_STEP',.1));record['effective_step']=effective_step
   assert effective_step==(.1 if isinstance(step,str) else step),record
   assert v[0]=='PASS' and abs(v[1]-2.2)<1e-12 and v[2]==(13 if effective_step==.1 else 4) and v[3]==reuse and v[4]<=1e-5 and v[6]==5*reuse,record
   field=np.array(app.Run(prefix+'SearchOptionFields'),dtype=float)
   if reference is None:reference=field
   record['max_completed_field_difference']=float(max(abs(field-reference)))
   assert record['max_completed_field_difference']<1e-10,record
  else:assert v[0]=='INPUT_ERROR' and v[1]==0 and v[2]==0 and v[5]==-2147218213,record
  records.append(record);(a.out/'results.json').write_text(json.dumps(records,indent=2),encoding='utf-8');print('PASS',json.dumps(record),flush=True)
  w.Close(False);w=None
finally:
 if w is not None:w.Close(False)
 app.Quit()
