"""Ensure the actual VBA sparse-fallback branch releases its native factor."""
from pathlib import Path
import argparse,json,shutil
import win32com.client

ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True)
a=ap.parse_args();root=Path(__file__).resolve().parents[1]
a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=False)
copy=a.out/'sparse_release_audit.xlsm';shutil.copy2(a.book,copy)
shutil.copy2(root/'native/bin/combined/soilfem_native_x64.dll',a.out/'soilfem_native_x64.dll')
shutil.copy2(root/'native/bin/sparse/soilfem_sparse_x64.dll',a.out/'soilfem_sparse_x64.dll')
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;w=None
try:
 w=app.Workbooks.Open(str(copy),UpdateLinks=0)
 for ws in w.Worksheets:
  if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
 for source in (root/'src/vba').glob('*.bas'):
  try:c=w.VBProject.VBComponents(source.stem)
  except Exception:c=w.VBProject.VBComponents.Add(1);c.Name=source.stem
  cm=c.CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(source.read_text(encoding='utf-8-sig'))
 cm=w.VBProject.VBComponents('FEMSparse').CodeModule;cm.AddFromString('''
Public Function SparseReleaseAudit() As Variant
 On Error GoTo Failed
 Dim nf As Boolean,rc As Long,oldLast As Long,oldToken As LongLong,released As Boolean,staleRC As Long
 SuppressUserMessages=True
 P0_RunAnalysis
 If ResultStatus<>RESULT_PASS Then Err.Raise 1001,,AnalysisMessage
 FEMSparseReset
 rc=FEMSparseTryBandSolve(nf)
 If rc<>1 Or mToken=0 Then Err.Raise 1002,,"Native factor was not established: " & FEMSparseReason
 oldToken=mToken:oldLast=lastDof
 ' Fail a guarded dimension check while a real native factor is live.
 lastDof=nDof-2
 rc=FEMSparseTryBandSolve(nf)
 lastDof=oldLast
 released=(mToken=0 And Not mUsed)
 staleRC=SparseSolve(oldToken,nDof,VarPtr(P6ForceBeforeBoundary(0)),VarPtr(Disp(0)),0.0000000001,VarPtr(Disp(0)),VarPtr(mMetrics(0)))
 SparseReleaseAudit=Array("PASS",rc,released,staleRC,FEMSparseFactorCalls,mRejectedGeneration=P6TangentGeneration)
 Exit Function
Failed:
 SparseReleaseAudit=Array("ERROR|" & Err.Description)
End Function
''')
 prefix="'"+w.Name+"'!"
 app.Run(prefix+'FEMWriteNumericSetting','NATIVE_SPARSE_DIRECT',1.)
 app.Run(prefix+'FEMWriteTextSetting','NATIVE_DLL','AUTO')
 app.Run(prefix+'FEMWriteTextSetting','EXPORT_LOAD','')
 cc=app.VBE.CommandBars.FindControl(1,578)
 if cc.Enabled:cc.Execute()
 assert not app.VBE.CommandBars.FindControl(1,578).Enabled,'VBA compile failed'
 result=list(app.Run(prefix+'SparseReleaseAudit'))
 assert result[0]=='PASS' and result[1]==2 and result[2] and result[3]==4 and result[4]>0 and result[5],result
 record={'status':'PASS','fallback_code':result[1],'token_cleared':result[2],'released_token_solve_rc':result[3],'factor_calls':result[4],'rejected_generation_remembered':result[5]}
 (a.out/'results.json').write_text(json.dumps(record,indent=2),encoding='utf-8');print(json.dumps(record),flush=True)
finally:
 if w is not None:w.Close(False)
 app.Quit()
