from pathlib import Path
import argparse,hashlib,json,shutil,time,csv
import win32com.client,win32process
R=Path(__file__).resolve().parent
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);args=ap.parse_args()
source=args.book.resolve();folder=args.out.resolve();folder.mkdir(parents=True,exist_ok=False)
target=folder/'gravity_verification.xlsm';shutil.copyfile(source,target)
shutil.copyfile(source.parent/'soilfem_band_fast_x64.dll',folder/'soilfem_band_fast_x64.dll')
initial_hash=hashlib.sha256(source.read_bytes()).hexdigest()
ex=win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Excel.Application'))
ex.Visible=False;ex.DisplayAlerts=False;ex.EnableEvents=False;ex.AutomationSecurity=1
try:
 w=ex.Workbooks.Open(str(target),UpdateLinks=0)
 q=w.VBProject.VBComponents.Add(1);q.Name='GravityVerification'
 q.CodeModule.AddFromString('''Option Explicit
Public GravityVerificationRecord As Variant
Public Sub CaptureGravityVerification(ByVal passed As Boolean)
 GravityVerificationRecord=Array(passed,ResultStatus,AnalysisMessage,P3LastConvergedLoadFactor,P3SuccessfulIncrementCount,RelativeResidualFree,P3_ENGINEERING_RESIDUAL,P3_ENGINEERING_CORRECTION,P3_RESIDUAL_TOLERANCE,P3_MAX_GLOBAL_ITERATIONS,P3_LINESEARCH_MAX,P6FactorizationCount,FEMBandFastFactorCalls,FEMBandFastSolveCalls,P3SrmTrialCount)
 Err.Raise vbObjectError+7450,,"EXPECTED_GRAVITY_VERIFICATION_BOUNDARY"
End Sub
Public Function VerifyGravity() As Variant
 SuppressUserMessages=True
 P0_RunAnalysis
 VerifyGravity=Array(ResultStatus,AnalysisMessage,GravityVerificationRecord)
End Function
Public Sub GravityVerificationTrace(ByVal loadFactor As Double,ByVal iteration As Long,ByVal residual As Double,ByVal correction As Double)
 Dim f As Integer
 f=FreeFile:Open ThisWorkbook.Path & "\\newton_trace.csv" For Append As #f
 Print #f,CStr(loadFactor) & "," & CStr(iteration) & "," & CStr(residual) & "," & CStr(correction)
 Close #f
End Sub
''')
 cm=w.VBProject.VBComponents('FEMEngine').CodeModule;code=cm.Lines(1,cm.CountOfLines)
 for needle,passed in [('      P6PerfEmit "GRAVITY", "PASS"',True),('          P6PerfEmit "GRAVITY", "DEFER"',False),('      If Not P3ReplayQuiet Then P6PerfEmit "GRAVITY", "FAIL"',False)]:
  assert code.count(needle)==1,needle
  code=code.replace(needle,needle+'\r\n      CaptureGravityVerification '+str(passed),1)
 code=code.replace('P3NextGlobalIteration:\r\n','P3NextGlobalIteration:\r\n      GravityVerificationTrace targetFactor,localIteration,RelativeResidualFree,correctionRatio\r\n',1)
 cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(code)
 control=ex.VBE.CommandBars.FindControl(1,578)
 if control.Enabled:control.Execute()
 assert not ex.VBE.CommandBars.FindControl(1,578).Enabled
 print('GRAVITY_VERIFICATION_STARTED',win32process.GetWindowThreadProcessId(ex.Hwnd)[1],str(folder),flush=True)
 t=time.time();result=ex.Run("'"+w.Name+"'!VerifyGravity")
 assert 'EXPECTED_GRAVITY_VERIFICATION_BOUNDARY' in result[1],result
 data=list(result[2]);logs=folder/'gravity_verification_out'
 rows=list(csv.DictReader((logs/'increment_summary.csv').open(encoding='cp932')))
 manifest=list(csv.reader((logs/'manifest.csv').open(encoding='cp932')))
 accepted=[r for r in rows if r.get('accepted')=='True']
 for increment in accepted:
  assert int(increment['iterations'])<=data[9],increment
  assert float(increment['relative_residual'])<=data[6],increment
  if increment['correction_ratio']!='NA':
   assert float(increment['correction_ratio'])<=data[7] or float(increment['relative_residual'])<=data[8],increment
 passed=bool(data[0])
 record=dict(passed=passed,verification_scope='Original stage 1 GRAVITY; intentional diagnostic stop at stage boundary before SRM. No input/stage/acceptance changes.',elapsed_seconds=time.time()-t,source_sha256=initial_hash,state=data,last_increment=rows[-1] if rows else None,accepted_increments=len(accepted),manifest=manifest,logs=str(logs),source_unchanged=hashlib.sha256(source.read_bytes()).hexdigest()==initial_hash)
 record['max_iterations_per_increment']=max((int(r['iterations']) for r in accepted),default=0)
 record['cutbacks']=sum(r.get('accepted')!='True' for r in rows)
 (folder/'gravity_raw_capture.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf8')
 assert record['source_unchanged']
 if passed:
  # P3LastConvergedLoadFactor is updated only in APPLY, not GRAVITY.
  # Use the accepted gravity increment and its exported checkpoint instead.
  assert data[14]==0
  assert manifest[-1][4]=='SUCCESS'
  assert float(accepted[-1]['lambda_target'])==1
  assert float(accepted[-1]['relative_residual'])<=data[6]
  assert float(accepted[-1]['correction_ratio'])<=data[7] or float(accepted[-1]['relative_residual'])<=data[8]
  record['load_factor']=float(accepted[-1]['lambda_target'])
  record['srm_trials']=data[14]
 (folder/'gravity_result.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf8')
 print('GRAVITY_VERIFICATION_RESULT',json.dumps(record,ensure_ascii=False),flush=True)
 assert passed,'Original gravity stage did not converge'
finally:
 for w in list(ex.Workbooks):w.Close(False)
 ex.Quit()
