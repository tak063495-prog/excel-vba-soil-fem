"""Audit the actual VBA accepted-step policy and strict initial-step input."""
from pathlib import Path
import argparse, json, shutil
import win32com.client

ap=argparse.ArgumentParser()
ap.add_argument('--book',type=Path,required=True)
ap.add_argument('--out',type=Path,required=True)
a=ap.parse_args(); root=Path(__file__).resolve().parents[1]
a.out=a.out.resolve(); a.out.mkdir(parents=True,exist_ok=False)
dest=a.out/'static_step_audit.xlsm'; shutil.copy2(a.book,dest)
app=win32com.client.DispatchEx('Excel.Application')
app.DisplayAlerts=False; app.EnableEvents=False; w=None
try:
 w=app.Workbooks.Open(str(dest),UpdateLinks=0)
 for source in (root/'src/vba').glob('*.bas'):
  try:c=w.VBProject.VBComponents(source.stem)
  except Exception:c=w.VBProject.VBComponents.Add(1); c.Name=source.stem
  cm=c.CodeModule; cm.DeleteLines(1,cm.CountOfLines); cm.AddFromString(source.read_text(encoding='utf-8-sig'))
 cm=w.VBProject.VBComponents('FEMEngine').CodeModule
 cm.AddFromString('''
Public Function StaticStepAudit(ByVal n As Long, ByVal ratio As Double) As Variant
 Dim step As Double,action As String
 step=0.1
 P3AcceptedNewtonStep step,n,ratio,action
 StaticStepAudit=Array(step,action)
End Function
Public Function StaticInitialAudit(ByVal legacy As Double) As Variant
 On Error GoTo Invalid
 StaticInitialAudit=P3ConfiguredInitialStep(legacy)
 Exit Function
Invalid:
 StaticInitialAudit="ERROR|" & CStr(Err.Number)
End Function
''')
 cc=app.VBE.CommandBars.FindControl(1,578)
 if cc.Enabled:cc.Execute()
 assert not app.VBE.CommandBars.FindControl(1,578).Enabled,'VBA compile failed'
 prefix="'"+w.Name+"'!"; results=[]
 for flag in (0,1):
  app.Run(prefix+'FEMWriteNumericSetting','STATIC_KEEP_ACCEPTED_STEP',float(flag))
  for iterations,legacy,on in ((2,.15,.15),(3,.1,.1),(5,.1,.1),(6,.07,.1),(8,.07,.1),(9,.05,.1),(1000,.05,.1)):
   step,action=app.Run(prefix+'StaticStepAudit',iterations,1.)
   assert abs(step-(on if flag else legacy))<1e-14,(flag,iterations,step,action)
   if flag and iterations>5:assert action=='KEEP_CONVERGED'
   results.append({'keep':flag,'iterations':iterations,'next_step':step,'action':action})
  step,action=app.Run(prefix+'StaticStepAudit',4,1.1)
  assert abs(step-.11)<1e-14 and action=='RECOVER_1_1'
 for count in (0,1,10,20,100000,-1,.5,100001):
  app.Run(prefix+'FEMWriteNumericSetting','STATIC_INITIAL_STEPS',float(count))
  value=app.Run(prefix+'StaticInitialAudit',.05)
  if count<0 or count>100000 or int(count)!=count:assert str(value).startswith('ERROR|'),(count,value)
  else:assert abs(value-(1./count if count else .05))<1e-14,(count,value)
  results.append({'initial_steps':count,'value':value})
 record={'status':'PASS','cases':results}
 (a.out/'results.json').write_text(json.dumps(record,indent=2),encoding='utf-8')
 print('PASS',len(results),'actual VBA step-policy cases',flush=True)
finally:
 if w is not None:w.Close(False)
 app.Quit()
