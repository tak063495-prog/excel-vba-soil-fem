"""Reject old DAVIS SRM operators while preserving compatible INCONSISTENT saves."""
from pathlib import Path
import argparse,json,shutil,win32com.client
ap=argparse.ArgumentParser();ap.add_argument('--legacy-runs',type=Path,required=True)
ap.add_argument('--out',type=Path,required=True);a=ap.parse_args()
root=Path(__file__).resolve().parents[1];a.out.mkdir(parents=True,exist_ok=False)
records=[]
for case in ('davis_elastic_auto','inconsistent_elastic_auto'):
 folder=a.legacy_runs.resolve()/case
 checkpoint=folder/(case+'_out')/'stg02_SRM'
 assert checkpoint.exists(),checkpoint
 case_dir=a.out.resolve()/case;case_dir.mkdir()
 dest=case_dir/(case+'.xlsm');shutil.copy2(folder/'fresh_results.xlsm',dest)
 for name in ('soilfem_native_x64.dll','soilfem_sparse_x64.dll'):
  if (folder/name).exists():shutil.copy2(folder/name,dest.parent/name)
 app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;w=None
 try:
  w=app.Workbooks.Open(str(dest),UpdateLinks=0)
  for ws in w.Worksheets:
   if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
  for source in (root/'src/vba').glob('*.bas'):
   cm=w.VBProject.VBComponents(source.stem).CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(source.read_text(encoding='utf-8-sig'))
  cm=w.VBProject.VBComponents.Add(1).CodeModule;cm.AddFromString('''
Public Function HourglassLegacyView(ByVal folder As String) As String
 On Error GoTo Failed
 SuppressUserMessages=True:P3LoadTimeFrame folder
 HourglassLegacyView="PASS":Exit Function
Failed:HourglassLegacyView="ERROR|" & Err.Description
End Function
Public Function HourglassLegacyResume(ByVal folder As String) As String
 On Error GoTo Failed
 SuppressUserMessages=True:FEMWriteTextSetting "EXPORT_LOAD",folder
 P0_RunAnalysis
 HourglassLegacyResume=ResultStatus & "|" & AnalysisMessage:Exit Function
Failed:HourglassLegacyResume="ERROR|" & Err.Description
End Function
''')
  control=app.VBE.CommandBars.FindControl(1,578)
  if control.Enabled:control.Execute()
  assert not app.VBE.CommandBars.FindControl(1,578).Enabled
  prefix="'"+w.Name+"'!"
  view=app.Run(prefix+'HourglassLegacyView',str(checkpoint))
  resume=app.Run(prefix+'HourglassLegacyResume',str(checkpoint))
  if case.startswith('davis'):
   assert view.startswith('ERROR|') and '署名' in view,view
   assert not resume.startswith('PASS|') and '署名' in resume,resume
  else:
   assert view=='PASS' and resume.startswith('PASS|'),(view,resume)
  record={'case':case,'frame':view,'resume':resume,'status':'PASS'};records.append(record)
  print(json.dumps(record),flush=True)
 finally:
  if w is not None:w.Close(False)
  app.Quit()
(a.out/'results.json').write_bytes(json.dumps({'status':'PASS','cases':records},indent=2).encode('utf-8'))
