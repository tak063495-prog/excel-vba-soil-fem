"""Audit rejected plastic trial restoration with a compiled face_support/static fixture book."""
from pathlib import Path
import argparse,json,shutil,win32com.client
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--native',choices=['AUTO','OFF'],default='AUTO');ap.add_argument('--expect-stale',action='store_true');a=ap.parse_args();root=Path(__file__).resolve().parents[1];a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=True);dst=a.out/'tangent_restore_audit.xlsm';assert not dst.exists();shutil.copy2(a.book,dst);shutil.copy2(a.book.resolve().parent/'soilfem_native_x64.dll',a.out/'soilfem_native_x64.dll')
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;w=None
try:
 w=app.Workbooks.Open(str(dst),UpdateLinks=0)
 for ws in w.Worksheets:
  if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
 for source in (root/'src/vba').glob('*.bas'):
  try:c=w.VBProject.VBComponents(source.stem)
  except Exception:c=w.VBProject.VBComponents.Add(1);c.Name=source.stem
  cm=c.CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(source.read_text(encoding='utf-8-sig'))
 cm=w.VBProject.VBComponents('FEMEngine').CodeModule;cm.AddFromString(Path(__file__).with_name('trial_tangent_restore_probe.bas').read_text(encoding='utf-8'))
 prefix="'"+w.Name+"'!";app.Run(prefix+'FEMWriteTextSetting','FLOW_POLICY','DAVIS');app.Run(prefix+'FEMWriteTextSetting','NATIVE_DLL',a.native);app.Run(prefix+'FEMWriteNumericSetting','NATIVE_STATE_COPY',1.);app.Run(prefix+'FEMWriteTextSetting','EXPORT_LOAD','')
 cc=app.VBE.CommandBars.FindControl(1,578)
 if cc.Enabled:cc.Execute()
 if app.VBE.CommandBars.FindControl(1,578).Enabled:
  pane=app.VBE.ActiveCodePane;sel=pane.GetSelection();print('COMPILE_LOCATION',pane.CodeModule.Parent.Name,sel,pane.CodeModule.Lines(sel[0],1),flush=True)
 assert not app.VBE.CommandBars.FindControl(1,578).Enabled
 v=list(app.Run(prefix+'TrialTangentRestoreAudit'));r={'native':a.native,'analysis':v[0],'plastic_gp':v[1],'elastic_yielded_gp':v[2],'dirty_after_restore':v[3],'plastic_vs_elastic_matrix':v[4],'restored_matrix_error':v[5],'native_state_copies':v[6]}
 assert v[0]=='PASS' and v[1]>0 and v[2]==0 and v[4]>1e-3,r
 assert (v[5]>1e-3 if a.expect_stale else v[5]<1e-8),r
 assert bool(v[3]) is (not a.expect_stale),r
 assert (v[6]>0 if a.native=='AUTO' else v[6]==0),r
 r['status']='STALE_REPRODUCED' if a.expect_stale else 'PASS';(a.out/'results.json').write_text(json.dumps(r,indent=2,allow_nan=False),encoding='utf-8');print(json.dumps(r,allow_nan=False),flush=True)
finally:
 if w is not None:w.Close(False)
 app.Quit()
