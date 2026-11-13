"""Fault-inject one rejected Newton direction and verify unchanged equilibrium."""
from pathlib import Path
import argparse,json,shutil,win32com.client
root=Path(__file__).resolve().parents[1]
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args();a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=True)
records=[]
for mode in ('AUTO','OFF'):
 folder=a.out/mode;folder.mkdir();dest=folder/'associated_recovery_test.xlsm';shutil.copy2(a.book,dest);shutil.copy2(root/'native/bin/combined/soilfem_native_x64.dll',folder/'soilfem_native_x64.dll')
 app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;wb=None
 try:
  wb=app.Workbooks.Open(str(dest),UpdateLinks=0);app.Calculation=-4135
  for ws in wb.Worksheets:
   if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
  for source in (root/'src/vba').glob('*.bas'):
   try:component=wb.VBProject.VBComponents(source.stem)
   except Exception:component=wb.VBProject.VBComponents.Add(1);component.Name=source.stem
   cm=component.CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(source.read_text(encoding='utf-8-sig'))
  cm=wb.VBProject.VBComponents('FEMEngine').CodeModule;code=cm.Lines(1,cm.CountOfLines)
  marker='  P3ApplyLineSearch = False\r\n';assert code.count(marker)==1
  code=code.replace(marker,marker+'  If AssociatedRejectOnce Then\r\n    AssociatedRejectOnce=False:AssociatedRejectHits=AssociatedRejectHits+1\r\n    Exit Function\r\n  End If\r\n',1)
  cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(code)
  mod=wb.VBProject.VBComponents.Add(1);mod.Name='AssociatedRecoveryDriver';mod.CodeModule.AddFromString('''
Public AssociatedRejectOnce As Boolean,AssociatedRejectHits As Long
Public Function AssociatedRun() As Variant
 SuppressUserMessages=True:AssociatedRejectOnce=True
 P0_RunAnalysis
 AssociatedRun=Array(ResultStatus,AnalysisMessage,AssociatedRejectHits,TotalSelfWeight,P6ReactionSumX,P6ReactionSumY,RelativeResidualFree,FEMKernRegCalls,TDisp)
End Function
''')
  prefix="'"+wb.Name+"'!"
  wb.Worksheets('水条件').Range('B3').Value2='NONE'
  app.Run(prefix+'FEMWriteTextSetting','EXPORT_LOAD','');app.Run(prefix+'FEMWriteTextSetting','EXPORT_MODE','OFF');app.Run(prefix+'FEMWriteTextSetting','FLOW_POLICY','DAVIS');app.Run(prefix+'FEMWriteTextSetting','NATIVE_DLL',mode)
  r=app.Run(prefix+'AssociatedRun');assert r[0]=='PASS' and r[2]==1 and abs(r[3]-18.)<1e-10 and abs(r[4])<1e-8 and abs(r[5]-18.)<1e-8 and r[6]<=1e-5,r[:8]
  assert (r[7]>0)==(mode=='AUTO'),r[:8]
  records.append({'native':mode,'status':r[0],'forced_rejections':r[2],'weight':r[3],'reaction_x':r[4],'reaction_y':r[5],'relative_residual':r[6],'native_recovery_calls':r[7],'displacement':list(r[8])})
 finally:
  if wb is not None:wb.Close(False)
  app.Quit()
error=max(abs(x-y) for x,y in zip(records[0]['displacement'],records[1]['displacement']));assert error<1e-9,error
(a.out/'results.json').write_text(json.dumps({'status':'PASS','maximum_native_vba_displacement_difference':error,'cases':records},indent=2),encoding='utf-8');print('PASS associated-flow recovery, AUTO/OFF, equilibrium and native/VBA agreement',flush=True)
