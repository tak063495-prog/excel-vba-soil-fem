"""Exercise algebra-only native recovery while physical-water material stays VBA."""
from pathlib import Path
import argparse,json,shutil,win32com.client
root=Path(__file__).resolve().parents[1]
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args();a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=True)
dest=a.out/'water_recovery_probe.xlsm';assert not dest.exists();shutil.copy2(a.book,dest);shutil.copy2(root/'native/bin/combined/soilfem_native_x64.dll',a.out/'soilfem_native_x64.dll')
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;wb=None
try:
 wb=app.Workbooks.Open(str(dest),UpdateLinks=0);app.Calculation=-4135
 for ws in wb.Worksheets:
  if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
 for name in ('FEMNativeMat','FEMNativeKern','FEMSolver'):
  cm=wb.VBProject.VBComponents(name).CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString((root/'src/vba'/f'{name}.bas').read_text(encoding='utf-8-sig'))
 cm=wb.VBProject.VBComponents('FEMSolver').CodeModule;cm.AddFromString('''
Public Function ProbeWaterRecovery() As Variant
 Dim i As Long, calls As Long, ok As Boolean, copy() As Double
 P6CSRReady=False
 For i=0 To lastDof
  force(i)=0#:Disp(i)=0#
  If NodeCond(i)=0 Then force(i)=Sin(CDbl(i)+1#)
 Next i
 calls=FEMKernRegCalls
 ok=P6SolveResidualRecovery(0.001)
 If Not P6CSRReady Then Err.Raise 1001,,"Recovery failed to build CSR lazily"
 copy=Disp
 ProbeWaterRecovery=Array(ok,FEMKernRegCalls-calls,GWActive,FEMMatStatusText(),FEMKernStatusText(),copy)
End Function
''')
 mod=wb.VBProject.VBComponents.Add(1);mod.Name='RecoveryKernelDriver';mod.CodeModule.AddFromString('''
Public Function RecoveryKernelRun() As String
 SuppressUserMessages=True
 P0_RunAnalysis
 RecoveryKernelRun=ResultStatus & "|" & AnalysisMessage
End Function
''')
 prefix="'"+wb.Name+"'!"
 app.Run(prefix+'FEMWriteTextSetting','NATIVE_PRECISION','AVX2');app.Run(prefix+'FEMWriteTextSetting','EXPORT_LOAD','')
 assert app.Run(prefix+'RecoveryKernelRun')=='PASS|'
 result=app.Run(prefix+'ProbeWaterRecovery');assert result[0] and result[1]==1 and result[2] and result[3]=='VBA(PHYSICAL_WATER_VBA_MATERIAL)',result[:5]
 record={'status':'PASS','native_recovery_calls':result[1],'physical_water_active':result[2],'material_path':result[3],'kernel_path':result[4],'direction':list(result[5])}
 (a.out/'results.json').write_text(json.dumps(record,indent=2),encoding='utf-8');print(json.dumps({k:v for k,v in record.items() if k!='direction'}),flush=True)
finally:
 if wb is not None:wb.Close(False)
 app.Quit()
