"""Exercise the production VBA progress controller in isolated native Excel."""
from pathlib import Path
import argparse,json
import win32com.client

root = Path(__file__).resolve().parents[1]
ap=argparse.ArgumentParser();ap.add_argument('--out',type=Path);args=ap.parse_args()
out = args.out or root / 'quality_20261113/progress_unit'
out.mkdir(parents=True, exist_ok=True)
app = win32com.client.DispatchEx('Excel.Application')
app.DisplayAlerts = False
wb = None
try:
    wb = app.Workbooks.Add()
    source = (root / 'src/vba/FEMIterationProgress.bas').read_text(encoding='utf-8-sig')
    mod = wb.VBProject.VBComponents.Add(1)
    mod.Name = 'FEMIterationProgress'
    mod.CodeModule.AddFromString(source)
    stub = wb.VBProject.VBComponents.Add(1)
    stub.Name = 'ProgressTest'
    stub.CodeModule.AddFromString('''Option Explicit
Private enabled As Double, span As Double, extensions As Long, reservations As Long
Public Function FEMReadSetting(ByVal key As String, ByVal fallback As Double) As Double
 If key = "NEWTON_PROGRESS_CONTINUE" Then FEMReadSetting=enabled Else FEMReadSetting=span
End Function
Public Function P2IsFinite(ByVal v As Double) As Boolean
 P2IsFinite=(Abs(v)<1E+300)
End Function
Public Sub FEMAppendRunLog(ByVal scope As String, ByVal eventName As String, ByVal detail As String)
 If eventName="EXTEND" Then extensions=extensions+1
 If eventName="RESERVE_RECOVERY" Then reservations=reservations+1
End Sub
Public Sub FEMInvalidateSettingCache()
End Sub
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise 1001,,message
End Sub
Public Function RunTests() As String
 Dim i As Long, base As Variant, v As Double, lim As Long, total As Long
 enabled=1:span=20
 For Each base In Array(100,1000)
  NewtonProgressReset CLng(base):v=1#
  For i=0 To 5000
   NewtonProgressRecord i,v:lim=NewtonProgressLimit(CLng(base),i)
   Check lim>i,"Improving residual stopped at " & CStr(i)
   v=v*0.999999
  Next i
 Next base
 total=extensions:Check total>400,"Continuation was not exercised"
 NewtonProgressReset 1000:v=1#
 For i=0 To 5000
  NewtonProgressRecord i,v:Check NewtonProgressLimit(1000,i)>i,"Tiny reduction was treated as stagnation"
  v=v*(1#-0.00000000001)
 Next i
 For i=0 To 2
  NewtonProgressReset 2:NewtonProgressRecord 0,1#:NewtonProgressRecord 1,0.9
  Select Case i
   Case 0: v=0.9
   Case 1: v=0.95
   Case 2: v=-1#
  End Select
  NewtonProgressRecord 2,v:Check NewtonProgressLimit(2,2)=2,"Invalid/no progress extended"
 Next i
 NewtonProgressReset 2:NewtonProgressRecord 0,1#:NewtonProgressRecord 1,0.9
 NewtonProgressRecord 2,0.8:Check NewtonProgressLimit(2,2)>2,"Small budget not extended"
 NewtonProgressReset 2:Check NewtonProgressLimit(2,2)=2,"History leaked across reset"
 NewtonProgressReset 2:NewtonProgressRecord 0,1#:NewtonProgressRecord 1,0.9
 Check NewtonProgressLimit(2,2)=2,"Stale accepted sample extended a rejected trial"
 NewtonProgressReset 2:NewtonProgressRecord 0,1#:NewtonProgressRecord 2,0.8
 NewtonProgressRecord 2,0.9:Check NewtonProgressLimit(2,2)>2,"Duplicate overwrote accepted sample"
 enabled=0:NewtonProgressReset 2:NewtonProgressRecord 0,1#:NewtonProgressRecord 2,0.8
 Check NewtonProgressLimit(2,2)=2,"Disabled setting ignored"
 Check Not NewtonProgressReduction(0.8,1#),"Disabled stagnation bypass"
 RunTests="PASS|extensions=" & CStr(total) & "|decreasing_iterations=15003|tiny_relative_decrease=1e-11|plateau,growth,invalid,reset,duplicate,disabled"
End Function
Public Function RecoveryTests() As String
 Dim lim As Long,i As Long,v As Double
 enabled=1:span=20:reservations=0
 NewtonProgressReset 100:NewtonProgressRecord 98,1.001:NewtonProgressRecord 99,1#
 Check NewtonProgressLimit(100,100)=100,"Strict outer stale guard changed"
 NewtonProgressReserveRecovery 100,100,21
 lim=NewtonProgressLimit(100,100):Check lim=141,"Fresh recovery work was cut off"
 NewtonProgressReserveRecovery 100,100,1000
 Check NewtonProgressLimit(100,100)=lim,"One sample reserved twice"
 NewtonProgressReserveRecovery 100,140,21
 Check NewtonProgressLimit(100,140)=lim,"Stale recovery sample extended"
 NewtonProgressRecord 140,0.999:NewtonProgressReserveRecovery 100,141,21
 Check NewtonProgressLimit(100,141)>lim,"New accepted progress did not renew recovery"
 For i=0 To 3
  NewtonProgressReset 100:NewtonProgressRecord 98,1#
  Select Case i
   Case 0:v=1#
   Case 1:v=1.1
   Case 2:v=-1#
   Case 3:v=0.9
  End Select
  NewtonProgressRecord 99,v
  If i=3 Then NewtonProgressReserveRecovery 100,101,21 Else NewtonProgressReserveRecovery 100,100,21
  Check NewtonProgressLimit(100,100)=100,"No-progress/invalid/stale recovery extended"
 Next i
 NewtonProgressReset 100:NewtonProgressRecord 80,0.5:NewtonProgressRecord 98,2#:NewtonProgressRecord 99,1#
 NewtonProgressReserveRecovery 100,100,21:Check NewtonProgressLimit(100,100)=100,"Growing window trend extended"
 NewtonProgressReset 100:NewtonProgressRecord 99,1#
 NewtonProgressReserveRecovery 100,100,21:Check NewtonProgressLimit(100,100)=100,"A first sample fabricated progress"
 enabled=0:NewtonProgressReset 100:NewtonProgressRecord 98,2#:NewtonProgressRecord 99,1#
 NewtonProgressReserveRecovery 100,100,21:Check NewtonProgressLimit(100,100)=100,"Disabled recovery continued"
 enabled=1:span=2:NewtonProgressReset 100:NewtonProgressRecord 95,2#:NewtonProgressRecord 99,1#
 NewtonProgressReserveRecovery 100,100,21:Check NewtonProgressLimit(100,100)=123,"Gapped accepted history in short window failed"
 NewtonProgressRecord 121,0.9:NewtonProgressReserveRecovery 100,122,21
 Check NewtonProgressLimit(100,122)=145,"Fresh gapped progress did not continue"
 span=20:NewtonProgressReset 100:NewtonProgressRecord 98,2#:NewtonProgressRecord 99,1#
 NewtonProgressReserveRecovery 100,100,0:NewtonProgressReserveRecovery 100,100,1001
 Check NewtonProgressLimit(100,100)=100,"Invalid work bound extended"
 Check reservations>=4,"Recovery reservation cases not exercised"
 RecoveryTests="recovery_reservations=" & CStr(reservations) & ";fresh,stale,duplicate,window,disabled,reset,bounds"
End Function
Public Function RunAllTests() As String
 On Error GoTo Failed
 RunAllTests=RunTests() & "|" & RecoveryTests():Exit Function
Failed:RunAllTests="ERROR|" & Err.Description
End Function
''')
    result = app.Run("'" + wb.Name + "'!RunAllTests")
    assert result.startswith('PASS|'), result
    (out / 'results.json').write_text(json.dumps({'status':'PASS','native_excel':True,'result':result},indent=2),encoding='utf-8')
    print(result, flush=True)
finally:
    if wb is not None: wb.Close(False)
    app.Quit()
