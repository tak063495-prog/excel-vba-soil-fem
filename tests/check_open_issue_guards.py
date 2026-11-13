"""Exercise existing load-lock and unsupported-mode protections for issue 1."""
from pathlib import Path
import argparse,json,shutil
import win32com.client
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args()
a.out.mkdir(parents=True,exist_ok=True);copy=a.out/'issue1_controls.xlsm';assert not copy.exists();shutil.copy2(a.book,copy)
app=win32com.client.DispatchEx('Excel.Application');app.EnableEvents=False;app.DisplayAlerts=False;wb=None
try:
 wb=app.Workbooks.Open(str(copy.resolve()),UpdateLinks=0)
 module=wb.VBProject.VBComponents.Item('FEMEngine').CodeModule
 module.AddFromString('''
Public Function I1HeldLoads() As Variant
 Dim t() As Double, values(0 To 4) As Double
 lastDof=0
 ReDim P3SelfWeightForce(0):ReDim P3LockedSelfWeight(0):ReDim P3AppliedForce(0):ReDim P3LockedApplied(0)
 ReDim NodeCond(0):ReDim P3BoundaryDisp(0):ReDim t(0)
 P3SelfWeightForce(0)=20
 P3BuildTargetForce 1,0,t:values(0)=t(0):P3CommitLoadLock True,False
 P3AppliedForce(0)=10
 P3BuildTargetForce 0,1,t:values(1)=t(0):P3CommitLoadLock False,True
 P3BuildTargetForce 0,0,t:values(2)=t(0)
 P3AppliedForce(0)=20:P3BuildTargetForce 0,1,t:values(3)=t(0):P3CommitLoadLock False,True
 P3AppliedForce(0)=0:P3BuildTargetForce 0,1,t:values(4)=t(0)
 I1HeldLoads=values
End Function
Public Function I1JointScope() As String
 Dim mat As Material_Data, msg As String
 mat.kn=1000:mat.ks=1000:mat.thickness=1:mat.cohesion=10:mat.psai=5
 If P3ValidateJointMaterial(mat,msg) Then I1JointScope="ERROR|unused psi accepted":Exit Function
 mat.psai=0
 If Not P3ValidateJointMaterial(mat,msg) Then I1JointScope="ERROR|" & msg:Exit Function
 I1JointScope="PASS|psi=5 rejected, declared stress-cap psi=0 accepted"
End Function
Public Function I1LegacyScope() As String
 On Error GoTo Rejected
 FEMWriteNumericSetting "CONSOL_ENABLE",1
 FEMValidateAnalysisScope
 I1LegacyScope="ERROR|legacy u-p accepted":Exit Function
Rejected:I1LegacyScope="REJECTED|" & Err.Description
End Function
''')
 prefix="'"+wb.Name+"'!";held=list(app.Run(prefix+'I1HeldLoads'));assert held==[20,30,30,40,20],held
 joint=app.Run(prefix+'I1JointScope');assert joint.startswith('PASS|'),joint
 legacy=app.Run(prefix+'I1LegacyScope');assert legacy.startswith('REJECTED|') and 'CONSOL_ENABLE' in legacy,legacy
 report={'status':'PASS','held_loads_gravity_apply_hold_add_unload':held,'joint_scope':joint,'legacy_up_scope':legacy,'scope':'existing supported-mode guards and real private load-lock routines; not a new JOINT plastic-slip law'}
 (a.out/'results.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(report,ensure_ascii=False),flush=True)
finally:
 if wb is not None:wb.Close(False)
 app.Quit()
