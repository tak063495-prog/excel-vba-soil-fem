"""Exercise real VBA checkpoint paths plus explicit state-transfer controls."""
from pathlib import Path
import argparse,json,shutil
import win32com.client
root=Path(__file__).resolve().parents[1]
ap=argparse.ArgumentParser();ap.add_argument('--runs-root',type=Path);ap.add_argument('--out',type=Path);a=ap.parse_args()
runs=a.runs_root or root/'evidence/targeted'
out=a.out or root/'evidence/checkpoint_repairs';out.mkdir(parents=True,exist_ok=True)
report=json.loads((out/'report.json').read_text(encoding='utf-8')) if (out/'report.json').exists() else {}

def run_case(name, action):
    source=runs/name/'fresh_results.xlsm'
    if not source.exists(): raise RuntimeError(f'Fixture not complete: {name}')
    copy=out/f'{name}_test.xlsm';shutil.copy2(source,copy)
    if not (out/'soilfem_native_x64.dll').exists():shutil.copy2(root/'delivery/soilfem_native_x64.dll',out/'soilfem_native_x64.dll')
    app=win32com.client.DispatchEx('Excel.Application');app.EnableEvents=False;app.DisplayAlerts=False
    wb=None
    try:
        wb=app.Workbooks.Open(str(copy),UpdateLinks=0)
        for ws in wb.Worksheets:
            if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
        module=wb.VBProject.VBComponents.Add(1);module.Name='CheckpointRepairTest'
        module.CodeModule.AddFromString('''Option Explicit
Public Function LoadCP(ByVal folder As String) As String
 On Error GoTo Failed
 SuppressUserMessages=True
 P3LoadTimeFrame folder
 LoadCP="PASS":Exit Function
Failed:LoadCP="ERROR|" & Err.Description
End Function
Public Function CPValues() As Variant
 CPValues=Array(FSS,P3FosPass,P3FosFail,P3FosMid,P3FosWidth,P3FosInterpretation,P3SrmTrialCount,FSTime,FSAccepted,FSStageElapsed,FSResumePending)
End Function
Public Function RunCP() As String
 On Error GoTo Failed
 SuppressUserMessages=True:P0_RunAnalysis
 RunCP=ResultStatus & "|" & AnalysisMessage:Exit Function
Failed:RunCP="ERROR|" & Err.Description
End Function
''')
        prefix="'"+wb.Name+"'!"
        call=lambda method,*args:app.Run(prefix+method,*args)
        action(app,wb,call,source.parent)
    finally:
        if wb is not None:wb.Close(False)
        app.Quit()
    (out/'report.json').write_text(json.dumps(report,indent=2,ensure_ascii=False,allow_nan=False),encoding='utf-8')

def checkpoints(folder):return list(folder.rglob('meta.txt'))
def kind(path,text):return f'KIND={text}' in path.read_text(encoding='cp932')

def dry(app,wb,call,folder):
    frame=next(p.parent for p in checkpoints(folder) if kind(p,'SRM'))
    assert call('LoadCP',str(frame))=='PASS'
    values=list(call('CPValues'))
    assert abs(values[0]-1.2)<1e-12 and abs(values[1]-1.2)<1e-12 and values[5]=='FIXED_FS_VERIFIED' and values[6]>0,values
    original=[list(row) for row in wb.Worksheets('ステージ結果').Range('A1:J5').Value2]
    legacy=out/'dry_legacy';shutil.copytree(frame,legacy,dirs_exist_ok=True)
    meta=legacy/'meta.txt';meta.write_text('\n'.join(s for s in meta.read_text(encoding='cp932').splitlines() if not s.startswith('FOS=')),encoding='cp932')
    assert call('LoadCP',str(legacy))=='PASS'
    old=list(call('CPValues'))
    assert abs(old[0]-1.2)<1e-12 and old[5]=='LEGACY_INTERVAL_UNAVAILABLE',old
    report['67']={'status':'PASS','restored':values,'legacy':old,'restored_stage_rows':original}

def missing(app,wb,call,folder):
    gravity=next(p.parent for p in checkpoints(folder) if kind(p,'GRAVITY'))
    assert call('LoadCP',str(gravity))=='PASS'
    first=min((p.parent for p in checkpoints(folder) if p.parent.name.startswith('C')),key=lambda p:p.name)
    assert call('LoadCP',str(first))=='PASS'
    bad=out/'missing_fixed_stress';shutil.copytree(first,bad,ignore=shutil.ignore_patterns('fixed_stress.csv'),dirs_exist_ok=True)
    result=call('LoadCP',str(bad));assert 'fixed_stress.csv' in result and result.startswith('ERROR|'),result
    meta=bad/'meta.txt';meta.write_text('\n'.join(s for s in meta.read_text(encoding='cp932').splitlines() if not s.startswith('HYDRAULICS=')),encoding='cp932')
    legacy=call('LoadCP',str(bad));assert 'fixed_stress.csv' in legacy and legacy.startswith('ERROR|'),legacy
    report['66']={'status':'PASS','pre_coupling_checkpoint':'PASS','normal_partial_checkpoint':'PASS','missing':result,'legacy_missing':legacy}

def flows(app,wb,call,folder):
    frame=next(p.parent for p in checkpoints(folder) if kind(p,'WAIT') and not p.parent.name.startswith('C'))
    wb.VBProject.VBComponents.Item('FEMFixedStress').CodeModule.AddFromString('''
Public Function I62FlowRoundtrip() As Variant
 Dim saved As Variant,n As Long,pre As Double,post As Double,maxErr As Double,expected() As Double,seeded As Boolean,phase As String,temp As Variant
 On Error GoTo Failed
 phase="read flow"
 For n=0 To NumberOfFreeNode-1:If Abs(mBoundaryFlow(n))>pre Then pre=Abs(mBoundaryFlow(n))
 Next n
 If pre=0# Then
  For n=0 To NumberOfFreeNode-1
   If mUsed(n) Then mBoundaryFlow(n)=0.125:seeded=True:Exit For
  Next n
 End If
 phase="flow copy":expected=mBoundaryFlow
 phase="water memory":temp=GWTransientState()
 phase="ledger memory":temp=FSLedgerMemory()
 phase="reference GP bounds":n=UBound(mReferenceGp,2)
 phase="pressure bounds":n=UBound(mP)
 phase="reference bounds":n=UBound(mReference)
 phase="used bounds":n=UBound(mUsed)
 phase="capture":saved=FSHydraulicMemory()
 For n=0 To NumberOfFreeNode-1:mBoundaryFlow(n)=999#:Next n
 phase="restore":FSRestoreHydraulicMemory saved
 For n=0 To NumberOfFreeNode-1
  If Abs(mBoundaryFlow(n)-expected(n))>maxErr Then maxErr=Abs(mBoundaryFlow(n)-expected(n))
  If Abs(mBoundaryFlow(n))>post Then post=Abs(mBoundaryFlow(n))
 Next n
 I62FlowRoundtrip=Array(pre,post,maxErr,seeded):Exit Function
Failed:I62FlowRoundtrip=Array("ERROR",phase,Err.Description)
End Function
''')
    assert call('LoadCP',str(frame))=='PASS'
    values=list(call('I62FlowRoundtrip'));assert values[0]!='ERROR',values
    assert values[1]>0 and values[2]==0,values
    report['62']={'status':'PASS','natural_accepted_max_flow':values[0],'restored_max_flow':values[1],'max_roundtrip_error':values[2],'synthetic_nonzero_control_used':bool(values[3])}

def split(app,wb,call,folder):
    frame=next(p.parent for p in checkpoints(folder) if kind(p,'WAIT') and not p.parent.name.startswith('C'))
    wb.VBProject.VBComponents.Item('FEMFixedStress').CodeModule.AddFromString('''
Public Function I65Split() As Variant
 Dim n As Long,a As Long,oldNode(0 To 3) As Long,result(0 To 3,0 To 2) As Double
 For n=0 To NumberOfFreeNode-1
  If mUsed(n) Then mP(n)=100#:mReference(n)=100#
 Next n
 mOldP=mP:FSUpdateField mP,True:GWCommit
 For a=0 To 3:oldNode(a)=mMap(a,1):Next a
 P3ElementActive(0)=False:P3ActiveSetGen=P3ActiveSetGen+1
 P6InvalidateActiveDependentCaches:GWTransientTopology:FSPrepare
 For a=0 To 3
  result(a,0)=oldNode(a):result(a,1)=mMap(a,1):result(a,2)=mP(mMap(a,1))
 Next a
 I65Split=result
End Function
''')
    assert call('LoadCP',str(frame))=='PASS'
    values=[list(row) for row in call('I65Split')]
    assert all(abs(r[2]-100)<1e-12 for r in values),values
    assert any(r[0]!=r[1] for r in values),values
    report['65']={'status':'PASS','test':'seeded accepted nonhydrostatic pressure 100 kPa; actual CONTINUOUS topology split and FSPrepare','old_representative_new_representative_pressure':values}

def steps(app,wb,call,folder):
    fs=wb.VBProject.VBComponents.Item('FEMFixedStress').CodeModule
    code=fs.Lines(1,fs.CountOfLines)
    code=code.replace('Option Explicit','Option Explicit\nPublic I63RejectFirst As Boolean\nPublic I63Takes As String',1)
    needle='If FSTryStep(take, target, locked, prescribed, limit, tolerance) Then'
    assert code.count(needle)==1
    code=code.replace(needle,'If I63Step(take, target, locked, prescribed, limit, tolerance) Then')
    code+='''
Private Function I63Step(ByVal dt As Double,ByRef target() As Double,ByRef locked() As Boolean,ByRef prescribed() As Double,ByVal limit As Long,ByVal tolerance As Double) As Boolean
 I63Takes=I63Takes & CStr(dt) & ","
 If I63RejectFirst Then I63RejectFirst=False:Exit Function
 I63Step=FSTryStep(dt,target,locked,prescribed,limit,tolerance)
End Function
Public Function I63Start() As String
 Dim before As String
 before=CStr(AnalysisRunning)
 I63RejectFirst=True:I63Takes="":SuppressUserMessages=True
 Call P0_RunAnalysis
 I63Start=ResultStatus & "|" & I63Takes
 If Len(ResultStatus)=0 Then I63Start="EMPTY|before=" & before & ";running=" & CStr(AnalysisRunning) & ";proc=" & FEMLastProc & ";message=" & AnalysisMessage
End Function
Public Function I63Trace() As String
 I63Trace=I63Takes
End Function
'''
    fs.DeleteLines(1,fs.CountOfLines);fs.AddFromString(code)
    call('FEMWriteTextSetting','EXPORT_LOAD','')
    result=call('I63Start');assert result.startswith('PASS|'),result
    takes=list(map(float,result.split('|',1)[1].strip(',').split(',')))
    assert takes[:2]==[.75,.5] and min(takes[1:])>=.5,takes
    first=min((p.parent for p in checkpoints(folder) if p.parent.name.startswith('C')),key=lambda p:p.name)
    legacy=out/'legacy_subminimum';shutil.copytree(first,legacy,dirs_exist_ok=True)
    f=legacy/'fixed_stress.csv';lines=f.read_text(encoding='cp932').splitlines()
    for i,line in enumerate(lines):
        if line.startswith('PHASE,'):
            parts=line.split(',');parts[4]='0.375';lines[i]=','.join(parts)
    f.write_text('\n'.join(lines)+'\n',encoding='cp932')
    call('FEMWriteTextSetting','EXPORT_LOAD',str(legacy))
    result2=call('RunCP');assert result2.startswith('PASS|'),result2
    state=list(call('CPValues'));assert abs(state[7]-3)<1e-12 and not state[10],state
    report['63']={'status':'PASS','controlled_first_trial_failure':True,'trial_steps_seconds':takes,'legacy_adaptive_dt_0_375_resume':result2,'resumed_state':state,'legacy_test_note':'constant, fully constrained field; only saved adaptive dt changed to exercise the historical subminimum compatibility path'}

for number,name,action in [('67','dry_fixed_fs',dry),('66','minimum_step',missing),('62','column_smoke',flows),('65','joint_pressure_split',split),('63','minimum_step',steps)]:
    if number not in report:run_case(name,action)
print('PASS: real VBA checkpoint paths for issues 62, 63, 65, 66, 67')
