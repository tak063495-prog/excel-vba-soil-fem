"""Real hydraulic transient followed by BIRTH, in both element traversal orders."""
from pathlib import Path
import argparse, copy, hashlib, json, shutil
import win32com.client

ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--expect-bug',action='store_true')
a=ap.parse_args();source=a.book.resolve();out=a.out.resolve();out.mkdir(parents=True,exist_ok=False)
root=Path(__file__).resolve().parents[1];digest=hashlib.sha256(source.read_bytes()).hexdigest()
base=json.loads((root/'tests/fixtures/fixed_stress/column_smoke.json').read_text(encoding='utf-8-sig'))
base['nodes']=base['nodes'][:13];base['elements']=base['elements'][:2];base['elements'][1][9]=2
for node in base['nodes']:node[3]=1;node[4]=1
base['materials']=[copy.deepcopy(base['materials'][0]),copy.deepcopy(base['materials'][0])]
base['materials'][1][0]=2;base['materials'][1][9]=0
base['stages']=[[1,'GRAVITY','',1,1,'initial'],[2,'WATER','W2',1,1,'hydraulic transient'],[3,'BIRTH',2,1,1,'activate upper soil'],[4,'WAIT','',1,1,'coupled step']]
base['loads']=[]
water=base['water'];water['boundaries']=[[1,1,'HEAD','*',None,'']]
# Fixed displacements isolate hydraulic history transfer without introducing
# mechanical unloading into the newborn soil. Pressure is solved, never seeded.
water['profiles']=[['W',0,4,''],['W',1,4,''],['W2',0,5,''],['W2',1,5,'']]
water['materials']=[[i,'EFFECTIVE',1e-4,'*',''] for i in [1,2]]
water['coupledMaterials']=[[i,1,1e-3,1e-4,1e-4,'',0.4] for i in [1,2]]
water['stagePlans']=[[2,'FIXED_STRESS','W2',1,1,'',200,1e-8,''],[3,'FIXED_STRESS','',0,0,'',200,1e-8,''],[4,'FIXED_STRESS','',1,1,'',200,1e-8,'']]
water['timeOptions']=[];water['drainSets']=[];water['timeCurves']=[]
instrumentation='''
Private I73Observe As Boolean
Private I73Map() As Long,I73Pressure() As Double,I73Reference() As Double,I73GpReference() As Double
Public I73Metrics As Variant
Public Sub I73BeforePrepare(ByVal checkpoint As Boolean)
 I73Observe=False
 If Not mReady Or checkpoint Or P3RunLogKind<>"BIRTH" Then Exit Sub
 I73Map=mMap:I73Pressure=mP:I73Reference=mReference:I73GpReference=mReferenceGp:I73Observe=True
End Sub
Public Sub I73AfterPrepare()
 Dim e As Long,a As Long,j As Long,b As Long,n As Long,old As Long,g As Long,found As Boolean
 Dim expected As Double,expectedReference As Double,dp As Double,dr As Double,dnew As Double,dgp As Double,dnewgp As Double,excess As Double,sharedCount As Long,newOnly As Long
 If Not I73Observe Then Exit Sub
 For e=0 To NumberOfElement-1
  If P3IsElementActive(e) And Material(Elem(e).MatNo).kind="SOIL" Then
   For a=0 To 3
    If a<3 Or Not FEMElementIsT6(e) Then
     n=mMap(a,e):found=False
     For j=0 To NumberOfElement-1
      If P3IsElementActive(j) Then
       For b=0 To 3
        If I73Map(b,j)>=0 And mMap(b,j)=n Then
         old=I73Map(b,j):expected=I73Pressure(old):expectedReference=I73Reference(old):found=True
        End If
       Next b
      End If
     Next j
     If found Then
      dp=Application.Max(dp,Abs(mP(n)-expected)):dr=Application.Max(dr,Abs(mReference(n)-expectedReference))
      If I73Map(a,e)<0 Then
       sharedCount=sharedCount+1:excess=Application.Max(excess,Abs(expected-GWInitialNodePressure(e,a)))
      End If
     Else
      newOnly=newOnly+1:dnew=Application.Max(dnew,Abs(mP(n)-GWInitialNodePressure(e,a)),Abs(mReference(n)-mP(n)))
     End If
    End If
   Next a
   For g=0 To 3
    If I73Map(0,e)>=0 Then
     dgp=Application.Max(dgp,Abs(mReferenceGp(g,e)-I73GpReference(g,e)))
    Else
     dnewgp=Application.Max(dnewgp,Abs(mReferenceGp(g,e)-mGp(g,e)))
    End If
   Next g
  End If
 Next e
 I73Metrics=Array(excess,dp,dr,dnew,dgp,dnewgp,sharedCount,newOnly)
 I73Observe=False
End Sub
Public Function I73GetMetrics() As Variant
 I73GetMetrics=I73Metrics
End Function
Public Function I73ConflictControl(ByVal referenceOnly As Boolean) As String
 Dim e As Long,a As Long,j As Long,b As Long,n As Long,spare As Long
 On Error GoTo Rejected
 For spare=0 To NumberOfFreeNode-1
  If Not mUsed(spare) Then Exit For
 Next spare
 For e=0 To NumberOfElement-1
  For a=0 To 3
   n=mMap(a,e)
   If n>=0 Then
    For j=e+1 To NumberOfElement-1
     For b=0 To 3
      If mMap(b,j)=n Then
       ' Explicit conflict control: simulate two previously separate histories
       ' merging into one representative. This is not a physical transient.
       mUsed(spare)=True:mP(spare)=mP(n):mReference(spare)=mReference(n)
       If referenceOnly Then mReference(spare)=mReference(spare)+1# Else mP(spare)=mP(spare)+1#
       mMap(b,j)=spare
       FSPrepare
       I73ConflictControl="UNEXPECTED_ACCEPT":Exit Function
      End If
     Next b
    Next j
   End If
  Next a
 Next e
 I73ConflictControl="NO_SHARED_NODE":Exit Function
Rejected:
 I73ConflictControl=Err.Description
End Function
'''

report=[]
for reverse in [False,True]:
    d=copy.deepcopy(base);d['name']='birth_newborn_first' if reverse else 'birth_survivor_first'
    if reverse:
        d['elements'].reverse()
        for i,row in enumerate(d['elements'],1):row[0]=i
        d['water']['boundaries'][0][0]=2
    folder=out/d['name'];folder.mkdir();(folder/'fixture.json').write_text(json.dumps(d,ensure_ascii=False,indent=2),encoding='utf8')
    target=folder/'fixture.xlsm';shutil.copyfile(source,target)
    shutil.copyfile(source.parent/'soilfem_band_fast_x64.dll',folder/'soilfem_band_fast_x64.dll')
    app=win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Excel.Application'))
    app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.AutomationSecurity=1
    try:
        w=app.Workbooks.Open(str(target),UpdateLinks=0)
        def run(n,*args):return app.Run("'"+w.Name+"'!"+n,*args)
        def rows(sn,r,c,data):
            if data:
                nc=max(map(len,data));w.Worksheets(sn).Cells(r,c).GetResize(len(data),nc).Value2=tuple(tuple(row+[None]*(nc-len(row))) for row in data)
        for ws in w.Worksheets:ws.Unprotect()
        run('FEMEnsureStageSheets')
        for sn,key in [('節点データ','nodes'),('要素データ','elements'),('材料データ','materials'),('ステージ','stages'),('載荷','loads'),('接合','jointPairs')]:
            w.Worksheets(sn).Range('A2:N60001').ClearContents();rows(sn,2,1,d.get(key))
        run('FEMWaterInstallUi');ws=w.Worksheets('水条件')
        for addr in ['A12:D211','F12:J211','L12:T211','V12:AB211','AC12:AL211','A216:F415','V216:AB415','V420:X2000']:ws.Range(addr).ClearContents()
        for addr,key in [('B3','mode'),('B4','gamma'),('B5','defaultProfile'),('B6','auto')]:ws.Range(addr).Value2=d['water'][key]
        for key,r,c in [('profiles',12,1),('materials',12,6),('boundaries',216,1),('stagePlans',12,12),('coupledMaterials',12,22),('timeOptions',12,29),('drainSets',216,22),('timeCurves',420,22)]:rows('水条件',r,c,d['water'].get(key))
        for k,v in dict(MESH_GEN_COMPARE=1,EXPORT_MODE='STAGE',EXPORT_LOAD='',FLOW_POLICY='INCONSISTENT',NATIVE_DLL='AUTO',OUTPUT_STAGE_MODE='ALL',RCM_POLICY='OFF').items():run('FEMWriteTextSetting' if isinstance(v,str) else 'FEMWriteNumericSetting',k,v)
        cm=w.VBProject.VBComponents('FEMFixedStress').CodeModule;code=cm.Lines(1,cm.CountOfLines)
        # Module-level declarations must precede all procedures.
        declarations,procedures=instrumentation.split('Public Sub I73BeforePrepare',1)
        code=code.replace('Option Explicit','Option Explicit\n'+declarations,1)
        code=code.replace('  initiallyReady = mReady','  I73BeforePrepare fromCheckpoint\n  initiallyReady = mReady',1)
        start=code.index('Public Sub FSPrepare(');end=code.index('End Sub',start)
        assert '  GWCommit' in code[start:end]
        code=code[:start]+code[start:end].replace('  GWCommit','  GWCommit\n  I73AfterPrepare')+code[end:]+'\nPublic Sub I73BeforePrepare'+procedures
        cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(code)
        driver=w.VBProject.VBComponents.Add(1);driver.Name='Issue73Driver';driver.CodeModule.AddFromString('''Option Explicit
Public Function I73Run() As Variant
 SuppressUserMessages=True:FEMInvalidateSettingCache:GWPreviewCompute
 ThisWorkbook.Worksheets("水条件").Range("B7").Value2=GWInputSignature()
 P0_RunAnalysis
 I73Run=Array(AnalysisOK,ResultStatus,AnalysisMessage,FSAccepted,FSRunningStage,FSTime)
End Function
''')
        ctl=app.VBE.CommandBars.FindControl(1,578)
        if ctl.Enabled:ctl.Execute()
        if app.VBE.CommandBars.FindControl(1,578).Enabled:
            pane=app.VBE.ActiveCodePane
            print('COMPILE_DIAG',pane.CodeModule.Name,pane.GetSelection(),flush=True)
            w.SaveCopyAs(str(folder/'compile_debug.xlsm'))
            raise AssertionError('Instrumented fixture did not compile')
        result=list(run('I73Run'));metrics=run('I73GetMetrics')
        rec=dict(name=d['name'],result=result,metrics=metrics,passed=bool(result[0]));report.append(rec)
        (folder/'result.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2),encoding='utf8')
        print('ISSUE73_CASE',json.dumps(rec,ensure_ascii=False),flush=True)
        if a.expect_bug:
            assert not result[0] and '共有圧力節点の初期水圧が矛盾' in result[2],result
        else:
            assert result[0],result
            assert metrics[0]>0.1 and metrics[6]>=2 and metrics[7]>=2,metrics
            assert max(metrics[1:6])<1e-10,metrics
            assert result[4]==4 and result[5]==2,result
            w.SaveCopyAs(str(folder/'fresh_results.xlsm'))
            conflict=run('I73ConflictControl',reverse)
            assert '合流する存続要素の受理水圧または参照水圧が矛盾' in conflict,conflict
            rec['explicit_conflict_control']={'reference_only':reverse,'rejected':True,'message':conflict}
            (folder/'result.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2),encoding='utf8')
    finally:
        for w in list(app.Workbooks):w.Close(False)
        app.Quit()
assert hashlib.sha256(source.read_bytes()).hexdigest()==digest
if not a.expect_bug:assert abs(report[0]['metrics'][0]-report[1]['metrics'][0])<1e-10
(out/'report.json').write_text(json.dumps(dict(passed=True,expected_old_bug=a.expect_bug,source_sha256=digest,source_unchanged=True,cases=report),ensure_ascii=False,indent=2),encoding='utf8')
print('ISSUE73_PASS',flush=True)
