"""Exercise real mesh transactions on disposable copies; never edit the input book."""
from pathlib import Path
import argparse, hashlib, json, shutil
import win32com.client

ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--expect-bug',action='store_true')
a=ap.parse_args();source=a.book.resolve();out=a.out.resolve();out.mkdir(parents=True,exist_ok=False)
digest=hashlib.sha256(source.read_bytes()).hexdigest();target=out/'mesh_transaction.xlsm';shutil.copyfile(source,target)
app=win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Excel.Application'))
app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.AutomationSecurity=1
record={}
try:
    w=app.Workbooks.Open(str(target),UpdateLinks=0)
    for ws in w.Worksheets:ws.Unprotect()
    refs=w.Worksheets.Add();refs.Name='Issue75Refs'
    ns=w.Worksheets('節点データ');es=w.Worksheets('要素データ');shape=w.Worksheets('形状入力');cache=w.Worksheets('SRM準備メッシュ')
    refs.Range('A1').Value2=ns.Range('B2').Value2;refs.Range('A2').Value2=es.Range('J2').Value2
    ns.Range('B2').Formula2="='Issue75Refs'!A1";es.Range('J2').Formula2="='Issue75Refs'!A2"
    literals=['=1+2','+123','00123','@ABC',"'text",'ordinary text']
    for i,v in enumerate(literals,2):ns.Cells(i,14).Formula2="'"+v
    ns.Range('N8').Formula2='=NA()';app.Calculate()
    n=int(ns.Cells(ns.Rows.Count,1).End(-4162).Row);e=int(es.Cells(es.Rows.Count,1).End(-4162).Row)
    areas=[ns.Range(f'A1:N{n}'),es.Range(f'A1:K{e}')]
    original=[(r.Formula2,r.Value2) for r in areas]
    mod=w.VBProject.VBComponents.Add(1);mod.Name='Issue75Driver';mod.CodeModule.AddFromString('''Option Explicit
Public Function Issue75Create() As Variant
 SuppressUserMessages=True:FEMInvalidateSettingCache
 ボタン19_Click
 Issue75Create=Array(FEMShapeGenerationOK,FEMUiMeshPreviewError)
End Function
''')
    def run():return list(app.Run("'"+w.Name+"'!Issue75Create"))
    def unchanged():
        app.Calculate()
        assert [(r.Formula2,r.Value2) for r in areas]==original,'Input contents or formulas changed'
        assert ns.Range('B2').HasFormula and es.Range('J2').HasFormula
        assert all(not ns.Cells(i,14).HasFormula and ns.Cells(i,14).Value2==v for i,v in enumerate(literals,2))
        before=refs.Range('A1').Value2;refs.Range('A1').Value2=before+0.125;app.Calculate()
        assert ns.Range('B2').Value2==before+0.125,'Coordinate reference no longer recalculates'
        refs.Range('A1').Value2=before
        before=refs.Range('A2').Value2;refs.Range('A2').Value2=before+1;app.Calculate()
        assert es.Range('J2').Value2==before+1,'Material reference no longer recalculates'
        refs.Range('A2').Value2=before;app.Calculate()
    spacing=shape.Range('J8').Value2;shape.Range('J8').Value2=float(spacing)+1
    result=run();assert not result[0] and '準備済みメッシュと異なります' in result[1],result
    if a.expect_bug:
        assert not ns.Range('B2').HasFormula and not es.Range('J2').HasFormula,'Old defect did not reproduce'
        assert ns.Range('B2').Value2==refs.Range('A1').Value2 and es.Range('J2').Value2==refs.Range('A2').Value2
        record.update(passed=True,expected_old_bug=True,formulas_lost=True,source_sha256=digest)
        (out/'result.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf8')
        print('ISSUE75_OLD_BUG_REPRODUCED',flush=True)
        raise SystemExit(0)
    unchanged();record['signature_mismatch_before_write']={'passed':True,'result':result}
    shape.Range('J8').Value2=spacing
    area=cache.Range('Z6').Value2;cache.Range('Z6').Value2=float(area)+10
    result=run();assert not result[0] and '材料別面積' in result[1],result
    unchanged();record['validation_failure_after_write']={'passed':True,'result':result,'formulas_values_and_text_constants_preserved':True,'references_recalculate':True}
    cache.Range('Z6').Value2=area
    result=run();assert result[0] and not result[1],result
    assert int(ns.Cells(ns.Rows.Count,1).End(-4162).Row)==n
    assert int(es.Cells(es.Rows.Count,1).End(-4162).Row)==e
    assert not ns.Range('B2').HasFormula and not es.Range('J2').HasFormula
    record['normal_prepared_restore']={'passed':True,'nodes':n-1,'elements':e-1,'result':result}
    record.update(passed=True,source_sha256=digest,source_unchanged=True)
    (out/'result.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf8')
    print('ISSUE75_PASS',json.dumps(record,ensure_ascii=False),flush=True)
finally:
    for w in list(app.Workbooks):w.Close(False)
    app.Quit()
assert hashlib.sha256(source.read_bytes()).hexdigest()==digest
