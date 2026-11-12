"""Run native Excel quality regression or regenerate the attached geometry."""
from pathlib import Path
import argparse,json,shutil,time,os
import win32com.client,win32process
root=Path(__file__).resolve().parents[1]
ap=argparse.ArgumentParser();ap.add_argument('mode',choices=['regression','generate','analyze','preview'])
ap.add_argument('--book',type=Path,required=True)
a=ap.parse_args();path=a.book.resolve();folder=path.parent
if a.mode=='regression':
    path=folder/'quality_regression.xlsm';assert not path.exists();shutil.copy2(a.book,path)
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;app.Visible=True
pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
(folder/(a.mode+'_process.json')).write_text(json.dumps({'pid':pid,'book':str(path),'python_pid':os.getpid()}),encoding='utf-8')
wb=None
try:
    wb=app.Workbooks.Open(str(path),UpdateLinks=0);app.Calculation=-4135
    for ws in wb.Worksheets:
        if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
    mod=wb.VBProject.VBComponents.Add(1);mod.Name='QualityTest'
    mod.CodeModule.AddFromString((root/'tests/mesh_quality_regression.bas').read_text(encoding='utf-8-sig'))
    prefix="'"+wb.Name+"'!"
    cc=app.VBE.CommandBars.FindControl(1,578)
    if cc.Enabled:cc.Execute()
    assert not app.VBE.CommandBars.FindControl(1,578).Enabled,'compile'
    started=time.monotonic()
    if a.mode=='regression':
        result=app.Run(prefix+'QualityRegression');assert result.startswith('PASS'),result
        print(result,flush=True)
        negative=app.Run(prefix+'QualityValidityAndDiagonal');print(negative,flush=True);assert negative.startswith('PASS'),negative
        flip=app.Run(prefix+'QualityEdgeFlip');print(flip,flush=True);assert flip.startswith('PASS'),flip
        # Solve the existing tapered Q8/T6 weight/reaction patch with the new code.
        j=json.loads(((root/'tests/fixtures/mixed_t6/q8_gravity.json')).read_text(encoding='utf-8-sig'))
        for name,key in [('節点データ','nodes'),('要素データ','elements'),('材料データ','materials'),('ステージ','stages')]:
            ws=wb.Worksheets(name);ws.Range('A2:N60010').ClearContents()
            data=j[key];cols=max(map(len,data));ws.Range(ws.Cells(2,1),ws.Cells(len(data)+1,cols)).Value2=tuple(tuple(row+[None]*(cols-len(row))) for row in data)
        for name in ('載荷','接合'):wb.Worksheets(name).Range('A2:N60010').ClearContents()
        wb.Worksheets('水条件').Range('B3').Value2='NONE';wb.Worksheets('水条件').Range('A216:F415').ClearContents()
        w=wb.Worksheets('節点データ');right=j['elements'][1]
        w.Cells(right[3]+1,2).Value2=3.;w.Cells(right[3]+1,3).Value2=.05;w.Cells(right[4]+1,2).Value2=2.
        for row in j['elements']:
            for k in range(4):
                a1=row[k+1]+1;b1=row[(k+1)%4+1]+1;m=row[k+5]+1
                for c in (2,3):w.Cells(m,c).Value2=(w.Cells(a1,c).Value2+w.Cells(b1,c).Value2)/2
        for key,value in [('MESH_T6_FALLBACK',1.),('MESH_QUALITY_MIN_ANGLE_DEG',10.),('MESH_QUALITY_MAX_ANGLE_DEG',170.)]:app.Run(prefix+'FEMWriteNumericSetting',key,value)
        app.Run(prefix+'FEMWriteTextSetting','EXPORT_MODE','OFF');app.Run(prefix+'FEMWriteTextSetting','EXPORT_LOAD','')
        helper=wb.VBProject.VBComponents.Add(1);helper.Name='MixedT6Driver'
        helper.CodeModule.AddFromString((root/'tests/mixed_t6_mesh_driver.bas').read_text(encoding='utf-8-sig'))
        assert app.Run(prefix+'M6ApplyFallback').startswith('PASS')
        assert app.Run(prefix+'M6Validate',1,2,14).startswith('PASS')
        print('PREFLIGHT',app.Run(prefix+'QualityPreflight'),flush=True)
        solve=app.Run(prefix+'M6Run');print(solve,flush=True);assert solve.startswith('PASS'),solve
        local=app.Run(prefix+'QualityRefinementFixture');print(local,flush=True);assert local.startswith('PASS'),local
        report={'status':'PASS','pure_quality':result,'validity_and_diagonal':negative,'edge_flip':flip,'mixed_patch':solve,'local_refinement':local}
        wb.VBProject.VBComponents.Remove(helper)
    elif a.mode=='generate':
        app.Run(prefix+'MeshQualityEnsureSettings')
        for key,value in [('MESH_T6_FALLBACK',1.),('MESH_GEN_COMPARE',1.),('MESH_REPAIR_TERRAIN_COMPARE',0.),('MESH_QUALITY_REFINE',1.),('SOLVER_MEMORY_LIMIT_MB',8192.)]:
            app.Run(prefix+'FEMWriteNumericSetting',key,value)
        # Sync canonical settings to the visible inputs before generation.
        app.Run(prefix+'FEMUiRefreshInputs')
        print('GENERATE_START',time.strftime('%H:%M:%S'),flush=True)
        result=app.Run(prefix+'QualityGenerate');print(result,flush=True);assert result.startswith('PASS'),result
        settings=wb.Worksheets('設定');last=settings.Cells(settings.Rows.Count,5).End(-4162).Row
        values={row[4]:row[2] for row in settings.Range('A5:E'+str(last)).Value2 if row[4]}
        assert not str(values.get('MESH_GEN_LAST_METHOD','')).startswith('COMPARE_'),values.get('MESH_GEN_LAST_COMPARE')
        report={'status':'PASS','result':result,'settings':values}
    elif a.mode=='analyze':
        app.Run(prefix+'FEMWriteTextSetting','DEBUG_MODE','ITER');app.Run(prefix+'FEMWriteNumericSetting','DEBUG_FLUSH',1.)
        print('ANALYZE_START',time.strftime('%H:%M:%S'),flush=True)
        result=app.Run(prefix+'QualityAnalyze');print(result,flush=True)
        report={'status':result}
    else:
        app.Run(prefix+'FEM24Model');print('PREVIEW_READY',flush=True);input();report={'status':'PASS'}
    report['seconds']=time.monotonic()-started
    (folder/(a.mode+'_result.json')).write_text(json.dumps(report,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
    wb.VBProject.VBComponents.Remove(mod)
    app.Calculation=-4105;wb.Save()
finally:
    if wb is not None:wb.Close(False)
    app.Quit()
