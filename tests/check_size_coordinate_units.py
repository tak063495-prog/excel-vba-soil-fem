"""Real VBA quality check: equal physical areas across Q8/T6 reference maps."""
from pathlib import Path
import argparse,json,shutil
import win32com.client
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--legacy',action='store_true');a=ap.parse_args()
root=Path(__file__).resolve().parents[1];a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=False)
src=a.book.resolve();book=a.out/'size_units.xlsm';shutil.copy2(src,book)
app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;w=None
records=[]
controls=[]
try:
    w=app.Workbooks.Open(str(book),UpdateLinks=0);app.Calculation=-4135
    for s in w.Worksheets:
        if s.ProtectContents:s.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
    for name in ('Mesh','MeshQuality','FEMIo'):
        text=(root/'src/vba'/ (name+'.bas')).read_text(encoding='utf-8-sig')
        if a.legacy and name=='Mesh':
            old='If area < P5ModelScale ^ 2 * P5ReadSetting("MESH_QUALITY_MIN_AREA_RATIO", 0.000001) Or minJacobian < P5ModelScale ^ 2 * P5ReadSetting("MESH_QUALITY_MIN_JACOBIAN_RATIO", 0.000001) Then P5SizeWarnings = P5SizeWarnings + 1'
            for triangle in ('True','False'):
                marker='If MeshElementHasSizeWarning('+triangle+', area, minJacobian, P5ModelScale) Then P5SizeWarnings = P5SizeWarnings + 1'
                assert text.count(marker)==1; text=text.replace(marker,old,1)
        cm=w.VBProject.VBComponents(name).CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(text)
    for name,cols in (('節点データ','I'),('要素データ','K'),('壁面支持','F')):
        s=w.Worksheets(name);end=s.Cells(s.Rows.Count,1).End(-4162).Row
        if end>1:s.Range('A2:'+cols+str(end)).ClearContents()
    cm=w.VBProject.VBComponents('Mesh').CodeModule
    cm.InsertLines(cm.CountOfLines+1,'''Public Function SizeUnitCheck() As Variant
SuppressUserMessages=True
FEMInvalidateSettingCache
P5EvaluateCurrentSheet "equal physical area size check"
SizeUnitCheck=Array(P5ModelScale,P5SizeWarnings,P5Q8Count,P5T6Count,P5Q8ShapeWarnings,P5T6ShapeWarnings,P5MinJacobian)
End Function''')
    cc=app.VBE.CommandBars.FindControl(1,578)
    if cc.Enabled:cc.Execute()
    assert not app.VBE.CommandBars.FindControl(1,578).Enabled
    prefix="'"+w.Name+"'!"
    for key,value in (('MESH_QUALITY_MIN_AREA_RATIO',1e-12),('MESH_QUALITY_MIN_ANGLE_DEG',45.),('MESH_QUALITY_MAX_ANGLE_DEG',135.),('MESH_QUALITY_MAX_ASPECT_RATIO',5.),('MESH_T6_MIN_ANGLE_DEG',20.),('MESH_T6_MAX_ANGLE_DEG',140.),('MESH_T6_MAX_ASPECT_RATIO',5.)):
        app.Run(prefix+'FEMWriteNumericSetting',key,value)
    base=[(0,0),(1,0),(1,1),(0,1),(.5,0),(1,.5),(.5,1),(0,.5),(3,0),(2,0),(2,.5)]
    w.Worksheets('要素データ').Range('A2:K3').Value2=((1,1,2,3,4,5,6,7,8,1,'Q8'),(2,2,9,3,3,10,11,3,6,1,'T6'))
    for scale,shift in ((1.,0.),(10.,0.),(1.,1234.)):
        nodes=[(i+1,shift+scale*x,shift+scale*y,0,0,0,0,0,0) for i,(x,y) in enumerate(base)]
        w.Worksheets('節点データ').Range('A2:I12').Value2=tuple(nodes)
        for threshold,expected in ((.1,0),(.2,1 if a.legacy else 0),(.3,2)):
            app.Run(prefix+'FEMWriteNumericSetting','MESH_QUALITY_MIN_JACOBIAN_RATIO',threshold/9.)
            values=list(app.Run(prefix+'SizeUnitCheck'))
            assert abs(values[0]-3*scale)<1e-7 and values[1]==expected and values[2:6]==[1.,1.,0.,0.],values
            # Raw integration-coordinate Jacobian remains unchanged and available.
            assert abs(values[6]-(1-1/3**.5)/4*scale**2)<1e-8,values
            records.append({'scale':scale,'translation':shift,'physical_square_area':scale**2,'physical_triangle_area':scale**2,'threshold_equivalent_jacobian':threshold*scale**2,'size_warnings':values[1],'raw_min_jacobian':values[6]})
    # Restore the production thresholds before independent geometry controls.
    for key in ('MESH_QUALITY_MIN_AREA_RATIO','MESH_QUALITY_MIN_JACOBIAN_RATIO'):
        app.Run(prefix+'FEMWriteNumericSetting',key,.000001)
    driver=w.VBProject.VBComponents.Add(1);driver.Name='SizeValidityControls'
    driver.CodeModule.AddFromString((root/'tests/mesh_quality_regression.bas').read_text(encoding='utf-8-sig'))
    cc=app.VBE.CommandBars.FindControl(1,578)
    if cc.Enabled:cc.Execute()
    assert not app.VBE.CommandBars.FindControl(1,578).Enabled
    for name in ('QualityRegression','QualityValidityAndDiagonal','QualityEdgeFlip'):
        result=app.Run(prefix+name);assert result.startswith('PASS'),(name,result)
        controls.append({'test':name,'result':result})
    w.Save()
finally:
    if w is not None:w.Close(False)
    app.Quit()
report={'status':'PASS','legacy_raw_coordinate_comparison':a.legacy,'checks':records,'validity_and_shape_controls':controls}
(a.out/'results.json').write_text(json.dumps(report,indent=2),encoding='utf-8');print(json.dumps(report),flush=True)
