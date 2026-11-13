"""Exercise conservative patch scope without changing the supplied workbook."""
from pathlib import Path
import argparse,json,shutil
import win32com.client
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args()
a.out.mkdir(parents=True,exist_ok=True);copy=a.out/'scope.xlsm';assert not copy.exists();shutil.copy2(a.book,copy)
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;wb=None
try:
    wb=app.Workbooks.Open(str(copy.resolve()),UpdateLinks=0)
    for ws in wb.Worksheets:
        if ws.ProtectContents:ws.Protect(Password='',DrawingObjects=False,Contents=True,Scenarios=True,UserInterfaceOnly=True)
    cm=wb.VBProject.VBComponents.Item('MeshT6Patch').CodeModule
    cm.AddFromString('''
Public Function PatchScopeTest() As String
 PatchScopeTest=CStr(PatchScopeAllowed())
End Function
''')
    call=lambda: app.Run("'"+wb.Name+"'!PatchScopeTest")
    load=wb.Worksheets('載荷');joint=wb.Worksheets('接合');water=wb.Worksheets('水条件')
    load.Range('A2:N100').ClearContents();joint.Range('A2:N100').ClearContents();water.Range('B3').Value2='NONE'
    try:wb.Worksheets('要素セット').Range('A2:N100').ClearContents()
    except Exception:pass
    checks={}
    assert call()=='True';checks['dry_no_external_IDs']='PASS'
    joint.Range('A2').Value2=1;assert call()=='False';checks['joint_blocks_moves_and_compaction']='PASS';joint.Range('A2').ClearContents()
    water.Range('B3').Value2='PHREATIC';assert call()=='False';checks['hydraulic_model_skipped']='PASS';water.Range('B3').Value2='NONE'
    for selector,mode,expected in [('TOP','DISP','True'),('LEFT','FIX','True'),('42','DISP','False'),('BOX','DISP','False'),('TOP','FORCE','False')]:
        load.Range('A2:D2').Value2=((1,selector,'Y',mode),)
        assert call()==expected,(selector,mode);checks[selector+'_'+mode]='PASS'
    result={'status':'PASS','checks':checks,'original_unchanged':True}
    (a.out/'results.json').write_text(json.dumps(result,indent=2),encoding='utf-8');print(json.dumps(result))
finally:
    if wb is not None:wb.Close(False)
    app.Quit()
