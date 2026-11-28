"""Import repository VBA into a disposable workbook copy and compile in Excel."""
from pathlib import Path
import argparse,hashlib,json,shutil
import win32com.client
from check_embedded_sources import normalize

ap=argparse.ArgumentParser()
ap.add_argument('--book',type=Path,required=True)
ap.add_argument('--out',type=Path,required=True)
args=ap.parse_args();source=args.book.resolve();out=args.out.resolve()
out.mkdir(parents=True,exist_ok=False)
root=Path(__file__).resolve().parents[1]
target=out/'source_import.xlsm';shutil.copyfile(source,target)
shutil.copyfile(source.parent/'soilfem_band_fast_x64.dll',out/'soilfem_band_fast_x64.dll')
original_hash=hashlib.sha256(source.read_bytes()).hexdigest()
app=win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Excel.Application'))
app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.AutomationSecurity=1
hashes={}
try:
    book=app.Workbooks.Open(str(target),UpdateLinks=0)
    catalog=json.loads((root/'src/vba/modules.json').read_text(encoding='utf-8-sig'))
    for item in catalog:
        name=item['name'];p=root/'src/vba'/item['file']
        code=p.read_text(encoding='utf-8-sig')
        code='\n'.join(l for l in code.splitlines() if not l.startswith('Attribute VB_'))
        try:component=book.VBProject.VBComponents(name)
        except Exception:
            assert item['type']==1,name
            component=book.VBProject.VBComponents.Add(1);component.Name=name
        cm=component.CodeModule
        if cm.CountOfLines:cm.DeleteLines(1,cm.CountOfLines)
        cm.AddFromString(code)
        hashes[p.relative_to(root).as_posix()]=hashlib.sha256(p.read_bytes()).hexdigest()
    control=app.VBE.CommandBars.FindControl(1,578)
    if control.Enabled:control.Execute()
    assert not app.VBE.CommandBars.FindControl(1,578).Enabled,'Native VBA compile failed'
    for item in catalog:
        cm=book.VBProject.VBComponents(item['name']).CodeModule
        assert normalize(cm.Lines(1,cm.CountOfLines))==normalize((root/'src/vba'/item['file']).read_text(encoding='utf-8-sig')),item['name']
    record=dict(passed=True,native_excel_compile=True,source_modules=len(hashes),workbook_sha256=original_hash,repository_source_sha256=hashes)
    (out/'source_import_result.json').write_text(json.dumps(record,indent=2)+'\n',encoding='utf8')
    print('NATIVE_SOURCE_IMPORT_COMPILE_PASS',len(hashes),'modules',flush=True)
finally:
    for b in list(app.Workbooks):b.Close(False)
    app.Quit()
assert hashlib.sha256(source.read_bytes()).hexdigest()==original_hash
