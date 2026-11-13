"""Verify real front-sheet edits, dirty marking and remesh preservation."""
from pathlib import Path
import argparse,json,shutil,win32com.client
ap=argparse.ArgumentParser();ap.add_argument('--book',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args();a.out=a.out.resolve();a.out.mkdir(parents=True,exist_ok=True)
root=Path(__file__).resolve().parents[1];copy=a.out/'face_ui_audit.xlsm';assert not copy.exists();shutil.copy2(a.book,copy);shutil.copy2(a.book.resolve().parent/'soilfem_native_x64.dll',a.out/'soilfem_native_x64.dll')
app=win32com.client.DispatchEx('Excel.Application');app.DisplayAlerts=False;app.EnableEvents=False;app.AutomationSecurity=1;wb=None
try:
 wb=app.Workbooks.Open(str(copy),UpdateLinks=0);app.Calculation=-4135
 for source in (root/'src/vba').glob('*.bas'):
  try:component=wb.VBProject.VBComponents(source.stem)
  except Exception:component=wb.VBProject.VBComponents.Add(1);component.Name=source.stem
  cm=component.CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(source.read_text(encoding='utf-8-sig'))
 cm=wb.VBProject.VBComponents('ThisWorkbook').CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString((root/'src/vba/ThisWorkbook.cls').read_text(encoding='utf-8-sig'))
 mod=wb.VBProject.VBComponents.Add(1);mod.Name='FaceUiObserver';mod.CodeModule.AddFromString('''
Public Sub FaceUiSetClean()
 AnalysisOK=True:ResultRevision=1:P1ResultReady=True:SuppressUserMessages=True
End Sub
Public Function FaceUiFlags() As Variant
 FaceUiFlags=Array(AnalysisOK,ResultRevision,P1ResultReady)
End Function
Public Function FaceUiMeshMessage() As String
 FaceUiMeshMessage=FEMUiMeshPreviewError
End Function
''')
 prefix="'"+wb.Name+"'!"
 # Follow Workbook_Open's canonical-settings refresh before rendering front pages.
 for key in ('NATIVE_SPARSE_DIRECT','NATIVE_WATER_TANGENT','ASSOCIATED_ENERGY_FIRST'):app.Run(prefix+'FEMWriteNumericSetting',key,1.)
 app.Run(prefix+'FEMWriteBuildStamp',False);app.Run(prefix+'FEM24Install');app.Run(prefix+'FEM24Water')
 app.EnableEvents=True;page=wb.Worksheets('境界・水条件');page.Range('C4').Value2='壁面支持'
 raw=wb.Worksheets('壁面支持');old=raw.Range('E2').Value2;assert page.Range('E8').Value2==old
 app.Run(prefix+'FaceUiSetClean');page.Range('E8').Value2=old*2.;assert raw.Range('E2').Value2==old*2.
 front=list(app.Run(prefix+'FaceUiFlags'));assert front==[False,0,False],front
 app.Run(prefix+'FaceUiSetClean');raw.Range('E2').Value2=old;direct=list(app.Run(prefix+'FaceUiFlags'));assert direct==[False,0,False],direct
 before=wb.Worksheets('節点データ').Range('A1:I20').Value2
 app.Run(prefix+'ボタン19_Click')
 assert wb.Worksheets('節点データ').Range('A1:I20').Value2==before,'Remeshing changed ID-based support nodes'
 assert '壁面支持' in app.Run(prefix+'FaceUiMeshMessage')
 options={}
 for key in ('NATIVE_SPARSE_DIRECT','NATIVE_WATER_TANGENT','ASSOCIATED_ENERGY_RECOVERY','ASSOCIATED_ENERGY_FIRST'):
  found=[]
  for name in ('メッシュ設定','解析設定','高速化設定','出力・表示','詳細設定'):
   sheet=wb.Worksheets(name);last=sheet.Cells(sheet.Rows.Count,6).End(-4162).Row
   for row,values in enumerate(sheet.Range(f'C1:F{last}').Value2,1):
    if values[3]==key:found.append((sheet,row))
  assert len(found)==1,(key,[(s.Name,r) for s,r in found])
  sheet,row=found[0];cell=sheet.Cells(row,3);original=cell.Value2
  assert cell.Validation.Type==3 and cell.Validation.Formula1.replace(';',',')=='0,1',(key,cell.Validation.Formula1)
  for value in (0.,1.,original):
   cell.Value2=value
   assert app.Run(prefix+'FEMReadSetting',key,-1.)==value,(key,value)
  options[key]={'page':sheet.Name,'row':row,'preserved_value':original,'write_through':'PASS'}
 record={'status':'PASS','front_setting_write_through':True,'front_edit_dirty_flags':front,'raw_edit_dirty_flags':direct,'remesh_preserved_nodes':True,'stiffness_value':old}
 record['solver_options']=options
 (a.out/'results.json').write_text(json.dumps(record,indent=2),encoding='utf-8');print(json.dumps(record),flush=True)
finally:
 if wb is not None:wb.Close(False)
 app.Quit()
