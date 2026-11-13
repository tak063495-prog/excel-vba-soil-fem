"""Real Excel regression for unused-point headings (#69) and JOINT/support input (#70)."""
from pathlib import Path
import argparse, json, shutil
import win32com.client

ap = argparse.ArgumentParser()
ap.add_argument('--book', type=Path, required=True, help='Compiled face_support/static fixture workbook')
ap.add_argument('--out', type=Path, required=True)
a = ap.parse_args()
root = Path(__file__).resolve().parents[1]
a.out = a.out.resolve()
assert not a.out.exists(), 'Preserve previous evidence'
a.out.mkdir(parents=True)
copy = a.out / 'pr68_followups.xlsm'
shutil.copy2(a.book, copy)
shutil.copy2(a.book.resolve().parent / 'soilfem_native_x64.dll', a.out / 'soilfem_native_x64.dll')
app = win32com.client.DispatchEx('Excel.Application')
app.DisplayAlerts = False
app.EnableEvents = False
app.AutomationSecurity = 1
w = None
records = []
try:
    w = app.Workbooks.Open(str(copy), UpdateLinks=0)
    for ws in w.Worksheets:
        if ws.ProtectContents:
            ws.Unprotect('')
    for name in ('FEMUiGeometry', 'FEMFaceSupport'):
        cm = w.VBProject.VBComponents(name).CodeModule
        cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString((root / 'src/vba' / (name + '.bas')).read_text(encoding='utf-8-sig'))
    w.VBProject.VBComponents('FEM').CodeModule.AddFromString('''
Public Function Pr68InputProbe() As String
 On Error GoTo Failed
 FEMInvalidateInputCache
 P6LoadInputCache
 Pr68InputProbe="PASS":Exit Function
Failed:Pr68InputProbe="ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function Pr68AnalysisProbe() As Variant
 SuppressUserMessages=True:P0_RunAnalysis
 Pr68AnalysisProbe=Array(ResultStatus,AnalysisMessage,TotalSelfWeight,P6ReactionSumX,P6ReactionSumY)
End Function
''')
    compile_control = app.VBE.CommandBars.FindControl(1, 578)
    if compile_control.Enabled:
        compile_control.Execute()
    assert not app.VBE.CommandBars.FindControl(1, 578).Enabled, 'Probe compile failed'
    prefix = "'" + w.Name + "'!"
    app.Run(prefix + 'FEMWriteTextSetting', 'EXPORT_LOAD', '')
    app.Run(prefix + 'FEMWriteTextSetting', 'EXPORT_MODE', 'OFF')
    nodes = w.Worksheets('節点データ')
    mats = w.Worksheets('材料データ')
    elems = w.Worksheets('要素データ')
    supports = w.Worksheets('壁面支持')
    for i, (x, y) in enumerate(((1, 0), (0, 0), (0.5, 0)), 14):
        nodes.Range(f'A{i+1}:I{i+1}').Value2 = ((i, x, y, 1, 1, 0, 0, 0, 0),)
    mats.Range('A4:M4').Value2 = ((3, 100000, 0, 1, 0, 0, 1000000, 0, 0, 1, 'JOINT', 1000000, 1000000),)
    for active in (1, 0):
        for padding in (None, 0):
            mats.Range('J4').Value2 = active
            elems.Range('A4:J4').Value2 = ((3, 1, 2, 5, 14, 15, 16, padding, padding, 3),)
            result = app.Run(prefix + 'Pr68InputProbe')
            assert result == 'PASS', (active, padding, result)
            records.append({'issue': 70, 'active_joint': active, 'padding': padding, 'input': result})
            if active and padding is None:
                result = list(app.Run(prefix + 'Pr68AnalysisProbe'))
                assert result[0] == 'PASS' and abs(result[2] - 36) < 1e-8 and abs(result[3] - 1) < 1e-8 and abs(result[4] - 36) < 1e-8, result
                records[-1]['full_static_analysis'] = result
    app.Run(prefix + 'FaceSupportWriteResults')
    assert '準備未完了' in str(w.Worksheets('壁面支持結果').Range('A1').Value2)
    records.append({'issue': 70, 'case': 'unmapped_result_guard', 'status': 'PASS'})
    for case, node_id in (('joint_only', 14), ('unconnected', 17)):
        if node_id == 17:
            nodes.Range('A18:I18').Value2 = ((17, 0, 3, 1, 1, 0, 0, 0, 0),)
        mats.Range('J4').Value2 = 1
        supports.Range('B2').Value2 = node_id
        result = app.Run(prefix + 'Pr68InputProbe')
        assert result.startswith('ERROR|') and '初期から有効な連続体要素' in result, (case, result)
        records.append({'issue': 70, 'case': case, 'rejected': result})
    supports.Range('B2').Value2 = 10
    shape = w.Worksheets('形状入力')
    legacy = w.Worksheets('図形定義')
    native_reference = None
    for unused in (2, 1, 0):
        shape.Range('B8:D2000').ClearContents()
        shape.Range('G8:J2000').ClearContents()
        shape.Range('B8:D11').Value2 = ((10, 0, 0), (20, 2, 0), (30, 2, 2), (40, 0, 2))
        for i in range(unused):
            shape.Range(f'B{12+i}:D{12+i}').Value2 = ((900+i, 100+i, 100+i),)
        shape.Range('H8:J8').Value2 = (('OUTER', '10,20,30,40', 1),)
        app.Run(prefix + 'FEMUiBuildGeometry')
        # Four used points imply diagnostics start at row 26.
        assert legacy.Range('A32:B32').Value2 == (('未使用の点番号', '元の行'),)
        assert legacy.Range('B27:B29').Value2 == ((4+unused,), (4,), (unused,))
        values = legacy.Range('A33:B34').Value2
        assert tuple((int(v[0]), int(v[1])) for v in values[:unused]) == tuple((900+i, 12+i) for i in range(unused)), values
        assert all(v is None for row in values[unused:] for v in row), values
        native = (legacy.Range('A1:N6').Value2, legacy.Range('D24:N24').Value2)
        if native_reference is None:
            native_reference = native
        assert native == native_reference, 'Unused ledger points changed mesh-generator input'
        records.append({'issue': 69, 'unused': unused, 'headings_preserved': True, 'native_geometry_unchanged': True})
    report = {'status': 'PASS', 'cases': records}
    (a.out / 'results.json').write_text(json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False, allow_nan=False), flush=True)
finally:
    if w is not None:
        w.Close(False)
    app.Quit()
