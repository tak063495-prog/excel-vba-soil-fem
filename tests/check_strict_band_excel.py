"""Compare the actual VBA private routines against the optional native kernel."""
from pathlib import Path
import argparse, hashlib, json, shutil
import win32com.client

parser=argparse.ArgumentParser()
parser.add_argument('--book', type=Path, required=True)
parser.add_argument('--out', type=Path, required=True)
parser.add_argument('--actual', action='store_true')
args=parser.parse_args(); args.book=args.book.resolve(); args.out=args.out.resolve()
args.out.mkdir(parents=True, exist_ok=False)
before=hashlib.sha256(args.book.read_bytes()).hexdigest()
target=args.out/'kernel_test.xlsm'; shutil.copyfile(args.book,target)
shutil.copyfile(args.book.parent/'soilfem_band_fast_x64.dll',args.out/'soilfem_band_fast_x64.dll')
app=win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx('Excel.Application'))
app.Visible=False; app.DisplayAlerts=False; app.EnableEvents=False; app.AutomationSecurity=1
try:
    book=app.Workbooks.Open(str(target),UpdateLinks=0)
    cm=book.VBProject.VBComponents('FEMSolver').CodeModule
    cm.InsertLines(3,'Private Declare PtrSafe Function QA_MemCmp Lib "msvcrt" Alias "memcmp" (ByVal a As LongPtr, ByVal b As LongPtr, ByVal n As LongPtr) As Long')
    source=cm.Lines(1,cm.CountOfLines)
    code=(Path(__file__).parent/'strict_band_driver.bas').read_text(encoding='utf8')
    for declaration in ['Public QA_ActualCompared As Boolean','Public QA_ActualRecord As Variant']:
        source=source.replace('Option Explicit','Option Explicit\r\n'+declaration,1)
        code=code.replace(declaration,'')
    if args.actual:
        needle='PackedFactorOk:\r\n'; assert source.count(needle)==1
        source=source.replace(needle,needle+'  If Not QA_ActualCompared Then QA_CompareActualBand n2, BandWidth, pivotTolerance\r\n',1)
    cm.DeleteLines(1,cm.CountOfLines); cm.AddFromString(source+'\r\n'+code)
    def call(name): return app.Run("'"+book.Name+"'!"+name)
    if args.actual:
        module=book.VBProject.VBComponents.Add(1); module.Name='StrictBandTest'
        module.CodeModule.AddFromString('Public Function StrictBandActualRun() As Variant\nSuppressUserMessages=True\nP0_RunAnalysis\nStrictBandActualRun=QA_ActualRecord\nEnd Function')
        result=call('StrictBandActualRun'); passed=bool(result[0])
    else:
        result=call('QA_KernelEquality'); passed=all(row[3] for row in result)
    record=dict(passed=passed,actual_matrix=args.actual,result=result,source_sha256=before)
    (args.out/'report.json').write_text(json.dumps(record,indent=2),encoding='utf8')
    assert passed
finally:
    for book in list(app.Workbooks): book.Close(False)
    app.Quit()
assert hashlib.sha256(args.book.read_bytes()).hexdigest()==before
print('PASS: bitwise equality, including failure decisions')
