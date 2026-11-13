"""Check production SRM material transformations against independent formulas."""
from pathlib import Path
import argparse, json, math, shutil
import win32com.client

ap = argparse.ArgumentParser()
ap.add_argument('--book', type=Path, required=True)
ap.add_argument('--out', type=Path, required=True)
ap.add_argument('--embedded', action='store_true')
ap.add_argument('--expect-old-cap', action='store_true')
a = ap.parse_args()
root = Path(__file__).resolve().parents[1]
a.out.mkdir(parents=True, exist_ok=False)
dest = a.out.resolve() / 'psi_policy_audit.xlsm'
shutil.copy2(a.book, dest)
app = win32com.client.DispatchEx('Excel.Application')
app.DisplayAlerts = False
app.EnableEvents = False
w = None
try:
    w = app.Workbooks.Open(str(dest), UpdateLinks=0)
    for ws in w.Worksheets:
        if ws.ProtectContents:
            ws.Protect(Password='', DrawingObjects=False, Contents=True,
                       Scenarios=True, UserInterfaceOnly=True)
    if not a.embedded:
        for source in (root / 'src/vba').glob('*.bas'):
            cm = w.VBProject.VBComponents(source.stem).CodeModule
            cm.DeleteLines(1, cm.CountOfLines)
            cm.AddFromString(source.read_text(encoding='utf-8-sig'))
    cm = w.VBProject.VBComponents.Add(1).CodeModule
    cm.AddFromString('''Option Explicit
Public Function PsiPolicyProbe(ByVal flow As String, ByVal policy As String, ByVal phi As Double, ByVal psi As Double, ByVal reduce As Boolean, ByVal fs As Double) As Variant
 On Error GoTo Failed
 FEMWriteTextSetting "FLOW_POLICY",flow
 FEMWriteTextSetting "SRM_PSI_POLICY",policy
 P3PolicyCacheReady=False
 NumberOfMaterial=1
 ReDim Material(0 To 0):ReDim P3OriginalMaterial(0 To 0)
 With P3OriginalMaterial(0)
  .young=10000#:.Poisson=0.3:.thickness=1#:.weight=18#
  .kind="SOIL":.fai=phi:.psai=psi:.cohesion=12#
  .ReduceStrength=reduce:.InitialActive=True
 End With
 ' Isolate the real transformation without unrelated element generations.
 P3CurrentStrengthFactor=fs
 If Not P3SetMaterialForStrengthFactor(fs) Then
  PsiPolicyProbe=Array("ERROR",AnalysisMessage):Exit Function
 End If
 PsiPolicyProbe=Array("PASS",Material(0).fai,Material(0).psai,Material(0).cohesion):Exit Function
Failed:
 PsiPolicyProbe=Array("ERROR",Err.Description)
End Function
''')
    records = []
    mismatches = []
    for flow in ('DAVIS', 'INCONSISTENT'):
        for policy in ('CAP', 'REDUCE', 'KEEP'):
            for phi, psi in ((30., 30.), (35., 5.)):
                for fs in (.8, 1., 1.2, 2.):
                    for reduce in (True, False):
                        p, q, c = phi, psi, 12.
                        if reduce:
                            p = math.degrees(math.atan(math.tan(math.radians(phi)) / fs))
                            if policy == 'REDUCE':
                                q = math.degrees(math.atan(math.tan(math.radians(psi)) / fs))
                            if policy != 'KEEP':
                                q = min(q, p)
                            c /= fs
                        if flow == 'DAVIS' and abs(p-q) > 1e-12:
                            pr, qr = map(math.radians, (p, q))
                            ratio = math.cos(qr)*math.cos(pr)/(1-math.sin(pr)*math.sin(qr))
                            c *= ratio
                            p = math.degrees(math.atan(math.tan(pr)*ratio))
                            q = p
                        actual = list(app.Run("'" + w.Name + "'!PsiPolicyProbe",
                                              flow, policy, phi, psi, reduce, fs))
                        assert actual[0] == 'PASS', actual
                        error = max(abs(x-y) for x, y in zip(actual[1:], (p, q, c)))
                        record = dict(flow=flow, policy=policy, original_phi=phi,
                                      original_psi=psi, reduce=reduce, fs=fs,
                                      expected=[p,q,c], actual=actual[1:], error=error)
                        records.append(record)
                        if error > 1e-10:
                            mismatches.append(record)
    if a.expect_old_cap:
        assert mismatches and all(r['flow']=='DAVIS' and r['policy']=='CAP'
                                  and r['reduce'] and r['fs']>1 for r in mismatches), mismatches
    else:
        assert not mismatches, mismatches
    report = dict(status='OLD_CAP_REPRODUCED' if a.expect_old_cap else 'PASS',
                  cases=len(records), mismatches=mismatches, records=records)
    (a.out / 'results.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    print(report['status'], 'cases=', len(records), 'mismatches=', len(mismatches), flush=True)
finally:
    if w is not None:
        w.Close(False)
    app.Quit()
