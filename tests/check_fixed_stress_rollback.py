"""Check rejected fixed-stress cases against their last accepted checkpoint.

Checks: accepted time/count from FIXED_STRESS_V3 STATE, node ux/uy, all
checkpoint pore pressures, the four directly persisted stress components
(sx, sy, txy, sz) for every Gauss point, and START_PROFILE.  The checkpoint
does not persist every VBA Stmat diagnostic slot, so those slots are reported
as an explicit coverage gap rather than guessed.
"""
import argparse, csv, json, math, pathlib, sys

TOL = 2e-8
def nums(v): return float(v)
def close(a,b): return math.isclose(a,b, rel_tol=2e-8, abs_tol=TOL)
def fail(msg): print('FAIL: '+msg, file=sys.stderr); return 1

ap=argparse.ArgumentParser(); ap.add_argument('--run', required=True); ap.add_argument('--report', required=True); a=ap.parse_args()
run=pathlib.Path(a.run); exp=json.loads((run/'expected_failure.json').read_text(encoding='utf-8'))
fs=sorted(run.rglob('fixed_stress.csv'), key=lambda p:p.stat().st_mtime)
if not fs: sys.exit(fail('no fixed_stress.csv checkpoint found'))
cp=fs[-1]; root=cp.parent
lines=cp.read_text(encoding='utf-8').splitlines()
state=next((x.split(',') for x in lines if x.startswith('STATE,')), None)
profile=next((x.split(',',1)[1] for x in lines if x.startswith('START_PROFILE,')), None)
if not state: sys.exit(fail('checkpoint lacks STATE'))
if not close(float(state[1]), float(exp['coupled'][1])): sys.exit(fail(f'time mismatch {state[1]} != {exp["coupled"][1]}'))
if int(state[2]) != int(exp['coupled'][2]): sys.exit(fail(f'accepted count mismatch {state[2]} != {exp["coupled"][2]}'))
if profile != exp.get('profile'): sys.exit(fail(f'profile mismatch {profile!r} != {exp.get("profile")!r}'))

nodefile=root/'node.csv'; rows=list(csv.DictReader(nodefile.open(encoding='utf-8')))
byid={int(r['node']):r for r in rows}
for n,(nid,ux,uy) in enumerate(exp['nodes']):
    r=byid.get(int(nid))
    if not r: sys.exit(fail(f'missing node {nid}'))
    if not close(float(r['ux']),ux) or not close(float(r['uy']),uy): sys.exit(fail(f'node {nid} displacement mismatch'))

water=root/'water.csv'
if not water.exists(): sys.exit(fail('checkpoint lacks water.csv pressure table'))
wlines=water.read_text(encoding='utf-8').splitlines(); start=wlines.index('element,gauss,pressure_kPa')+1
p=[float(r.split(',')[2]) for r in wlines[start:] if r.strip()]
if len(p) != len(exp['pressure']): sys.exit(fail(f'pressure count {len(p)} != {len(exp["pressure"])}'))
for i,(x,y) in enumerate(zip(p,exp['pressure'])):
    if not close(x,y): sys.exit(fail(f'pressure mismatch at {i}: {x} != {y}'))

gfile=root/'gauss.csv'; grows=list(csv.DictReader(gfile.open(encoding='utf-8')))
if len(grows)*17 != len(exp['stress']): sys.exit(fail('stress evidence shorter than checkpoint gauss rows'))
for i,r in enumerate(grows):
    for j,k in ((0,'sx'),(1,'sy'),(2,'txy'),(16,'sz')):
        if not close(float(r[k]), float(exp['stress'][i*17+j])): sys.exit(fail(f'stress {k} mismatch at gauss {i}'))

report={'status':'PASS','checkpoint':str(cp),'case':exp.get('case'),'time':float(state[1]),'accepted':int(state[2]),'profile':profile,
        'checked':['STATE time/accepted count','START_PROFILE','all node ux/uy','all water.csv pressures','gauss sx/sy/txy/sz'],
        'coverage_gap':'Stmat slots beyond sx/sy/txy/sz and boundary head internals are not persisted in FIXED_STRESS_V3 CSV; no assertion made.'}
pathlib.Path(a.report).write_text(json.dumps(report,indent=2),encoding='utf-8'); print('PASS')
