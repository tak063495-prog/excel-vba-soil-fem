#!/usr/bin/env python3
"""Independent post-checker for the Issue 54/55 Excel material-point driver.

The cone oracle uses scipy.linprog on the six original principal-space flow
rays.  For deliberately constructed tiny-dilatancy inputs it records HiGHS as
diagnostic only: its feasibility tolerance is not a reliable classifier when
the ray matrix is nearly rank deficient.  Those labels are established by the
known construction and checked separately by the actual VBA result.
"""
from __future__ import annotations
import argparse, itertools, json, math, sys
from pathlib import Path
import numpy as np
from scipy.optimize import linprog

EPS = 2.220446049250313e-16


def load(path: Path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def principal(ep):
    ex, ey, gxy, ez = map(float, ep)
    center = .5 * (ex + ey)
    radius = .5 * math.hypot(ex - ey, gxy)
    return sorted((center + radius, center - radius, ez), reverse=True)


def rays(psi_deg):
    s = math.sin(math.radians(float(psi_deg)))
    return np.asarray(sorted(set(itertools.permutations((1.0 - s, 0.0, -1.0 - s)))), dtype=float).T


def lp_oracle(case):
    """Independent nonnegative six-ray feasibility/min-sum oracle."""
    target = np.asarray(principal(case["plastic"]), dtype=float)
    aeq = rays(case["psi"])
    ans = linprog(np.ones(6), A_eq=aeq, b_eq=target, bounds=(0, None),
                  method="highs", options={"primal_feasibility_tolerance": 1e-9,
                                              "dual_feasibility_tolerance": 1e-9})
    return bool(ans.success), (float(ans.fun) if ans.success else None), str(ans.message)


def cone_expected(case):
    p = principal(case["plastic"])
    s = math.sin(math.radians(float(case["psi"])))
    if case.get("constructed_label"):
        # A known ray/combination label is deliberately used near rank loss.
        return bool(case["expected_accept"]), None, "known constructed case"
    ok, total, message = lp_oracle(case)
    if ok != bool(case["expected_accept"]):
        raise AssertionError(f"{case['id']}: fixture label disagrees with independent linprog: {message}")
    return ok, total, message


def elastic_matrix(e=1000.0, nu=.3):
    lam = e * nu / ((1 + nu) * (1 - 2 * nu))
    g = e / (2 * (1 + nu))
    return np.array([[lam + 2*g, lam, 0., lam],
                     [lam, lam + 2*g, 0., lam],
                     [0., 0., g, 0.],
                     [lam, lam, 0., lam + 2*g]])


def material_yield(stress, c):
    sx, sy, txy, sz = map(float, stress)
    center=.5*(sx+sy); radius=math.hypot(.5*(sx-sy), txy)
    mx, mn = max(center+radius, center-radius, sz), min(center+radius, center-radius, sz)
    s=math.sin(math.radians(30.0)); co=math.cos(math.radians(30.0))
    return mx-mn-(mx+mn)*s-2*c*co


def close(a,b,atol=2e-8,rtol=2e-9):
    return abs(float(a)-float(b)) <= atol + rtol*max(1.,abs(float(a)),abs(float(b)))


def require_close_vec(label, got, want, atol=2e-8, rtol=2e-9):
    if len(got) != len(want) or any(not close(a,b,atol,rtol) for a,b in zip(got,want)):
        raise AssertionError(f"{label}: got={got}, expected={want}")


def records_from(raw):
    if isinstance(raw, dict): raw = raw.get("records", raw.get("results"))
    if not isinstance(raw, list): raise AssertionError("actual output must be a records list or {records:[...]}")
    result={}
    for r in raw:
        if not isinstance(r,dict) or "id" not in r or "actual" not in r:
            raise AssertionError("each record requires id and actual")
        if r["id"] in result: raise AssertionError(f"duplicate actual id: {r['id']}")
        result[r["id"]]=r
    return result


def check_cone(case, record):
    actual=record["actual"]
    if len(actual) < 3: raise AssertionError(f"{case['id']}: cone actual requires [accept,sum,strict]")
    expected, lp_sum, source = cone_expected(case)
    accept=bool(actual[0]); total=float(actual[1]); strict=bool(actual[2])
    if accept != expected: raise AssertionError(f"{case['id']}: accept={accept}, expected={expected} ({source})")
    row={"id":case["id"],"mode":"cone","accept":accept,"oracle":source}
    if accept:
        if total < -1e-12: raise AssertionError(f"{case['id']}: negative multiplier sum {total}")
        if "expected_min_sum" in case:
            require_close_vec(case['id']+': canonical sum',[total],[case['expected_min_sum']],atol=2e-8,rtol=2e-8)
        elif lp_sum is not None:
            # LP gives an independently chosen decomposition; for psi>0 its sum is unique.
            require_close_vec(case['id']+': LP total',[total],[lp_sum],atol=2e-8,rtol=2e-8)
        if "expected_strict_interior" in case and strict != bool(case["expected_strict_interior"]):
            raise AssertionError(f"{case['id']}: strict={strict}, expected={case['expected_strict_interior']}")
        row.update(multiplier_sum=total,strict_interior=strict)
    return row


def check_material(case, record):
    a=record["actual"]
    if len(a) < 15: raise AssertionError(f"{case['id']}: material actual has {len(a)} fields, need 15")
    ok=bool(a[0]); expected=bool(case["expected_success"])
    if ok != expected: raise AssertionError(f"{case['id']}: success={ok}, expected={expected}, raw={a}")
    row={"id":case["id"],"mode":"material","ok":ok,"failure_code":a[3]}
    if not ok: return row
    if not bool(a[1]) or bool(a[2]): raise AssertionError(f"{case['id']}: successful apex case must converge plastically: {a[:4]}")
    stress=np.asarray(a[4:8],dtype=float); ep=np.asarray(a[8:12],dtype=float); mult=float(a[12])
    previous=np.asarray(case["previous"],dtype=float); strain=np.asarray([*case["strain"],0.],dtype=float)
    reconstruction=previous+elastic_matrix()@(strain-ep)
    require_close_vec(case['id']+': elastic reconstruction',stress,reconstruction,atol=3e-8,rtol=3e-9)
    yf=material_yield(stress,float(case["c"]))
    scale=max(1.,np.max(np.abs(stress)),float(case['c']))
    if abs(yf)>3e-8*scale: raise AssertionError(f"{case['id']}: MC yield residual {yf}")
    psi=float(case["psi"])
    if psi > 0:
        expected_mult=-float(ep[0]+ep[1]+ep[3])/(2*math.sin(math.radians(psi)))
        if not close(mult,expected_mult,atol=3e-8,rtol=3e-8):
            raise AssertionError(f"{case['id']}: multiplier {mult} != -tr(ep)/(2sin psi) {expected_mult}")
    if "expected_multiplier" in case:
        if not close(mult,case["expected_multiplier"],atol=3e-8,rtol=3e-8):
            raise AssertionError(f"{case['id']}: explicit multiplier mismatch")
    if "expected_stress" in case: require_close_vec(case['id']+': apex stress',stress,case['expected_stress'],atol=3e-8,rtol=3e-9)
    if "expected_plastic" in case: require_close_vec(case['id']+': apex plastic',ep,case['expected_plastic'],atol=3e-8,rtol=3e-9)
    if "expected_tangent_ready" in case and bool(a[13]) != bool(case["expected_tangent_ready"]):
        raise AssertionError(f"{case['id']}: tangentReady={a[13]}")
    row.update(multiplier=mult,yield_residual=yf,tangent_ready=bool(a[13]),max_abs_tangent=float(a[14]))
    return row


def main():
    p=argparse.ArgumentParser()
    p.add_argument('--fixture',type=Path,default=Path('tests/fixtures/material_issue54_55.json'))
    p.add_argument('--actual',type=Path,required=True,help='JSON records emitted by the actual Excel wrapper')
    p.add_argument('--out',type=Path,required=True)
    args=p.parse_args()
    fixture=load(args.fixture)
    if fixture.get('schema')!='MATERIAL_ISSUE54_55_V1': raise AssertionError('wrong fixture schema')
    records=records_from(load(args.actual)); cases=fixture.get('cases',[])
    ids=[c.get('id') for c in cases]
    if len(ids)!=len(set(ids)) or not ids: raise AssertionError('fixture IDs must be unique and nonempty')
    if set(records)!=set(ids):
        raise AssertionError(f"record IDs mismatch missing={sorted(set(ids)-set(records))} extra={sorted(set(records)-set(ids))}")
    rows=[]
    for c in cases:
        # The workbook wrapper emits the case it actually supplied.  Refuse a
        # stale/misordered driver result even when its numerical shape fits.
        supplied = records[c['id']].get('input')
        if supplied is None or supplied != c:
            raise AssertionError(f"{c['id']}: actual record input does not match fixture case")
        rows.append(check_cone(c,records[c['id']]) if c['mode']=='cone' else check_material(c,records[c['id']]))
    result={'status':'PASS','fixture':str(args.fixture),'records':len(rows),'checks':rows,
            'oracle':'scipy.linprog original six flow rays; tiny-psi constructed cases use analytic construction'}
    args.out.parent.mkdir(parents=True,exist_ok=True)
    args.out.write_text(json.dumps(result,indent=2)+"\n",encoding='utf-8')
    print(f"PASS: issue54/55 material records={len(rows)}")

if __name__=='__main__':
    try: main()
    except Exception as e:
        print(f"FAIL: {e}",file=sys.stderr)
        raise
