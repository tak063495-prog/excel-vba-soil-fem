"""Independent equilibrium, geometry and loading-strip settlement checks."""
import argparse
import hashlib
import json
from pathlib import Path
import numpy as np
import check_manual_case as geom


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def check(path, fixture_path=None, baseline_path=None):
    path = Path(path).resolve()
    d = load(path)
    if d['status'] != 'PASS':
        return dict(status='FAIL', solver_status=d['status'], q_kPa=None,
                    reason='Nonconverged trial is not a successful pressure result')
    f = load(fixture_path or d['fixture'])
    p = f['bearing_protocol']
    bp = Path(baseline_path) if baseline_path else path.parent.parent / p['baseline'] / 'result.json'
    assert hashlib.sha256(bp.read_bytes()).hexdigest() == p['baseline_result_sha256']
    b = load(bp)
    assert b['status'] == 'PASS' and b['application_code_unchanged'] and d['application_code_unchanged']
    t = d['tables']
    for key in ['節点データ', '要素データ', '材料データ']:
        assert t[key] == b['tables'][key], key
    nodes = {int(r[0]): r for r in t['節点データ'][1:]}
    mats = {int(r[0]): r for r in t['材料データ'][1:]}
    foot = sorted(i for i, n in nodes.items() if 1-1e-9 <= n[1] <= 3+1e-9 and abs(n[2]-5) < 1e-9)
    assert len(foot) >= 3 and foot == sorted(p['footing_nodes'])
    assert set(foot) == {int(r[1]) for r in f['loads']}
    s, width, thickness = float(p['additional_settlement_m']), float(p['width_m']), float(p['thickness_m'])
    assert s > 0 and width == 2 and thickness == mats[1][3]
    assert abs(max(nodes[i][1] for i in foot)-min(nodes[i][1] for i in foot)-width) < 1e-10
    states = {(int(r[0]), int(r[1])): r for r in t['結果節点'][1:]}
    base = {(int(r[0]), int(r[1])): r for r in b['tables']['結果節点'][1:]}
    target_errors, increment_errors = [], []
    for r in f['loads']:
        i = int(r[1]); ref = p['reference_uy_m'][str(i)]
        assert r[0] == 1 and r[2:4] == ['Y', 'DISP']
        assert abs(ref-base[(1, i)][5]) < 1e-12 and abs(r[4]-(ref-s)) < 1e-12
        assert abs(states[(1, i)][5]-ref) < 1e-9
        target_errors.append(abs(states[(2, i)][5]-r[4]))
        increment_errors.append(abs(states[(2, i)][5]-states[(1, i)][5]+s))
    assert max(target_errors) < 1e-9 and max(increment_errors) < 1e-9
    geom.ROOT = path.parent.parent
    geometry = geom.check(path.parent.name)
    areas = {}
    for e in t['要素データ'][1:]:
        m = int(e[9]); xy = np.array([nodes[int(i)][1:3] for i in e[1:9]], float)
        areas[m] = areas.get(m, 0.) + geom.quadrature(xy)
    weight = sum(a*mats[m][3]*mats[m][4] for m, a in areas.items())
    records = []
    res2 = force2 = 0.
    for stage in [1, 2]:
        sx = sy = free_x = free_y = 0.
        for i, n in nodes.items():
            r = states[(stage, i)]; cx = bool(n[3]); cy = bool(n[4]) or (stage == 2 and i in foot)
            if cx: sx += r[8]
            else: free_x += r[8]
            if cy: sy += r[9]
            else: free_y += r[9]
            if stage == 2:
                for j, fixed in [(0, cx), (1, cy)]:
                    if not fixed: res2 += r[8+j]**2; force2 += r[6+j]**2
        logged = next(r for r in t['ステージ結果'][1:] if r[0] == stage)
        assert logged[3] == 'PASS'
        tol = max(1e-5, abs(weight)*1e-8)
        assert abs(sx-logged[6]) < tol and abs(sy-logged[7]) < tol
        assert abs(sy+free_y-weight) < tol and abs(sx+free_x) < tol
        records.append(dict(stage=stage, support_Rx_kN=sx, support_Ry_kN=sy,
                            free_residual_sum_y_kN=free_y, weight_kN=weight,
                            equilibrium_closure_y_kN=sy+free_y-weight))
    diag = {r[0]: r[1] for r in t['診断'][1:] if r[0]}
    rel = float(np.sqrt(res2)/max(np.sqrt(force2), 1e-12))
    assert abs(rel-diag['RelativeResidualFree']) < 1e-9
    reaction = {str(i): states[(2, i)][9] for i in foot}
    q = -sum(reaction.values())/(width*thickness)
    assert q > 0, 'Expected compressive loading-strip pressure'
    actual = -float(np.mean([states[(2, i)][5]-states[(1, i)][5] for i in foot]))
    return dict(status='PASS', solver_status=d['status'], q_kPa=q, actual_settlement_m=actual,
                requested_settlement_m=s, target_error_max_m=max(target_errors),
                increment_error_max_m=max(increment_errors), width_m=width, thickness_m=thickness,
                seconds=d['seconds'], relative_residual=diag['RelativeResidualFree'],
                reconstructed_relative_residual=rel, footing_nodes=foot, footing_reaction_y_kN=reaction,
                tension_reaction_node_count=sum(v > 1e-8 for v in reaction.values()),
                stage_equilibrium=records, geometry=geometry, no_contact_model=True,
                interpretation='Pressure at prescribed additional settlement; ultimate bearing capacity is not automatically determined')


if __name__ == '__main__':
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('result', type=Path); ap.add_argument('--fixture', type=Path); ap.add_argument('--baseline', type=Path)
    args = ap.parse_args()
    record = check(args.result, args.fixture, args.baseline)
    (args.result.parent/'bearing_checks.json').write_text(json.dumps(record, ensure_ascii=False, indent=2)+'\n', encoding='utf8')
    print(record['status'], 'q_kPa=', record['q_kPa'])
    raise SystemExit(0 if record['status'] == 'PASS' else 2)
