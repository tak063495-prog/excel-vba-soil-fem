"""Independent checks for the water-display regression package.

This checker does not start Excel.  It compares JSON records produced by the
fixture runner and reports measured runtimes without making them pass/fail
criteria.
"""
from __future__ import annotations

import argparse
import json
import math
from pathlib import Path


def load(path: Path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def flat(value):
    if isinstance(value, list):
        return [x for item in value for x in flat(item)]
    return [value]


def maxdiff(a, b):
    aa, bb = flat(a), flat(b)
    if len(aa) != len(bb):
        return math.inf
    return max((abs(float(x) - float(y)) for x, y in zip(aa, bb)), default=0.0)


def check_record(candidate: Path, reference: Path, label: str, report: dict):
    c, r = load(candidate), load(reference)
    if not c.get("status", "").startswith("PASS") or not r.get("status", "").startswith("PASS"):
        raise AssertionError(f"{label}: candidate/reference status is not PASS")
    if not all(isinstance(c.get(k), list) and c.get(k) for k in ("disp", "stress", "metrics")):
        raise AssertionError(f"{label}: candidate vectors are missing")
    if not all(isinstance(r.get(k), list) and r.get(k) for k in ("disp", "stress", "metrics")):
        raise AssertionError(f"{label}: reference vectors are missing")
    fields = ("disp", "stress", "metrics")
    diffs = {k: maxdiff(c.get(k, []), r.get(k, [])) for k in fields}
    tol = {"disp": 1e-10, "stress": 1e-7, "metrics": 1e-8}
    ok = c.get("status", "").startswith("PASS") and all(diffs[k] <= tol[k] for k in fields)
    report["records"].append({"label": label, "candidate": str(candidate), "reference": str(reference),
                              "maxdiff": diffs, "candidate_seconds": c.get("seconds"),
                              "reference_seconds": r.get("seconds"), "pass": ok})
    if not ok:
        raise AssertionError(f"{label}: status/diff mismatch {diffs}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", type=Path, required=True, help="final run directory, e.g. tests/tmp/water_display/final7")
    ap.add_argument("--baseline", type=Path, default=None, help="acceptance20 directory")
    ap.add_argument("--davis-baseline", type=Path, default=None, help="final5 DAVIS result.json")
    ap.add_argument("--ui", type=Path, default=None, help="optional UI evidence directory")
    args = ap.parse_args()
    root = args.root.resolve()
    baseline = (args.baseline or root.parents[1] / "groundwater" / "acceptance20").resolve()
    report = {"root": str(root), "baseline": str(baseline), "records": [], "rejections": [], "ui": {}}

    mapping = {
        "cutoff": "cutoff",
        "box_inside": "box_inside",
        "stages": "stages",
    }
    for candidate_name, reference_name in mapping.items():
        check_record(root / candidate_name / "result.json", baseline / reference_name / "result.json",
                     candidate_name, report)

    check_record(root / "wet_srm_INCONSISTENT" / "result.json", baseline / "wet_srm" / "result.json",
                 "wet_srm_INCONSISTENT", report)
    davis_ref = (args.davis_baseline or (root / "final" / "wet_srm_DAVIS" / "result.json")).resolve()
    if not davis_ref.exists():
        raise AssertionError(f"missing final5 DAVIS baseline: {davis_ref}")
    check_record(root / "wet_srm_DAVIS" / "result.json", davis_ref,
                 "wet_srm_DAVIS_final5_reference", report)

    one_way = load(root / "one_way" / "result.json")
    stages = load(root / "stages" / "result.json")
    for key in ("disp", "stress", "metrics"):
        d = maxdiff(one_way.get(key, []), stages.get(key, []))
        if d > {"disp": 1e-10, "stress": 1e-7, "metrics": 1e-8}[key]:
            raise AssertionError(f"one_way/stages {key} differs by {d}")
    report["one_way_vs_stages"] = {k: maxdiff(one_way.get(k, []), stages.get(k, []))
                                    for k in ("disp", "stress", "metrics")}
    tab=load(root/'box_inside'/'水圧結果.json')
    forces=[r for r in tab[35:] if len(r)>29 and isinstance(r[23],(float,int))]
    fx,fy,moment=(sum(float(r[c]) for r in forces) for c in (27,28,29))
    ry=load(root/'box_inside'/'result.json')['metrics'][2]
    assert abs(fx)<1e-8 and abs(fy-85)<1e-8 and abs(moment-127.5)<1e-8 and abs(ry-107)<1e-7
    report['box_inside_balance']={'Fx_kN':fx,'Fy_kN':fy,'M_kNm':moment,'Ry_kN':ry}

    # Resume records must reproduce the saved final state to numerical tolerance.
    for name in ("cutoff", "box_inside", "stages", "one_way", "joint", "wet_srm_INCONSISTENT"):
        p = root / name / "resume.json"
        if p.exists():
            c = load(root / name / "result.json")
            q = load(p)
            diffs = {k: maxdiff(c.get(k, []), q.get(k, [])) for k in ("disp", "stress", "metrics")}
            if any(v > t for v, t in zip(diffs.values(), (1e-9, 1e-5, 1e-7))):
                raise AssertionError(f"{name} resume differs: {diffs}")
            report.setdefault("resume", {})[name] = diffs
        elif name == "stages":
            raise AssertionError(f"missing required resume record: {p}")

    rejection_names = ("fixed_stress", "time_reserved", "missing_stage", "conflicting_profile",
                       "orphan", "duplicate", "unknown_profile")
    for name in rejection_names:
        candidates = [root / f"reject_{name}" / "input_check.json",
                      root.parent / f"reject_{name}" / "input_check.json"]
        path = next((p for p in candidates if p.exists()), None)
        if path is None:
            raise AssertionError(f"missing rejection record: {name}")
        item = load(path)
        if item.get("status") != "PASS" or not item.get("actual_error"):
            raise AssertionError(f"invalid rejection record: {path}")
        report["rejections"].append({"name": name, "path": str(path), "pass": True,
                                      "actual_error": item["actual_error"]})

    ui = args.ui or root.parent
    shape_files = [p for p in (list(ui.glob("**/preview.json")) + list(ui.glob("**/reopen.json")) + list(ui.glob("**/result.json")))
                   if any(part.lower().startswith("ui_") for part in p.parts)]
    report["ui"]["evidence_files"] = len(shape_files)
    if args.ui:
        if not shape_files:
            raise AssertionError(f"UI evidence requested but no action JSON found under {ui}")
        ui_checks = []
        for path in shape_files:
            item = load(path)
            if item.get("status") != "PASS" or not item.get("shapes"):
                raise AssertionError(f"invalid UI action record: {path}")
            names = {str(s.get("name", "")) for s in item["shapes"]}
            title=next((s.get('text','') for s in item['shapes'] if s.get('name')=='GW_VIEW_TITLE'),'')
            assert any(n.startswith('GW_VIEW_PROFILE_') for n in names),path
            if 'STEADY' in title:
                assert 'GW_VIEW_HEAD_RANGE' in names and any('HEAD_FIXED' in n for n in names),path
                rt=next(s['text'] for s in item['shapes'] if s.get('name')=='GW_VIEW_HEAD_RANGE')
                if '一定水頭' in rt:
                    assert not any(n.startswith('GW_VIEW_FLUX_') for n in names),path
                else:
                    assert any(n.startswith('GW_VIEW_FLUX_') for n in names),path
            elif 'HYDROSTATIC' in title:
                assert 'GW_VIEW_HEAD_RANGE' not in names and not any('HEAD_FIXED' in n or n.startswith('GW_VIEW_FLUX_') for n in names),path
            if "GW_VIEW_LEGEND" in names and "GW_VIEW_HEAD_RANGE" in names:
                legend = next(s for s in item["shapes"] if s.get("name") == "GW_VIEW_LEGEND")
                rng = next(s for s in item["shapes"] if s.get("name") == "GW_VIEW_HEAD_RANGE")
                if float(rng["top"]) + float(rng["height"]) > float(legend["top"]):
                    raise AssertionError(f"head range overlaps legend: {path}")
            ui_checks.append({"path": str(path), "shape_count": len(item["shapes"]),
                              "steady": any("HEAD_CONTOUR" in n or "FLUX" in n for n in names),
                              "fixed_head": any("HEAD_FIXED" in n for n in names)})
        report["ui"]["checks"] = ui_checks
        stale=list(ui.glob('**/stale_error.txt'))
        assert stale and any('ERROR|入力が変更' in p.read_text(encoding='utf-8-sig') for p in stale),'No verified stale-input rejection'
    report["ui"]["note"] = "UI action JSON is checked only when --ui is supplied; no Excel is started by this checker."
    report['status']='PASS'
    out = root / "water_display_check.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
