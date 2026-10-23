"""Check the issue #50 WATER resume matrix.

The runner may optionally emit ``resume_water.json`` and
``resume_水圧結果.json`` when the boundary observables hook is enabled.  The
checks below intentionally require those records: displacement-only resume
checks cannot detect a stale boundary profile when the saved pore pressure is
still readable.
"""
from __future__ import annotations
import argparse, json, math
from pathlib import Path


def load(p: Path):
    return json.loads(p.read_text(encoding="utf-8-sig"))


def finite_values(value):
    if isinstance(value, (int, float)):
        return [float(value)] if math.isfinite(float(value)) else []
    if isinstance(value, list):
        out = []
        for x in value:
            out.extend(finite_values(x))
        return out
    return []


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", type=Path, required=True)
    ap.add_argument("--out", type=Path, required=True)
    args = ap.parse_args()
    cases = ["default_start_no_detail", "explicit_hold", "explicit_start", "explicit_linear"]
    report = {"issue": 50, "cases": {}, "status": "PASS"}
    for suffix in cases:
        d = args.root / ("issue50_" + suffix)
        fresh = load(d / "result.json")
        resumed = load(d / "resume.json")
        if not str(fresh.get("status", "")).startswith("PASS|") or not str(resumed.get("status", "")).startswith("PASS|"):
            raise SystemExit(f"non-PASS run: {suffix}")
        disp = fresh.get("disp", [])
        rdisp = resumed.get("disp", [])
        if len(disp) != len(rdisp) or any(abs(float(a) - float(b)) > 1e-9 for a, b in zip(disp, rdisp)):
            raise SystemExit(f"resume displacement mismatch: {suffix}")
        # The fixture runner extension writes these values before closing the
        # fresh workbook and after reopening the checkpoint.  A missing file is
        # a test failure rather than an implicit pass.
        for name in ("water_observables.json", "resume_water_observables.json"):
            p = d / name
            if not p.exists():
                raise SystemExit(f"missing boundary observables: {d.name}/{name}")
        a = load(d / "water_observables.json")
        b = load(d / "resume_water_observables.json")
        keys = ("profile", "external_load", "flow", "next_wait_profile")
        missing = [k for k in keys if k not in a or k not in b]
        if missing:
            raise SystemExit(f"missing observable keys {missing}: {suffix}")
        for k in keys:
            av, bv = a[k], b[k]
            va, vb = finite_values(av), finite_values(bv)
            if len(va) != len(vb) or any(abs(x - y) > 1e-10 for x, y in zip(va, vb)):
                raise SystemExit(f"resume {k} mismatch: {suffix}")
        report["cases"][suffix] = {"status": "PASS", "dofs": len(disp), "observables": list(keys)}
    args.out.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("PASS")


if __name__ == "__main__":
    main()
