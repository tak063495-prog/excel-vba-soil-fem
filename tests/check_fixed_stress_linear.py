"""Diagnostic 1-D Terzaghi comparison for the linear-ramp fixed-stress column.

This independently derives constrained modulus, storage, and diffusivity from
the fixture.  It compares excess pressure at reconstructed Q8 Gauss points and
the average top settlement against the analytical ramp convolution, then hold.
No numerical tolerance is treated as an acceptance criterion here.
"""
import argparse
import csv
import json
import math
from pathlib import Path


def read_water(path):
    lines = path.read_text(encoding="utf-8-sig").splitlines()
    i = lines.index("element,gauss,pressure_kPa")
    return {(int(r["element"]) - 1, int(r["gauss"]) - 1): float(r["pressure_kPa"])
            for r in csv.DictReader(lines[i:])}


def read_nodes(path):
    with path.open(newline="", encoding="utf-8-sig") as f:
        return {int(r["node"]): float(r["uy"]) for r in csv.DictReader(f)}


def time_of(folder):
    lines = (folder / "fixed_stress.csv").read_text(encoding="utf-8-sig").splitlines()
    state = next(line.split(",") for line in lines if line.startswith("STATE,"))
    return float(state[1])


def q8_y(coords, gp):
    z = 1.0 / math.sqrt(3.0)
    xi, eta = [(-z, -z), (z, -z), (z, z), (-z, z)][gp]
    n = [-(1-xi)*(1-eta)*(1+xi+eta)/4, -(1+xi)*(1-eta)*(1-xi+eta)/4,
         -(1+xi)*(1+eta)*(1-xi-eta)/4, -(1-xi)*(1+eta)*(1+xi-eta)/4,
         (1-xi*xi)*(1-eta)/2, (1+xi)*(1-eta*eta)/2,
         (1-xi*xi)*(1+eta)/2, (1-xi)*(1-eta*eta)/2]
    return sum(w * p[1] for w, p in zip(n, coords))


def parameters(fixture):
    material = fixture["materials"][0]
    e, nu = float(material[1]), float(material[2])
    alpha, storage, kx = map(float, fixture["water"]["coupledMaterials"][0][1:4])
    gamma = float(fixture["water"]["gamma"])
    constrained = e * (1.0 - nu) / ((1.0 + nu) * (1.0 - 2.0 * nu))
    cv = kx / (gamma * (storage + alpha * alpha / constrained))
    q = abs(sum(float(row[4]) for row in fixture["loads"] if row[3] == "FORCE"))
    top = max(float(n[2]) for n in fixture["nodes"]); bottom = min(float(n[2]) for n in fixture["nodes"])
    ramp = float(fixture["water"]["stagePlans"][0][3]) * 86400.0  # fixture specifies D in timeOptions
    initial_effective_weight = float(material[4]) - gamma
    gravity_top = -initial_effective_weight * (top-bottom)**2 / (2.0 * constrained)
    return constrained, cv, alpha * q / (alpha * alpha + storage * constrained), q, top-bottom, ramp, gravity_top, gamma


def terms(z, h, cv, t, ramp, p0, count=800):
    """Excess pressure for a top-drained / bottom-impermeable linear load ramp."""
    total = 0.0
    for m in range(count):
        odd = 2*m + 1; lam = odd*odd*math.pi*math.pi*cv/(4*h*h)
        coefficient = 4*p0/(math.pi*odd) * math.sin(odd*math.pi*z/(2*h))
        if t <= ramp:
            total += coefficient * (1.0 - math.exp(-lam*t)) / (lam*ramp)
        else:
            total += coefficient * (math.exp(-lam*(t-ramp)) - math.exp(-lam*t)) / (lam*ramp)
    return total


def settlement(t, h, cv, ramp, final, count=800):
    total = 0.0
    for m in range(count):
        odd = 2*m + 1; lam = odd*odd*math.pi*math.pi*cv/(4*h*h)
        coefficient = 8.0/(math.pi*math.pi*odd*odd)
        if t <= ramp:
            total += coefficient * (1.0 - math.exp(-lam*t)) / (lam*ramp)
        else:
            total += coefficient * (math.exp(-lam*(t-ramp)) - math.exp(-lam*t)) / (lam*ramp)
    return -final * ((min(t, ramp) / ramp) - total)


def main(fixture_file, run_root, report):
    fixture = json.loads(Path(fixture_file).read_text(encoding="utf-8")); run_root = Path(run_root)
    nodes = {int(n[0]): (float(n[1]), float(n[2])) for n in fixture["nodes"]}
    elements = {i: [nodes[int(n)] for n in e[1:9]] for i, e in enumerate(fixture["elements"])}
    md, cv, p0, q, h, ramp, gravity_top, gamma = parameters(fixture)
    top_ids = [int(n[0]) for n in fixture["nodes"] if abs(float(n[2]) - h) < 1e-12]
    frames = []
    for folder in sorted(run_root.rglob("C*")):
        if not (folder / "water.csv").exists(): continue
        t, field, disp = time_of(folder), read_water(folder / "water.csv"), read_nodes(folder / "node.csv")
        perr = []
        for (e, gp), p in field.items():
            depth = h - q8_y(elements[e], gp)
            hydrostatic = gamma * depth
            perr.append(abs((p - hydrostatic) - terms(depth, h, cv, t, ramp, p0)))
        actual = sum(disp[n] - gravity_top for n in top_ids) / len(top_ids)
        expected = settlement(t, h, cv, ramp, q*h/md)
        frames.append({"folder": str(folder), "time_s": t, "max_pressure_error_kPa": max(perr),
                       "mean_pressure_error_kPa": sum(perr)/len(perr), "top_settlement_m": actual,
                       "ramp_series_settlement_m": expected, "settlement_error_m": abs(actual-expected)})
    if not frames: raise RuntimeError("no fixed-stress time frames found")
    result = {"status": "DIAGNOSTIC", "constrained_modulus_kPa": md, "cv_m2_s": cv,
              "ramp_duration_s": ramp, "p0_kPa": p0, "gravity_top_settlement_m": gravity_top, "frames": frames,
              "max_pressure_error_kPa": max(r["max_pressure_error_kPa"] for r in frames),
              "max_settlement_error_m": max(r["settlement_error_m"] for r in frames),
              "scope": "independent 1-D ramp/hold series; hydrostatic and gravity baselines are derived from fixture data because this time-only export has no stg01_GRAVITY folder; report only, pending mesh/time refinement criterion"}
    Path(report).write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--fixture", required=True); ap.add_argument("--run", required=True); ap.add_argument("--report", required=True)
    args = ap.parse_args(); main(args.fixture, args.run, args.report)
