"""Independent pressure-history check for the prescribed-displacement Mandel fixture.

Pressure is compared at the physical Q8 Gauss coordinates reconstructed from
the fixture.  The imposed top displacement is checked only as a prescribed
reference curve; it is deliberately not presented as an independent FE result.
"""
import argparse
import csv
import json
import math
from functools import lru_cache
from pathlib import Path

from scipy.optimize import brentq


def constants(m):
    e, nu = m['materials'][0][1:3]
    alpha,storage,k,ky = m['water']['coupledMaterials'][0][1:5]
    gamma=m['water']['gamma']
    if k != ky: raise ValueError('Mandel series requires isotropic permeability')
    g = e / (2 * (1 + nu)); kdr = e / (3 * (1 - 2 * nu)); modulus = 1 / storage
    ku = kdr + alpha * alpha * modulus
    nuu = (3 * ku - 2 * g) / (2 * (3 * ku + g)); biot_b = alpha * modulus / ku
    cv = k / (gamma * (storage + alpha * alpha / (kdr + 4 * g / 3)))
    return g, nuu, biot_b, cv


@lru_cache(maxsize=4)
def roots(nuu, nu=.3, count=1000):
    ratio = (1-nu) / (nuu-nu)
    return [brentq(lambda r: math.tan(r) - ratio * r, n * math.pi + 1e-12,
                   (n + .5) * math.pi - 1e-12) for n in range(count)]


def pressure(x, t, fixture):
    g, nuu, biot_b, cv = constants(fixture)
    nu=fixture['materials'][0][2]
    a, q = fixture["mandel"]["a_m"], fixture["mandel"]["q_kPa"]
    p0 = q * biot_b * (1 + nuu) / 3
    return 2 * p0 * sum(math.sin(r) / (r - math.sin(r) * math.cos(r)) *
                        (math.cos(r * x / a) - math.cos(r)) *
                        math.exp(-r * r * cv * t / (a * a)) for r in roots(nuu,nu))


def top_displacement(t, fixture):
    g, nuu, _, cv = constants(fixture)
    nu=fixture['materials'][0][2]
    a, b, q = fixture["mandel"]["a_m"], fixture["mandel"]["b_m"], fixture["mandel"]["q_kPa"]
    if t == 0.0: return -q * (1 - nuu) * b / (2 * g)
    return -q * (1-nu) * b / (2 * g) + q * (1 - nuu) * b / g * sum(
        math.sin(r) * math.cos(r) / (r - math.sin(r) * math.cos(r)) * math.exp(-r*r*cv*t/(a*a)) for r in roots(nuu,nu))


def displacement(x, y, t, fixture):
    # Plane-strain equilibrium: sigma_x=Md*ux,x+lambda*uy,y-alpha*p=0.
    # Integrate the independent Mandel pressure series, with ux(0)=0.
    g, nuu, biot_b, cv = constants(fixture)
    a, b, q = (fixture['mandel'][k] for k in ('a_m', 'b_m', 'q_kPa'))
    e, nu = fixture['materials'][0][1:3]
    lam = e*nu/((1+nu)*(1-2*nu)); md=e*(1-nu)/((1+nu)*(1-2*nu))
    alpha=fixture['water']['coupledMaterials'][0][1]
    p0=q*biot_b*(1+nuu)/3
    if t == 0:
        integral=p0*x
    else:
        integral=2*p0*sum(math.sin(r)/(r-math.sin(r)*math.cos(r))*
            (a/r*math.sin(r*x/a)-x*math.cos(r))*math.exp(-r*r*cv*t/(a*a)) for r in roots(nuu,nu))
    strain_y=top_displacement(t,fixture)/b
    return (alpha*integral-lam*strain_y*x)/md, strain_y*y


def q8_point(coords, gp):
    z = 1 / math.sqrt(3); xi, eta = [(-z, -z), (z, -z), (z, z), (-z, z)][gp]
    n = [-(1-xi)*(1-eta)*(1+xi+eta)/4, -(1+xi)*(1-eta)*(1-xi+eta)/4,
         -(1+xi)*(1+eta)*(1-xi-eta)/4, -(1-xi)*(1+eta)*(1+xi-eta)/4,
         (1-xi*xi)*(1-eta)/2, (1+xi)*(1-eta*eta)/2,
         (1-xi*xi)*(1+eta)/2, (1-xi)*(1-eta*eta)/2]
    return sum(w*p[0] for w,p in zip(n, coords)), sum(w*p[1] for w,p in zip(n, coords))


def water(path):
    with path.open(newline="", encoding="utf-8-sig") as f:
        rows = list(f)
    header = next(i for i, line in enumerate(rows) if line.strip() == "element,gauss,pressure_kPa")
    return {(int(r["element"]) - 1, int(r["gauss"]) - 1): float(r["pressure_kPa"])
            for r in csv.DictReader(rows[header:]) if r.get("element")}


def state_time(folder):
    with (folder / "fixed_stress.csv").open(encoding="utf-8-sig") as f:
        line = next(f).strip(); state = next(f).strip().split(",")
    if not line.startswith(("FIXED_STRESS_V2,", "FIXED_STRESS_V3,")) or state[0] != "STATE":
        raise RuntimeError(f"invalid fixed-stress checkpoint: {folder}")
    return float(state[1])


def node_displacements(path):
    with path.open(newline="", encoding="utf-8-sig") as f:
        return {int(r["node"]): float(r["uy"]) for r in csv.DictReader(f) if r.get("node")}


def node_vectors(path):
    with path.open(newline="", encoding="utf-8-sig") as f:
        return {int(r['node']):(float(r['ux']),float(r['uy'])) for r in csv.DictReader(f) if r.get('node')}


def main(fixture_path, run, report):
    fixture = json.loads(Path(fixture_path).read_text(encoding="utf-8")); run = Path(run)
    baseline_dirs = list(run.rglob("stg01_GRAVITY"))
    if len(baseline_dirs) == 1:
        baseline = water(baseline_dirs[0] / "water.csv")
        baseline_uy = node_displacements(baseline_dirs[0] / "node.csv")
        baseline_vector = node_vectors(baseline_dirs[0] / "node.csv")
    elif not baseline_dirs:
        # EXPORT_MODE=OFF still retains OUTPUT_STAGE_MODE=ALL sheet records.
        rows = json.loads((run / "水圧結果.json").read_text(encoding="utf-8-sig"))
        baseline = {(int(r[1])-1, int(r[2])-1): float(r[5]) for r in rows[35:] if r[0] == 1}
        rows = json.loads((run / "結果節点.json").read_text(encoding="utf-8-sig"))
        baseline_uy = {int(r[1]): float(r[5]) for r in rows[1:] if r[0] == 1}
        baseline_vector = {int(r[1]): (float(r[4]),float(r[5])) for r in rows[1:] if r[0] == 1}
        if not baseline or not baseline_uy: raise RuntimeError("initial gravity records missing")
    else: raise RuntimeError("ambiguous gravity checkpoint")
    nodes = {int(n[0]): (float(n[1]), float(n[2])) for n in fixture["nodes"]}
    coords = {i: [nodes[int(v)] for v in e[1:9]] for i, e in enumerate(fixture["elements"])}
    centre_key = min(baseline, key=lambda key: (q8_point(coords[key[0]], key[1])[0], abs(q8_point(coords[key[0]], key[1])[1] - fixture['mandel']['b_m']/2)))
    frames = []
    for folder in sorted(run.rglob("C*")):
        if not folder.is_dir() or not (folder / "water.csv").exists(): continue
        t, values = state_time(folder), water(folder / "water.csv")
        errors = []
        for (e, gp), p in values.items():
            if (e, gp) not in baseline: raise RuntimeError("Gauss key missing from gravity baseline")
            x, _ = q8_point(coords[e], gp)
            if t > 0.0: errors.append(abs((p - baseline[e, gp]) - pressure(x, t, fixture)))
        top = [n[0] for n in fixture["nodes"] if abs(float(n[2]) - fixture["mandel"]["b_m"]) < 1e-14]
        got = node_displacements(folder / "node.csv")
        vectors = node_vectors(folder / 'node.csv')
        ux_error=uy_error=0.
        for node,(x,y) in nodes.items():
            ux,uy=displacement(x,y,t,fixture)
            ux_error=max(ux_error,abs(vectors[node][0]-baseline_vector[node][0]-ux))
            uy_error=max(uy_error,abs(vectors[node][1]-baseline_vector[node][1]-uy))
        prescribed = top_displacement(t, fixture)
        disp_error = max(abs((got[int(n)] - baseline_uy[int(n)]) - prescribed) for n in top)
        excess_max = max(p - baseline[key] for key, p in values.items())
        frames.append({"folder": str(folder), "time_s": t, "max_pressure_error_kPa": max(errors) if errors else None,
                       "mean_pressure_error_kPa": sum(errors) / len(errors) if errors else None, "top_uy_reference_m": prescribed,
                       "top_uy_prescribed_error_m": disp_error, "max_excess_pressure_kPa": excess_max,
                       "centre_excess_pressure_kPa": values[centre_key]-baseline[centre_key],
                       "all_node_ux_error_m":ux_error,"all_node_uy_error_m":uy_error})
    if len(frames) < 100: raise RuntimeError(f"expected >=100 accepted time frames, found {len(frames)}")
    transient = [r for r in frames if r["max_pressure_error_kPa"] is not None]
    if not transient: raise RuntimeError("no t>0 transient time frame")
    peak = max(transient, key=lambda r: r["max_pressure_error_kPa"])
    pressure_history = [r["centre_excess_pressure_kPa"] for r in frames]
    cryer = max(pressure_history) > pressure_history[0] and pressure_history[-1] < max(pressure_history)
    result = {"status": "DIAGNOSTIC_REFINEMENT_REQUIRED", "frames": frames, "max_transient_pressure_error_kPa": peak["max_pressure_error_kPa"],
              "worst_frame": peak, "scope": "initial t=0 discrete pressure projection is excluded; transient pressure uses GEOS Mandel series at reconstructed physical Gauss points"}
    result["cryer_peak_detected"] = cryer
    result["pass_policy"] = "No PASS is emitted: establish a grid-refinement convergence criterion before acceptance."
    try:
        import matplotlib.pyplot as plt
        plt.plot([r["time_s"] for r in frames], pressure_history, marker=".")
        plt.xlabel("time [s]"); plt.ylabel("centre excess pressure [kPa]"); plt.title("Mandel/Cryer pressure history")
        plt.tight_layout(); plt.savefig(Path(report).with_suffix(".png"), dpi=160); plt.close()
    except ImportError: result["plot"] = "matplotlib unavailable"
    Path(report).write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    p = argparse.ArgumentParser(); p.add_argument("--fixture", required=True); p.add_argument("--run", required=True); p.add_argument("--report", required=True)
    a = p.parse_args(); main(a.fixture, a.run, a.report)
