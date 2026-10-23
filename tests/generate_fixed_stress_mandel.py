"""Create the Q8/P1 quarter-Mandel fixed-stress fixture.

The prescribed top displacement is the GEOS Mandel series evaluated at each
time point.  ``載荷`` column K selects a named load set:
M000 is the undrained t=0 state and M001..M100 are absolute stage targets.
"""
import json
import math
from pathlib import Path

from scipy.optimize import brentq

A, B, NX, NY = 1.0, 0.25, 16, 2
E, NU, ALPHA, STORAGE, K, GAMMA, Q = 10000.0, 0.3, 1.0, 1e-6, 1e-8, 9.81, 10.0


def constants():
    g = E / (2.0 * (1.0 + NU))
    kdr = E / (3.0 * (1.0 - 2.0 * NU))
    m = 1.0 / STORAGE
    ku = kdr + ALPHA * ALPHA * m
    nuu = (3.0 * ku - 2.0 * g) / (2.0 * (3.0 * ku + g))
    biot_b = ALPHA * m / ku
    cv = K / (GAMMA * (STORAGE + ALPHA * ALPHA / (kdr + 4.0 * g / 3.0)))
    return g, kdr, m, ku, nuu, biot_b, cv


def roots(count=1000):
    _, _, _, _, nuu, _, _ = constants()
    ratio = (1.0 - NU) / (nuu - NU)
    fun = lambda r: math.tan(r) - ratio * r
    values = []
    # The physical fundamental is in (0,pi/2); subsequent roots are in
    # (n*pi,n*pi+pi/2).  Start just right of zero to exclude the trivial root.
    for n in range(count):
        lo = n * math.pi + 1e-12
        hi = (n + 0.5) * math.pi - 1e-12
        values.append(brentq(fun, lo, hi))
    return values


ROOTS = roots()


def uy(t):
    g, _, _, _, nuu, _, cv = constants()
    if t == 0.0:
        return -Q * (1.0 - nuu) * B / (2.0 * g)
    tail = sum(math.sin(r) * math.cos(r) / (r - math.sin(r) * math.cos(r)) *
               math.exp(-r * r * cv * t / (A * A)) for r in ROOTS)
    return -Q * (1.0 - NU) * B / (2.0 * g) + Q * (1.0 - nuu) * B * tail / g


def build(dtau=.002, steps=100):
    nodes, lookup = [], {}
    def node(ix, iy):
        key = (ix, iy)
        if key not in lookup:
            ident = len(nodes) + 1
            x, y = A * ix / (2 * NX), B * iy / (2 * NY)
            # ux symmetry at x=0, uy support at y=0.
            nodes.append([ident, x, y, int(ix == 0), int(iy == 0), 0, 0, 0, 0])
            lookup[key] = ident
        return lookup[key]

    elements = []
    for j in range(NY):
        for i in range(NX):
            i0, j0 = 2 * i, 2 * j
            q8 = [node(i0, j0), node(i0 + 2, j0), node(i0 + 2, j0 + 2), node(i0, j0 + 2),
                  node(i0 + 1, j0), node(i0 + 2, j0 + 1), node(i0 + 1, j0 + 2), node(i0, j0 + 1)]
            elements.append([len(elements) + 1, *q8, 1])

    _, _, _, _, _, _, cv = constants()
    dt = dtau * A * A / cv
    top = [n[0] for n in nodes if abs(n[2] - B) < 1e-14]
    stages = [[1, "GRAVITY", "", 1, 1, "hydrostatic initial equilibrium"],
              [2, "LOAD", 1, 1, 1, "Mandel undrained t=0"]]
    plans = []
    options = []
    loads = []
    loads.append([1, "TOP", "Y", "DISP", uy(0.0), None, None, None, None, "", "M000"])
    plans.append([2, "FIXED_STRESS", "", 0.0, 0.0, "", 200, 1e-8, "undrained"])
    options.append([2, "S", "HOLD", "INSTANT", 0.0, 0.0, 0.0, "", "M000", "undrained"])
    for step in range(1, steps+1):
        stage = step + 2
        set_id = f"M{step:03d}"
        stages.append([stage, "LOAD", 1, 1, 1, f"Mandel t={step}dt"])
        plans.append([stage, "FIXED_STRESS", "", dt, dt, "", 200, 1e-8, set_id])
        options.append([stage, "S", "HOLD", "LINEAR", dt / 1024.0, dt, dt, "", set_id, "Mandel target"])
        loads.append([1, "TOP", "Y", "DISP", uy(step * dt), None, None, None, None, "", set_id])

    return {
        "name": "fixed_stress_mandel",
        "nodes": nodes,
        "elements": elements,
        "materials": [[1, E, NU, 1, GAMMA, 30, 1e9, 0, 1, 1, "SOIL", 0, 0]],
        "stages": stages,
        "loads": loads,
        "settings": {"OUTPUT_STAGE_MODE": "ALL"},
        "water": {
            "mode": "HYDROSTATIC", "gamma": GAMMA, "defaultProfile": "W", "auto": "OFF",
            "profiles": [["W", 0.0, B, "top head"], ["W", A, B, "top head"]],
            "materials": [[1, "EFFECTIVE", K, "*", ""]],
            "boundaries": [[NX, 2, "EXPOSED", "W", None, "lower right exposed"],
                           [2 * NX, 2, "EXPOSED", "W", None, "upper right exposed"]],
            "stagePlans": plans,
            "coupledMaterials": [[1, ALPHA, STORAGE, K, K, "Mandel", 0.0]],
            "timeOptions": options,
        },
        "mandel": {"a_m": A, "b_m": B, "nx": NX, "ny": NY, "q_kPa": Q,
                   "dt_s": dt, "cv_m2_s": cv, "uy_t0_m": uy(0.0), "roots": ROOTS},
    }


if __name__ == "__main__":
    import argparse
    ap=argparse.ArgumentParser()
    ap.add_argument('--nx',type=int,default=16)
    ap.add_argument('--dtau',type=float,default=.002)
    ap.add_argument('--steps',type=int,default=100)
    ap.add_argument('--name',default='fixed_stress_mandel')
    ap.add_argument('--out',required=True)
    args=ap.parse_args()
    if args.nx<1 or args.dtau<=0 or not 1<=args.steps<=1900: ap.error('invalid mesh/time/stage count')
    NX=args.nx
    assert abs(uy(1e-12)-uy(0.0))<1e-8, 'Mandel root-series t=0 identity failed'
    fixture=build(args.dtau,args.steps);fixture['name']=args.name
    fixture['settings']['EXPORT_MODE']='STAGE'
    Path(args.out).write_text(json.dumps(fixture,indent=2),encoding='utf-8')
