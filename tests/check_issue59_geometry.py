"""Independent Q8 signed-Jacobian evidence for issue 59.

This checker deliberately uses dense sampling and a symbolic polynomial as an
independent audit; production code is expected to use Bernstein certification.
"""
import json
from pathlib import Path
import numpy as np
try:
    import sympy as sp
except ImportError:
    sp = None

ROOT = Path(__file__).resolve().parents[1]
FIX = ROOT / "tests" / "fixtures" / "issue59" / "q8_jacobian_cases.json"

def deriv(xi, eta):
    dx = [.25*(1-eta)*(2*xi+eta), .25*(1-eta)*(2*xi-eta), .25*(1+eta)*(2*xi+eta), .25*(1+eta)*(2*xi-eta), -xi*(1-eta), .5*(1-eta*eta), -xi*(1+eta), -.5*(1-eta*eta)]
    de = [.25*(1-xi)*(xi+2*eta), .25*(1+xi)*(-xi+2*eta), .25*(1+xi)*(xi+2*eta), .25*(1-xi)*(-xi+2*eta), -.5*(1-xi*xi), -eta*(1+xi), .5*(1-xi*xi), -eta*(1-xi)]
    return np.array(dx), np.array(de)

def detj(xi, eta, p):
    p = p - p[0]
    dx, de = deriv(xi, eta)
    return float(np.linalg.det([[dx @ p[:, 0], de @ p[:, 0]], [dx @ p[:, 1], de @ p[:, 1]]]))

def audit(name, coords):
    p = np.asarray(coords, dtype=float)
    grid = np.linspace(-1.0, 1.0, 161)
    dense = min(detj(x, y, p) for x in grid for y in grid)
    boundary = min([detj(x, z, p) for x in grid for z in (-1.0, 1.0)] + [detj(z, y, p) for y in grid for z in (-1.0, 1.0)])
    old_points = [(-1,-1),(1,-1),(1,1),(-1,1),(-1/np.sqrt(3),-1/np.sqrt(3)),(1/np.sqrt(3),-1/np.sqrt(3)),(1/np.sqrt(3),1/np.sqrt(3)),(-1/np.sqrt(3),1/np.sqrt(3))]
    old = min(detj(x, y, p) for x, y in old_points)
    edge_min = []
    t = np.linspace(-1., 1., 4)
    for edge in (lambda z: detj(z, -1., p), lambda z: detj(z, 1., p),
                 lambda z: detj(-1., z, p), lambda z: detj(1., z, p)):
        coef = np.polynomial.polynomial.polyfit(t, [edge(z) for z in t], 3)
        roots = np.polynomial.polynomial.polyroots(np.polynomial.polynomial.polyder(coef))
        candidates = [-1., 1.] + [r.real for r in roots if abs(r.imag) < 1e-8 and -1. <= r.real <= 1.]
        edge_min.append(min(edge(z) for z in candidates))
    if sp is None:
        poly = "sympy-unavailable"
        center = float(detj(0, 0, p))
    else:
        x, y = sp.symbols("xi eta")
        dsx = [sp.Rational(1,4)*(1-y)*(2*x+y), sp.Rational(1,4)*(1-y)*(2*x-y), sp.Rational(1,4)*(1+y)*(2*x+y), sp.Rational(1,4)*(1+y)*(2*x-y), -x*(1-y), sp.Rational(1,2)*(1-y*y), -x*(1+y), -sp.Rational(1,2)*(1-y*y)]
        dsy = [sp.Rational(1,4)*(1-x)*(x+2*y), sp.Rational(1,4)*(1+x)*(-x+2*y), sp.Rational(1,4)*(1+x)*(x+2*y), sp.Rational(1,4)*(1-x)*(-x+2*y), -sp.Rational(1,2)*(1-x*x), -y*(1+x), sp.Rational(1,2)*(1-x*x), -y*(1-x)]
        rp = [[sp.Rational(str(z)) for z in row] for row in coords]
        poly_expr = sp.expand(sum(dsx[i]*rp[i][0] for i in range(8))*sum(dsy[i]*rp[i][1] for i in range(8)) - sum(dsy[i]*rp[i][0] for i in range(8))*sum(dsx[i]*rp[i][1] for i in range(8)))
        poly = str(poly_expr)
        center = float(poly_expr.subs({x:0, y:0}))
    return {"name": name, "boundary_min": boundary, "boundary_extrema_min": min(edge_min), "legacy_8point_min": old, "dense_161_min": dense, "detJ_polynomial": poly, "center_detJ": center}

def find_random_internal():
    rng = np.random.default_rng(59)
    grid = np.linspace(-1.0, 1.0, 11)
    base = np.array([[-1,-1],[1,-1],[1,1],[-1,1],[-.5,-1],[1,0],[.5,1],[-1,0]], dtype=float)
    for _ in range(20000):
        p = base.copy(); p[4:] += rng.normal(0.0, 0.8, (4, 2))
        b = min([detj(x,z,p) for x in grid for z in (-1,1)] + [detj(z,y,p) for y in grid for z in (-1,1)])
        old = min(detj(x,y,p) for x,y in [(-1,-1),(1,-1),(1,1),(-1,1),(-1/np.sqrt(3),-1/np.sqrt(3)),(1/np.sqrt(3),-1/np.sqrt(3)),(1/np.sqrt(3),1/np.sqrt(3)),(-1/np.sqrt(3),1/np.sqrt(3))])
        dense = min(detj(x,y,p) for x in grid for y in grid)
        if b > 1e-3 and old > 1e-3 and dense < -1e-3:
            return p.tolist()
    raise AssertionError("deterministic random search did not find an interior inversion")

def main():
    cases = json.loads(FIX.read_text(encoding="utf-8"))["cases"]
    extra = json.loads((FIX.parent / 'q8_subdivision_case.json').read_text(encoding='utf-8'))['cases']
    cases += extra
    report = [audit(c["name"], c["coordinates"]) for c in cases]
    t = np.linspace(0., 1., 4)
    inverse = np.linalg.inv(np.array([[(1-z)**3,3*z*(1-z)**2,3*z*z*(1-z),z**3] for z in t]))
    for c in extra:
        p = np.array(c['coordinates'], dtype=float)
        values = np.array([[detj(2*x-1,2*y-1,p) for y in t] for x in t])
        lower = float((inverse @ values @ inverse.T).min())
        assert lower < 0 and abs(lower-c['root_bernstein_min']) < 1e-10
        next(r for r in report if r['name'] == c['name'])['root_bernstein_min'] = lower
    for r in report:
        if r["name"] in {"internal_inversion", "random_internal"}:
            if r["name"] == "random_internal":
                assert r["boundary_extrema_min"] > 0 and r["legacy_8point_min"] > 0
            assert r["dense_161_min"] < 0
        elif r["name"] in {"square", "rotated_scaled", "curved_valid", "square_small", "near_degenerate", "scaled_translated", "curved_subdivision"}:
            assert r["dense_161_min"] > 0
        else:
            assert r["dense_161_min"] <= 0
    print(json.dumps({"status":"PASS", "cases":report}, ensure_ascii=False, indent=2))

if __name__ == "__main__": main()
