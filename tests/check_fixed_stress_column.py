"""Independent Terzaghi comparison for a saved fixed-stress column result.

Pressure rows are matched by (x,y), never by workbook row order.  This is a
diagnostic checker; it does not claim that the finite element problem is the
same as the one-dimensional series problem.
"""
import argparse, json, math
from pathlib import Path


def series(z, H, cv, t, p0, terms=500):
    return p0 * sum(4/(math.pi*(2*m+1))*math.sin((2*m+1)*math.pi*z/(2*H)) *
                    math.exp(-((2*m+1)**2*math.pi**2*cv*t)/(4*H**2))
                    for m in range(terms))


def main(root):
    root = Path(root)
    rows = json.loads((root/'水圧結果.json').read_text(encoding='utf-8'))
    gp = {}
    for row in rows[35:]:
        if isinstance(row, list) and len(row) > 5 and isinstance(row[0], (int, float)):
            gp.setdefault(int(row[0]), []).append({'x': float(row[3]), 'y': float(row[4]), 'p': float(row[5])})
    if not {1, 2, 3}.issubset(gp):
        raise RuntimeError('expected pressure stages 1,2,3')
    s = 1e-6; alpha = 1.; E = 10000.; nu = .3; k = 1e-8; gamma = 9.81; H = 2.; t = 86400.; q = 10.
    Md = E*(1-nu)/((1+nu)*(1-2*nu))
    cv = k/(gamma*(s + alpha*alpha/Md))
    p0 = alpha*q/(alpha*alpha + s*Md)
    stage = lambda rows: {(round(v['x'], 10), round(v['y'], 10)): v for v in rows}
    a, b, c = stage(gp[1]), stage(gp[2]), stage(gp[3])
    matched = []
    for key, r3 in c.items():
        if key in a and key in b:
            excess0 = b[key]['p'] - a[key]['p']; excess1 = r3['p'] - a[key]['p']
            matched.append({'x': key[0], 'y': key[1], 'p0_fe': excess0, 'p1_fe': excess1,
                            'p1_terzaghi': series(H-key[1], H, cv, t, p0)})
    if not matched:
        raise RuntimeError('no coordinate-matched pressure points')
    err = [abs(v['p1_fe']-v['p1_terzaghi']) for v in matched]
    out = {'status':'DIAGNOSTIC', 'points':len(matched), 'Md_kPa':Md, 'cv_m2_s':cv,
           'p0_kPa':p0, 'Tv':cv*t/(H*H), 'max_abs_error_kPa':max(err),
           'mean_abs_error_kPa':sum(err)/len(err), 'points_data':matched,
           'scope':'diagnostic series comparison; finite-element boundary/load projection is not assumed identical'}
    print(json.dumps(out, indent=2, ensure_ascii=False))


if __name__ == '__main__':
    ap = argparse.ArgumentParser(); ap.add_argument('root'); main(ap.parse_args().root)
