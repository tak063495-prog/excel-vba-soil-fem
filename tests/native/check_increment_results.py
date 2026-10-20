"""Independent checks on the four-GP increment table exported by the Excel test."""
import argparse
import json
import math
from pathlib import Path


def check(rows_path, run_path):
    rows = json.loads(rows_path.read_text(encoding="utf-8-sig"))
    run = json.loads(run_path.read_text(encoding="utf-8-sig"))
    if isinstance(run, list):
        assert len(run) == 1
        run = run[0]
    assert rows
    errors, groups = [], {}
    for a in rows:
        assert len(a) == 20
        assert a[17] == run["Bracket"][0], "increment data must follow restored successful Fs"
        assert a[19] == 1, "final accepted load factor"
        groups.setdefault(a[1], []).append(a[2])

        def eig(x, y, gamma):
            center = (x + y) / 2
            radius = math.hypot((x - y) / 2, gamma / 2)
            return center + radius, center - radius

        total = eig(*a[5:8])
        plastic = eig(*a[8:11])
        expected = [max(*total, 0) - min(*total, 0),
                    max(*plastic, a[11]) - min(*plastic, a[11]), *plastic]
        errors.extend(abs(x - y) for x, y in zip(expected, a[12:16]))
        c, s = math.cos(a[16]), math.sin(a[16])
        reconstructed = [plastic[0] * c*c + plastic[1] * s*s,
                         plastic[0] * s*s + plastic[1] * c*c,
                         2 * (plastic[0] - plastic[1]) * s*c]
        errors.extend(abs(x - y) for x, y in zip(reconstructed, a[8:11]))
    assert all(sorted(g) == [1, 2, 3, 4] for g in groups.values())
    assert max(errors) < 1e-14, max(errors)
    return dict(status="PASS", gauss_rows=len(rows), elements=len(groups), fs=run["Bracket"][0],
                lambda_final=1, max_abs_formula_error=max(errors),
                nonzero_total=sum(a[12] > 0 for a in rows),
                nonzero_plastic=sum(a[13] > 0 for a in rows))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", type=Path, required=True)
    parser.add_argument("--run", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    result = check(args.rows, args.run)
    args.out.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print("PASS all Gauss point principal values, 3D engineering max shear and successful Fs")
