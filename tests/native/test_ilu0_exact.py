"""Bitwise sfm_ilu0 reference/candidate regression and microbenchmark.

The test deliberately includes valid sorted CSR and malformed (unsorted or
duplicate) rows.  It compares return code, failure pair, minimum pivot and
the complete ILU buffer, including partial output on failure.
"""
from __future__ import annotations
import argparse, ctypes as C, hashlib, json, math, pathlib, random, statistics, struct, time


def load(path: pathlib.Path):
    d = C.WinDLL(str(path))
    f = d.sfm_ilu0
    f.argtypes = [C.c_int32, C.c_int32, C.POINTER(C.c_int32),
                  C.POINTER(C.c_int32), C.POINTER(C.c_int32),
                  C.POINTER(C.c_double), C.POINTER(C.c_double),
                  C.POINTER(C.c_int32), C.POINTER(C.c_double)]
    f.restype = C.c_int32
    d.sfm_init.restype = C.c_int32
    if d.sfm_init() != 4:
        raise RuntimeError("sfm_init ABI mismatch: %s" % path)
    if hasattr(d,'sfm_probe_cw'):
        d.sfm_probe_cw.argtypes=[C.c_ushort];d.sfm_probe_cw.restype=C.c_ushort
    return d, f


def cases(count: int, seed: int = 0x1A):
    rng = random.Random(seed)
    out = []
    fixed = [(1, 0), (2, 1), (9, 2), (25, 4), (64, 7), (169, 8)]
    for n, bw in fixed:
        for malformed in (False, True):
            out.append(make_case(n, bw, rng, malformed, None))
    for i in range(count):
        n = rng.choice([1, 2, 3, 4, 7, 9, 16, 25, 49, 64, 100, 169])
        bw = rng.randrange(0, min(12, n))
        malformed = (i % 5 == 0)
        special = rng.choice([None, "zero", "tiny", "nan", "inf"] if i % 3 == 0 else [None])
        out.append(make_case(n, bw, rng, malformed, special, sparse=(i % 4 == 1)))
    return out


def make_case(n, bw, rng, malformed, special, sparse=False):
    rows = []
    for r in range(n):
        lo, hi = max(0, r - bw), min(n - 1, r + bw)
        cols = list(range(lo, hi + 1))
        if sparse and len(cols) > 3:
            cols = [c for c in cols if c == r or rng.random() < 0.55]
        if malformed and len(cols) > 2:
            rng.shuffle(cols)
            if r % 2 == 0:
                cols.insert(1, cols[0])
        else:
            cols.sort()
        rows.append(cols)
    rp, ci, vals, dp = [0], [], [], []
    for r, cols in enumerate(rows):
        start = len(ci)
        for c in cols:
            ci.append(c)
            vals.append((2.0 if c == r else 0.01) + rng.uniform(-1e-4, 1e-4))
        # sfm_ilu0 assumes a diagonal position; choose the first equal entry.
        try:
            dp.append(start + cols.index(r))
        except ValueError:
            dp.append(start)
        rp.append(len(ci))
    if special:
        p = dp[min(n - 1, rng.randrange(n))]
        vals[p] = {"zero": 0.0, "tiny": 1e-20, "nan": float("nan"), "inf": float("inf")}[special]
    return n, rp, ci, dp, vals


def bind(f, case):
    n, rp, ci, dp, vals = case
    ar, ac = (C.c_int32 * len(rp))(*rp), (C.c_int32 * len(ci))(*ci)
    ad, av = (C.c_int32 * len(dp))(*dp), (C.c_double * len(vals))(*vals)
    ilu = (C.c_double * len(vals))(*([123.25] * len(vals)))
    fail, minimum = (C.c_int32 * 2)(-77, -77), C.c_double(-77.0)
    def call():
        return int(f(n, len(ci), ar, ac, ad, av, ilu, fail, C.byref(minimum)))
    return call


def run(f, case):
    n, rp, ci, dp, vals = case
    ar = (C.c_int32 * len(rp))(*rp)
    ac = (C.c_int32 * len(ci))(*ci)
    ad = (C.c_int32 * len(dp))(*dp)
    av = (C.c_double * len(vals))(*vals)
    ilu = (C.c_double * len(vals))(*([123.25] * len(vals)))
    fail = (C.c_int32 * 2)(-77, -77)
    minimum = C.c_double(-77.0)
    rc = int(f(n, len(ci), ar, ac, ad, av, ilu, fail, C.byref(minimum)))
    return rc, bytes(ilu), bytes(fail), struct.pack("=d", minimum.value), minimum.value


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--reference", type=pathlib.Path, required=True)
    ap.add_argument("--candidate", type=pathlib.Path, required=True)
    ap.add_argument("--count", type=int, default=100)
    ap.add_argument("--out", type=pathlib.Path)
    ap.add_argument("--benchmark", action="store_true")
    args = ap.parse_args()
    # Keep the same CRT dependency ordering as the VBA host.
    C.WinDLL(str(pathlib.Path('C:/Windows/System32/msvcr100.dll')))
    if C.sizeof(C.c_double) != 8:
        raise SystemExit("binary64 required")
    ref, rf = load(args.reference)
    cand, cf = load(args.candidate)
    rows = []
    words=(0x027f,0x037f) if hasattr(ref,'sfm_probe_cw') and hasattr(cand,'sfm_probe_cw') else (None,)
    saved=None
    try:
        for cw in words:
            if cw is not None:
                old=ref.sfm_probe_cw(cw)
                if saved is None:saved=old
            for idx, case in enumerate(cases(args.count)):
                a, b = run(rf, case), run(cf, case)
                if a[:4] != b[:4]:
                    raise AssertionError("cw %s case %d differs: rc/fail/min/ilu" % (cw,idx))
                rows.append({"index": idx, "n": case[0], "nnz": len(case[2]), "rc": a[0]})
    finally:
        if saved is not None:ref.sfm_probe_cw(saved)
    bench = None
    if args.benchmark:
        bench = {}
        for n, bw in ((169, 8), (676, 20)):
            c = make_case(n, bw, random.Random(4 + n), False, None)
            funcs = [bind(rf, c), bind(cf, c)]
            times = [[], []]
            for rep in range(5):
                for side in (0, 1) if rep % 2 == 0 else (1, 0):
                    t = time.perf_counter()
                    for _ in range(100):
                        assert funcs[side]() == 1
                    times[side].append((time.perf_counter() - t) / 100.0)
            med = [statistics.median(x) for x in times]
            bench[str(n)] = {"n": n, "nnz": len(c[2]), "calls": 100,
                             "reference_samples_s": times[0], "candidate_samples_s": times[1],
                             "reference_median_s": med[0], "candidate_median_s": med[1],
                             "reduction_pct": (med[0] - med[1]) / med[0] * 100.0}
    report = {"reference": str(args.reference), "candidate": str(args.candidate),
              "reference_sha256": sha(args.reference), "candidate_sha256": sha(args.candidate),
              "seed": "0x1A", "cases": len(rows), "matches": len(rows),
              "x87_control_words":words,"fp_control_restored":True,"benchmark": bench,
              "return_code_counts":{str(rc):sum(row['rc']==rc for row in rows) for rc in {row['rc'] for row in rows}}}
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
