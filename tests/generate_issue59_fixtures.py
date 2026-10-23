"""Generate runnable one-element Excel fixtures from the existing Q8 gravity case."""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / "tests/fixtures/mixed_t6/q8_gravity.json"
CASES = ROOT / "tests/fixtures/issue59/q8_jacobian_cases.json"
OUT = ROOT / "tests/fixtures/issue59/excel"

def main():
    base = json.loads(BASE.read_text(encoding="utf-8-sig"))
    cases = json.loads(CASES.read_text(encoding="utf-8"))["cases"]
    OUT.mkdir(parents=True, exist_ok=True)
    for case in cases:
        d = json.loads(json.dumps(base))
        d["name"] = "issue59_" + case["name"]
        coords = case["coordinates"]
        nodes = []
        for i, (x, y) in enumerate(coords, 1):
            # Q8 gravity nodes use [id,x,y,fixUx,fixUy,...]. Preserve the
            # original supports by copying the corresponding canonical row.
            row = list(base["nodes"][i - 1])
            row[1], row[2] = x, y
            nodes.append(row)
        d["nodes"] = nodes
        d["elements"] = [list(base["elements"][0])]
        d["stages"] = [[1, "GRAVITY", "", 1, 1, "gravity"]]
        if case['name'] in {'internal_inversion', 'random_internal', 'degenerate', 'near_degenerate', 'square_small'}:
            d['expectedAnalysisError'] = '要素=1'
        d['checkNoCommit'] = bool(d.get('expectedAnalysisError'))
        (OUT / (d["name"] + ".json")).write_text(json.dumps(d, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"generated {len(cases)} issue59 Excel fixtures in {OUT}")

if __name__ == "__main__":
    main()
