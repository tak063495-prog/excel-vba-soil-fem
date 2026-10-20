"""Exact calculation comparison for a UI-only workbook change."""
import argparse
import json
from pathlib import Path


def records(path):
    value = json.loads(path.read_text(encoding="utf-8-sig"))
    return value if isinstance(value, list) else [value]


def compare(before, after):
    old, new = records(before), records(after)
    assert len(old) == len(new) and old, "run count differs"
    result = []
    fixed_keys = ("Case", "Variant", "Mode", "Precision", "StateCopy", "GmresSimd", "IoFast",
                  "EventFast", "AssemblyFast", "Flow", "GmresTuned", "GmresPrepared",
                  "NativeDllSha256", "MaterialDllSha256")
    for a, b in zip(old, new):
        assert a["ExclusiveExcel"] and b["ExclusiveExcel"], "timing has another Excel process"
        assert a.get("UiPrewarmed", False) == b.get("UiPrewarmed", False), "different UI initialization timing scope"
        for key in fixed_keys:
            assert a[key] == b[key], f"configuration changed: {key}"
        for key in ("Bracket", "Disp", "Fos", "Assembly", "TunedCalls", "PreparedCalls",
                    "SolverEventsSha256", "SolverEventsBytes", "DecisionEventsSha256"):
            assert a[key] == b[key], f"calculation differs: {a['Case']} {a['Flow']} {key}"
        ar, br = a["Run"].split("|"), b["Run"].split("|")
        assert ar[1:9] == br[1:9] and ar[13:] == br[13:], "status/residual/solver counters differ"
        assert br[1] == "PASS" and br[2] == "True", "analysis did not pass"
        assert len(a["Disp"]) == len(b["Disp"]) > 0
        delta = max(abs(x - y) for x, y in zip(a["Disp"], b["Disp"]))
        assert delta == 0
        result.append(dict(case=b["Case"], flow=b["Flow"], status="PASS", bracket=b["Bracket"],
                           displacement_count=len(b["Disp"]), max_abs_displacement_delta=delta,
                           final_relative_residual=float(br[3]),
                           baseline_seconds=a["WallSeconds"], candidate_seconds=b["WallSeconds"],
                           full_solver_csv_sha256=b["SolverEventsSha256"],
                           full_solver_csv_bytes=b["SolverEventsBytes"], full_solver_csv_equal=True))
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--pair", action="append", nargs=2, type=Path, required=True,
                        metavar=("BASELINE_JSON", "CANDIDATE_JSON"))
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    results = [row for pair in args.pair for row in compare(*pair)]
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(dict(status="PASS", comparisons=len(results), runs=results),
                                   ensure_ascii=False, indent=2) + "\n", encoding="utf8")
    print(f"PASS {len(results)} exact comparisons: Fs, all displacements, residuals, full event CSV")
