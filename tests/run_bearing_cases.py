"""Run gravity-referenced footing displacement examples in isolated Excel copies.

Run from the repository root. No source workbook or VBA is modified.
Each trial starts from the original input and retains the solver tolerances.
"""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import time


def read(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def run_fixture(fixture, folder, flow, log, stage_limit=999):
    result = folder / "result.json"
    if result.exists():
        data = read(result)
        assert data.get("application_code_unchanged"), folder
        return data
    if folder.exists():
        raise RuntimeError(f"Incomplete output exists; inspect before restarting: {folder}")
    command = ["pwsh", "-NoProfile", "-File", "tests/run_manual_case.ps1",
               "-Fixture", str(fixture), "-OutDirectory", str(folder), "-Flow", flow, "-StageLimit", str(stage_limit)]
    with log.open("w", encoding="utf8") as stream:
        subprocess.run(command, stdout=stream, stderr=subprocess.STDOUT, check=True)
    time.sleep(3)  # Allow the owned Excel COM server to finish its natural exit.
    data = read(result)
    assert data.get("application_code_unchanged"), folder
    print(folder.name, data["status"], data["seconds"], flush=True)
    return data


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--flow", choices=["INCONSISTENT", "DAVIS"], default="INCONSISTENT")
    parser.add_argument("--coarse-settlements", nargs="+", type=float, default=[.001, .005, .01, .02, .05])
    parser.add_argument("--fine-settlements", nargs="+", type=float, default=[.01, .02, .05])
    args = parser.parse_args()
    out = args.out.resolve()
    out.mkdir(parents=True, exist_ok=True)
    inputs = out / "inputs"
    inputs.mkdir(exist_ok=True)
    plan = {"flow": args.flow, "coarse_settlements_m": args.coarse_settlements,
            "fine_settlements_m": args.fine_settlements, "fixture_sha256": {}}
    for terrain in ("flat", "slope"):
        for fine in (False, True):
            name = f"bearing_{terrain}{'_fine' if fine else ''}_mesh.json"
            path = Path("tests/manual_cases") / name
            plan["fixture_sha256"][name] = hashlib.sha256(path.read_bytes()).hexdigest()
    plan_path = out / "plan.json"
    if plan_path.exists():
        assert read(plan_path) == plan, "Existing queue plan differs"
    else:
        plan_path.write_text(json.dumps(plan, indent=2), encoding="utf8")
    records = []
    for fine in (False, True):
        for terrain in ("flat", "slope"):
            tag = f"bearing_{terrain}{'_fine' if fine else ''}"
            fixture = Path("tests/manual_cases") / f"{tag}_mesh.json"
            template = read(fixture)
            assert not template["stages"][1][4], "Template must disable LOAD until gravity reference exists"
            baseline_folder = out / f"{tag}_gravity"
            baseline = run_fixture(fixture, baseline_folder, args.flow, out / f"{tag}_gravity.log", stage_limit=1)
            records.append({"case": baseline_folder.name, "status": baseline["status"], "seconds": baseline["seconds"]})
            if baseline["status"] != "PASS":
                raise RuntimeError(f"Gravity reference did not converge: {tag}")
            reference = {int(r[1]): r for r in baseline["tables"]["結果節点"][1:] if r[0] == 1}
            foot = [int(n[0]) for n in template["nodes"]
                    if 1 - 1e-9 <= n[1] <= 3 + 1e-9 and abs(n[2] - 5) < 1e-9]
            assert len(foot) >= 3
            settlements = args.fine_settlements if fine else args.coarse_settlements
            for settlement in settlements:
                assert 0 < settlement <= .1, "Settlement is an increment in metres"
                trial = json.loads(json.dumps(template))
                trial["stages"][1][4] = 1
                trial["loads"] = [[1, node, "Y", "DISP", reference[node][5] - settlement,
                                    None, None, None, None, "Gravity-referenced vertical settlement"]
                                   for node in foot]
                trial["bearing_protocol"] = {"baseline": baseline_folder.name,
                    "baseline_result_sha256": hashlib.sha256((baseline_folder / "result.json").read_bytes()).hexdigest(),
                    "additional_settlement_m": settlement, "width_m": 2, "thickness_m": 1,
                    "footing_nodes": foot,
                    "reference_uy_m": {str(node): reference[node][5] for node in foot},
                    "scope": "Uniform incremental vertical displacement of a soil loading strip; no rigid footing/contact/separation element"}
                case = f"{tag}_s{round(settlement * 1e6):06d}"
                trial_path = inputs / f"{case}.json"
                payload = json.dumps(trial, ensure_ascii=False, indent=2) + "\n"
                if trial_path.exists():
                    assert trial_path.read_text(encoding="utf8") == payload
                else:
                    trial_path.write_text(payload, encoding="utf8")
                data = run_fixture(trial_path, out / case, args.flow, out / f"{case}.log")
                records.append({"case": case, "status": data["status"], "seconds": data["seconds"],
                                "settlement_m": settlement, "terrain": terrain, "fine": fine})
                (out / "queue_results.json").write_text(json.dumps(records, indent=2), encoding="utf8")
    print("QUEUE COMPLETE", flush=True)


if __name__ == "__main__":
    main()
