"""Check retained numerical and UI records; does not start Excel."""
from __future__ import annotations
import argparse
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def flat(x):
    if isinstance(x, list):
        return [v for part in x for v in flat(part)]
    return [float(x)]


def compare(a, b):
    result = {}
    for key in ("disp", "stress", "metrics"):
        av, bv = flat(a[key]), flat(b[key])
        assert av and len(av) == len(bv), key
        result[key] = max(abs(x-y) for x, y in zip(av, bv))
        assert result[key] == 0.0, (key, result[key])
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", type=Path, required=True)
    ap.add_argument("--book", type=Path, required=True)
    ap.add_argument("--generation", type=int, default=15)
    ap.add_argument("--baseline", type=Path, default=Path("tests/tmp/groundwater/acceptance20"))
    ap.add_argument("--davis-baseline", type=Path, default=Path("tests/tmp/water_display/final7/wet_srm_DAVIS/result.json"))
    ap.add_argument("--base", default="ecc05e794847601267437f627d7a6249b3290bd0")
    args = ap.parse_args()
    root = args.root
    g = str(args.generation)
    book_hash = hashlib.sha256(args.book.read_bytes()).hexdigest().upper()
    package = load(root / f"package{g}.json")
    assert package["status"] == "PASS" and len(package["modules"]) == 35
    assert package["package_sha256"] == book_hash
    ui = load(root / f"ui{g}/checks.json")
    assert ui["status"] == "PASS" and ui["book_sha256"] == book_hash
    assert len(ui["checks"]) >= 60
    assert all(x["status"] == "PASS" for x in ui["checks"])
    assert "ALL PASS" in (root / f"regression{g}/queue.log").read_text()
    subprocess.run([sys.executable, "tests/check_water_display.py", "--root", str(root / f"regression{g}"),
                    "--baseline", str(args.baseline), "--davis-baseline", str(args.davis_baseline)], check=True, stdout=subprocess.DEVNULL)
    water = load(root / f"regression{g}/water_display_check.json")
    front = load(root / f"front{g}/result.json")
    disabled = load(root / f"disabled{g}/result.json")
    assert load(root / f"front{g}/front_inputs.json")["status"] == "PASS"
    for case in (front, disabled, load(root / f"disabled_ref{g}/result.json")):
        assert case["status"].startswith("PASS")
    report = {"status": "PASS", "book_sha256": book_hash, "ui_check_count": len(ui["checks"]),
              "frontend_vs_canonical": compare(front, load(root / f"regression{g}/one_way/result.json")),
              "disabled_water_plan_vs_no_plan": compare(disabled, load(root / f"disabled_ref{g}/result.json")),
              "water_regression": water, "unchanged_sources": {}, "base": args.base}
    basic = load(root / f"regression{g}/basic/results.json")
    basic_ref = {x["case"]: x for x in load(args.baseline / "basic/results.json")}
    report["basic"] = []
    for case in basic:
        if "disp" not in case:
            assert case["result"].startswith("PASS"), case
            report["basic"].append(case)
            continue
        ref = basic_ref[case["case"]]
        assert case["status"].startswith("PASS")
        assert case["disp"] == ref["disp"] and case["bracket"] == ref["bracket"]
        report["basic"].append({"case": case["case"], "status": "PASS", "displacement_count": len(case["disp"]),
                                "maxdiff": 0.0, "candidate_seconds": case["seconds"], "reference_seconds": ref["seconds"]})
    unchanged = ["FEMCore", "FEMEngine", "FEMSolver", "FEMNative", "FEMNativeKern", "FEMNativeMat",
                 "FEMSeepage", "FEMSeepageFlux", "FEMWaterGeometry", "FEMWaterView", "FEMJointWater",
                 "FEMStrainMath", "FEMViewIncrement", "FEMIncrementSvg", "Mesh", "MeshAdapt", "FEMTerrainBlocks"]
    for name in unchanged:
        path = f"src/vba/{name}.bas"
        old = subprocess.check_output(["git", "show", f"{args.base}:{path}"])
        new = Path(path).read_bytes()
        assert old == new, f"Unexpected calculation/mesh source change: {name}"
        report["unchanged_sources"][name] = hashlib.sha256(new).hexdigest()
    path = "src/vba/FEMWater.bas"
    old = subprocess.check_output(["git", "show", f"{args.base}:{path}"]).decode("utf-8-sig").replace("\r\n", "\n")
    new = Path(path).read_text(encoding="utf-8-sig")
    pattern = r"Public Sub GWValidateStagePlans\(\).*?\nEnd Sub"
    assert re.sub(pattern, "", old, flags=re.S) == re.sub(pattern, "", new, flags=re.S)
    report["water_change_scope"] = "GWValidateStagePlans only; inactive stage plans retained and skipped"
    dll = args.book.parent / "soilfem_native_x64.dll"
    old_dll = subprocess.check_output(["git", "show", f"{args.base}:workbook/ui_20261023/soilfem_native_x64.dll"])
    assert dll.read_bytes() == old_dll
    report["unchanged_dll_sha256"] = hashlib.sha256(old_dll).hexdigest()
    out = root / f"stage_ui24_check{g}.json"
    out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in report.items() if k != "water_regression"}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
