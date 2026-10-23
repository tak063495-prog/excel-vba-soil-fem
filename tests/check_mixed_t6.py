"""Check actual Excel results against analytic fields, independent of VBA assembly."""
import argparse
import json
from pathlib import Path

import numpy as np


def read(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def check(root, patch_only=False):
    root = Path(root)
    report = {"cases": {}, "comparisons": {}, "scope": "patches" if patch_only else "full"}
    if not patch_only:
        for name in ("mixed_plastic_auto", "mixed_plastic_off", "mixed_joint_water_wet", "mixed_joint_water_stages", "mixed_slope_srm"):
            assert (root/name/"result.json").is_file(), ("missing required case",name)
        for name in ("mixed_joint_water_wet", "mixed_joint_water_stages"):
            data=read(root/name/"result.json");resume=read(root/name/"resume.json")
            assert data["status"].startswith("PASS|") and resume["status"].startswith("PASS|")
            assert data["metrics"][3] <= 1e-7 and resume["metrics"][3] <= 1e-7
            assert abs(data["metrics"][0]-88) <= 1e-8 and abs(data["metrics"][2]-88) <= 1e-6
            du=float(np.max(np.abs(np.array(data["disp"])-resume["disp"])))
            assert du <= 1e-9
            report["comparisons"][name+"_resume"]=dict(max_du_m=du,residual=data["metrics"][3],seconds=data["seconds"],reaction_y=data["metrics"][2])
    modulus = 10000 * (1 - .3) / ((1 + .3) * (1 - 2 * .3))
    for kind in ("q8", "mixed", "t6"):
        for load in ("gravity", "traction", "seepage"):
            name = f"{kind}_{load}"
            folder = root / name
            if name == "mixed_gravity" and not folder.exists():
                folder = root / "mixed_gravity01"
            data = read(folder / "result.json")
            assert data["status"].startswith("PASS|"), (name, data["status"])
            assert data["metrics"][3] <= 1e-7, (name, "mechanics residual")
            assert np.isfinite(data["disp"]).all()
            assert np.isfinite(data["stress"]).all()
            rows = read(folder / "結果節点.json")[1:]
            rows = [r for r in rows if isinstance(r[0], (int, float))]
            final = max(r[0] for r in rows)
            rows = [r for r in rows if r[0] == final]
            entry = {"seconds": data["seconds"], "residual": data["metrics"][3],
                     "reaction_y": data["metrics"][2], "nodes": len(rows)}
            if load != "seepage":
                error = 0.
                for row in rows:
                    y = row[3]
                    expected = (-18 * (y - y*y/2) if load == "gravity" else -100*y) / modulus
                    error = max(error, abs(row[4]), abs(row[5] - expected))
                assert error <= 1e-9, (name, "analytic displacement", error)
                expected_reaction = 36 if load == "gravity" else 200
                assert abs(data["metrics"][2] - expected_reaction) <= 1e-6
                entry["max_displacement_error_m"] = error
            else:
                water = read(folder / "水圧結果.json")
                heads = [r for r in water[35:] if len(r) > 38 and isinstance(r[33], (int, float))]
                assert heads, (name, "missing head output")
                head_error = max(abs(r[33] - (3 - .5*r[37])) for r in heads)
                points = [r for r in water[35:] if isinstance(r[0], (int, float)) and isinstance(r[5], (int, float))]
                pressure_error = max(abs(r[5] - 9.81*(3 - .5*r[3] - r[4])) for r in points)
                flux_error = max(max(abs(r[12] - 5e-5), abs(r[13])) for r in points)
                assert head_error <= 1e-8, (name, "head", head_error)
                assert pressure_error <= 1e-7, (name, "pressure", pressure_error)
                assert flux_error <= 1e-10, (name, "flux", flux_error)
                for row in (1, 2, 5):
                    assert abs(water[row][17]) <= 1e-8, (name, "water residual/balance", row, water[row][17])
                entry.update(head_error_m=head_error, pressure_error_kpa=pressure_error, flux_error_m_s=flux_error)
            report["cases"][name] = entry
    off = read(root / "mixed_gravity_off" / "result.json")
    auto_path = root / "mixed_gravity" / "result.json"
    if not auto_path.exists():
        auto_path = root / "mixed_gravity01" / "result.json"
    auto = read(auto_path)
    du = float(np.max(np.abs(np.array(off["disp"]) - auto["disp"])))
    # Principal directions are undefined at coincident eigenvalues.
    stress_a = np.array(auto["stress"]).reshape(-1, 17)
    stress_b = np.array(off["stress"]).reshape(-1, 17)
    stress_error = float(np.max(np.abs(np.delete(stress_a-stress_b, 6, axis=1))))
    assert du <= 1e-10 and stress_error <= 1e-7, ("native parity", du, stress_error)
    report["comparisons"]["native_off_auto"] = dict(max_du_m=du, max_stress_delta=stress_error)
    plastic_auto = root / "mixed_plastic_auto" / "result.json"
    if plastic_auto.exists():
        pa = read(plastic_auto)
        pb = read(root / "mixed_plastic_off" / "result.json")
        assert pa["status"].startswith("PASS|") and pb["status"].startswith("PASS|")
        assert pa["metrics"][3] <= 1e-7 and pb["metrics"][3] <= 1e-7
        du = float(np.max(np.abs(np.array(pa["disp"])-pb["disp"])))
        ds = float(np.max(np.abs(np.delete(np.array(pa["stress"]).reshape(-1,17)-np.array(pb["stress"]).reshape(-1,17),6,axis=1))))
        assert du <= 1e-9 and ds <= 1e-6, ("plastic native parity",du,ds)
        plastic = read(root / "mixed_plastic_auto" / "結果塑性.json")[1:]
        rows = [r for r in plastic if isinstance(r[0],(int,float))]
        multiplier = max(abs(float(r[9])) for r in rows)
        assert multiplier > 1e-10, "Dry native case did not exercise plasticity"
        assert "eval_elements=0;" not in pa["native"] and "eval_elements=" in pa["native"]
        report["comparisons"]["dry_plastic_native_off_auto"] = dict(max_du_m=du,max_stress_delta=ds,max_plastic_multiplier=multiplier,auto_seconds=pa["seconds"],off_seconds=pb["seconds"])
    srm = root / "mixed_slope_srm" / "result.json"
    if srm.exists() and not patch_only:
        data = read(srm)
        resume = read(srm.parent / "resume.json")
        assert data["status"].startswith("PASS|") and resume["status"].startswith("PASS|")
        assert data["metrics"][5] > 0 and data["metrics"][6] >= data["metrics"][5]
        du = float(np.max(np.abs(np.array(data["disp"]) - resume["disp"])))
        assert du <= 1e-9
        assert np.allclose(data["metrics"][4:8], resume["metrics"][4:8], atol=1e-10, rtol=0)
        report["comparisons"]["mixed_srm_resume"] = dict(fs_pass=data["metrics"][5], fs_fail=data["metrics"][6], max_du_m=du, seconds=data["seconds"])
    baseline = root / "q8_reference_srm" / "result.json"
    if baseline.exists() and not patch_only:
        old = read(baseline)
        new = read(root / "q8_candidate_srm" / "result.json")
        assert old["status"].startswith("PASS|") and new["status"].startswith("PASS|")
        du = float(np.max(np.abs(np.array(old["disp"]) - new["disp"])))
        stress_old = np.array(old["stress"]).reshape(-1, 17)
        stress_new = np.array(new["stress"]).reshape(-1, 17)
        ds = float(np.max(np.abs(np.delete(stress_old-stress_new, 6, axis=1))))
        dm = float(np.max(np.abs(np.array(old["metrics"]) - new["metrics"])))
        assert du <= 1e-9 and ds <= 1e-6 and dm <= 1e-7, ("Q8 regression", du, ds, dm)
        assert np.allclose(old["metrics"][4:8], new["metrics"][4:8], atol=1e-10, rtol=0)
        report["comparisons"]["q8_original_candidate_srm"] = dict(
            max_du_m=du, max_stress_delta=ds, max_metric_delta=dm,
            fs_pass=new["metrics"][5], fs_fail=new["metrics"][6],
            original_seconds=old["seconds"], candidate_seconds=new["seconds"])
    if baseline.exists() and srm.exists() and not patch_only:
        def final_nodes(folder):
            rows = [r for r in read(root/folder/"結果節点.json")[1:] if isinstance(r[0],(int,float))]
            stage = max(r[0] for r in rows)
            return {int(r[1]): np.array(r[2:6],dtype=float) for r in rows if r[0] == stage}
        original_nodes = final_nodes("q8_reference_srm")
        mixed_nodes = final_nodes("mixed_slope_srm")
        assert original_nodes.keys() <= mixed_nodes.keys()
        common = sorted(original_nodes)
        assert all(np.max(np.abs(original_nodes[n][:2]-mixed_nodes[n][:2])) <= 1e-8 for n in common)
        delta = max(float(np.linalg.norm(original_nodes[n][2:]-mixed_nodes[n][2:])) for n in common)
        report["comparisons"]["q8_mixed_srm_common_nodes"] = dict(
            common_nodes=len(common), max_displacement_vector_difference_m=delta,
            q8_max_displacement_m=max(float(np.linalg.norm(v[2:])) for v in original_nodes.values()),
            mixed_max_displacement_m=max(float(np.linalg.norm(v[2:])) for v in mixed_nodes.values()),
            note="Different meshes; this is a model comparison, not an equality test.")
    report["status"] = "PASS"
    return report


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("root", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--patch-only", action="store_true", help="Check only completed analytic patches; excludes SRM/staged acceptance")
    args = parser.parse_args()
    report = check(args.root, args.patch_only)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.out:
        args.out.write_text(text + "\n", encoding="utf-8")
    print(text)
