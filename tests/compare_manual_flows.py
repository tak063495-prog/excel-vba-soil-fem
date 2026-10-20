"""Compare two completed manual-case result.json files without Excel or plotting."""
import argparse
import csv
import json
import sys
from pathlib import Path

import numpy as np

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from check_manual_case import check as check_manual_case, state  # noqa: E402


def _load(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def _settings(data):
    return {row[4]: row[2] for row in data["tables"]["設定"][1:] if len(row) > 4 and row[4]}


def _diagnostics(data):
    return {row[0]: row[1] for row in data["tables"]["診断"][1:] if row and row[0]}


def _positive(diag, key):
    value = diag.get(key)
    return float(value) if isinstance(value, (int, float)) and value > 0 else None


def _run_record(path, data, geometry):
    diag = _diagnostics(data)
    interpretation = diag.get("FOS_INTERPRETATION")
    fail = None if interpretation == "LOWER_BOUND" else _positive(diag, "P3SrmUpper")
    return {
        "result": str(Path(path)),
        "status": data["status"],
        "seconds": data.get("seconds"),
        "Fs_PASS": _positive(diag, "P3SrmLower"),
        "Fs_FAIL": fail,
        "interpretation": interpretation,
        "note": diag.get("P3SrmNote"),
        "relative_residual": diag.get("RelativeResidualFree"),
        "geometry": geometry,
        "geometry_valid": bool(geometry.get("deformed_geometry_valid")),
    }


def _write_displacements(out, label, inco, davis):
    _, nodes_i, ui, _, _, _ = state(inco, 3 if label == "construction_Fs1" else 4)
    _, nodes_d, ud, _, _, _ = state(davis, 3 if label == "construction_Fs1" else 4)
    assert ui and ud, (label, "missing saved stage")
    if label == "construction_Fs1":
        for data in (inco, davis):
            inc = [r for r in data["tables"]["結果増分"][1:] if r[0] == 3]
            assert inc and all(abs(float(r[17]) - 1) < 1e-12 for r in inc), "construction state must be Fs=1"
    assert ui.keys() == ud.keys(), (label, "node sets differ")
    rows = []
    all_delta = []
    max_i = max_d = 0.0
    for node in sorted(ui):
        vi, vd = np.asarray(ui[node], float), np.asarray(ud[node], float)
        delta = vd - vi
        xy = nodes_i[node][1:3]
        rows.append([node, *xy, *vi, *vd, *delta])
        all_delta.extend(delta.tolist())
        max_i = max(max_i, float(np.linalg.norm(vi)))
        max_d = max(max_d, float(np.linalg.norm(vd)))
    csv_path = out / f"{label}_all_displacements.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["node", "x_m", "y_m", "ux_INCO_m", "uy_INCO_m",
                         "ux_DAVIS_m", "uy_DAVIS_m", "delta_ux_m", "delta_uy_m"])
        writer.writerows(rows)
    delta = np.asarray(all_delta, float)
    return {
        "nodes": len(ui),
        "components": 2 * len(ui),
        "csv": str(csv_path),
        "max_abs_component_delta_m": float(np.max(np.abs(delta))),
        "l2_delta_m": float(np.linalg.norm(delta)),
        "max_norm_INCO_m": max_i,
        "max_norm_DAVIS_m": max_d,
    }


def compare(inco_path, davis_path, out_path):
    inco, davis = _load(inco_path), _load(davis_path)
    for data in (inco, davis):
        stages = {int(r[0]): str(r[1]).upper() for r in data["tables"]["ステージ"][1:] if isinstance(r[0], (int, float)) and r[4]}
        assert stages == {1: "GRAVITY", 2: "BIRTH", 3: "BIRTH", 4: "SRM"}, "this comparator requires the published wall four-stage fixture"
    assert inco["status"] == "PASS" and davis["status"] == "PASS", "both results must PASS"
    for table in ("材料データ", "節点データ", "要素データ", "ステージ", "載荷"):
        assert inco["tables"][table] == davis["tables"][table], table
    si, sd = _settings(inco), _settings(davis)
    keys = si.keys() | sd.keys()
    diff = {key: [si.get(key), sd.get(key)] for key in keys if si.get(key) != sd.get(key)}
    assert set(diff) == {"FLOW_POLICY"}, diff
    assert diff["FLOW_POLICY"] == ["INCONSISTENT", "DAVIS"], diff

    out = Path(out_path)
    out.mkdir(parents=True, exist_ok=True)
    warnings = []
    geometries = {}
    for name, path in (("INCONSISTENT", inco_path), ("DAVIS", davis_path)):
        # Report shape inversion, but keep all other independent checks strict.
        geometry = check_manual_case(str(Path(path).resolve().parent), strict_deformed=False)
        geometries[name] = geometry
        if not geometry.get("deformed_geometry_valid", False):
            warnings.append(f"WARNING: {name} deformed geometry invalid; not a geometry PASS")

    record = {
        "same_mesh_material_stages_loads": True,
        "setting_differences": diff,
        "runs": [_run_record(inco_path, inco, geometries["INCONSISTENT"]),
                  _run_record(davis_path, davis, geometries["DAVIS"])],
        "warnings": warnings,
        "displacements": {},
    }
    for label in ("construction_Fs1", "final_saved_SRM"):
        item = _write_displacements(out, label, inco, davis)
        if item is not None:
            record["displacements"][label] = item
    (out / "flow_comparison.json").write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(record, ensure_ascii=False, indent=2))
    for warning in warnings:
        print(warning)
    return record


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Compare INCONSISTENT and DAVIS manual results")
    parser.add_argument("--inco", required=True, type=Path, help="INCONSISTENT result.json")
    parser.add_argument("--davis", required=True, type=Path, help="DAVIS result.json")
    parser.add_argument("--out", required=True, type=Path, help="output directory")
    args = parser.parse_args()
    compare(args.inco, args.davis, args.out)
