"""Independent force reconstruction for Issue 51 saved fixed-stress frames.

This deliberately reads only the persisted frame CSVs and the native result
sheet dump.  It does not use the VBA force arrays as an oracle.
"""
import argparse
import csv
import json
import sys
from collections import defaultdict
from pathlib import Path

import numpy as np

from test_fixed_stress_reference import grad, pressure_shape, q8_shape, reference_local


TOL = 1.0e-7
AFFINE_TOL = 1.0e-9
GP = 1.0 / np.sqrt(3.0)
GP_NATURAL = ((-GP, -GP), (GP, -GP), (GP, GP), (-GP, GP))


def fail(message):
    raise ValueError(message)


def rows_after_header(path, header):
    lines = path.read_text(encoding="utf-8-sig").splitlines()
    try:
        start = lines.index(header)
    except ValueError:
        fail(f"{path}: missing {header!r}")
    return list(csv.DictReader(lines[start:]))


def fixture_model(path):
    fixture = json.loads(path.read_text(encoding="utf-8-sig"))
    nodes = {int(r[0]): (float(r[1]), float(r[2]), int(r[3]), int(r[4]))
             for r in fixture["nodes"]}
    elements = {int(r[0]): [int(v) for v in r[1:9]] for r in fixture["elements"]}
    material = fixture["materials"][0]
    coupled = fixture["water"]["coupledMaterials"][0]
    water = fixture["water"]
    return fixture, nodes, elements, material, coupled, water


def read_dof(path, count):
    values = np.zeros(count)
    for r in csv.DictReader(path.open(encoding="utf-8-sig", newline="")):
        i = int(r["dof"])
        if i < 0 or i >= count:
            fail(f"{path}: original dof {i} outside fixture")
        values[i] = float(r["locked_sw"]) + float(r["locked_ap"])
    if len(values) != count:
        fail(f"{path}: malformed DOF table")
    return values


def read_stress(path):
    out = {}
    for r in csv.DictReader(path.open(encoding="utf-8-sig", newline="")):
        key = (int(r["elem"]), int(r["gp"]))
        out[key] = np.array([float(r["sx"]), float(r["sy"]), float(r["txy"])])
    return out


def read_water(path):
    return {(int(r["element"]), int(r["gauss"])): float(r["pressure_kPa"])
            for r in rows_after_header(path, "element,gauss,pressure_kPa")}


def element_B(coords, xi, eta):
    dxi = np.array([grad(lambda a, b, i=i: q8_shape(a, b)[i], xi, eta)[0]
                    for i in range(8)])
    deta = np.array([grad(lambda a, b, i=i: q8_shape(a, b)[i], xi, eta)[1]
                     for i in range(8)])
    jac = np.array([[dxi @ coords[:, 0], deta @ coords[:, 0]],
                    [dxi @ coords[:, 1], deta @ coords[:, 1]]])
    det = np.linalg.det(jac)
    if det <= 0:
        fail("non-positive Q8 Jacobian")
    dxy = np.linalg.inv(jac).T @ np.vstack((dxi, deta))
    b = np.zeros((3, 16))
    b[0, 0::2] = dxy[0]
    b[1, 1::2] = dxy[1]
    b[2, 0::2] = dxy[1]
    b[2, 1::2] = dxy[0]
    return b, det


def corner_pressures(water, element):
    try:
        p = np.array([water[(element, i + 1)] for i in range(4)])
    except KeyError as exc:
        fail(f"water table lacks four GP values for element {element}: {exc}")
    n = np.array([pressure_shape("q8", xi, eta) for xi, eta in GP_NATURAL])
    return np.linalg.solve(n, p)


def reconstruct_internal(nodes, elements, material, coupled, stress, water):
    """Return -B^T sigma' (2x2) - Q p_corner in original DOF order."""
    count = 2 * len(nodes)
    internal = np.zeros(count)
    alpha = float(coupled[1])
    storage, kx, ky = map(float, coupled[2:5])
    E, nu = float(material[1]), float(material[2])
    gamma = float(material[4])  # only passed to reference Q; Q itself is gamma-free
    for eid, enodes in elements.items():
        coords = np.array([(nodes[n][0], nodes[n][1]) for n in enodes])
        _, _, _, q, _ = reference_local(coords, "q8", alpha, storage, kx, ky,
                                        gamma, E, nu)
        local = -q @ corner_pressures(water, eid)
        for gp, (xi, eta) in enumerate(GP_NATURAL):
            try:
                total = stress[(eid, gp)]
            except KeyError:
                fail(f"gauss table lacks element {eid}, GP {gp}")
            p = float(pressure_shape("q8", xi, eta) @ corner_pressures(water, eid))
            effective = total - np.array([alpha * p, alpha * p, 0.0])
            b, det = element_B(coords, xi, eta)
            local += -b.T @ effective * det
        for local_node, nid in enumerate(enodes):
            internal[2 * (nid - 1):2 * nid] += local[2 * local_node:2 * local_node + 2]
    return internal


def numeric(value, label):
    if value is None or value == "":
        fail(f"missing {label}")
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"invalid {label}: {value!r}") from exc


def sheet_field(record, nodes):
    rows = record["sheet_rows"]
    if not rows or len(rows[0]) < 10:
        fail(f"{record['id']}: result sheet has no A:J header")
    values = {}
    for raw in rows[1:]:
        if not raw or raw[1] in (None, ""):
            continue
        nid = int(numeric(raw[1], "node ID"))
        if nid not in nodes:
            fail(f"{record['id']}: unknown node {nid} in result sheet")
        if nid in values:
            fail(f"{record['id']}: duplicate node {nid} in result sheet")
        values[nid] = np.array([numeric(raw[i], f"column {i + 1}") for i in range(2, 10)])
    if set(values) != set(nodes):
        fail(f"{record['id']}: expected {len(nodes)} nodes, found {len(values)}")
    return values


def assert_one_dimensional_field(sheet):
    # The vertical displacement need not be affine in y.  The relevant zero
    # hourglass property for this Q8 column is the exact one-dimensional field:
    # ux=0 and every node at a common y has the same uy.
    ux = max(abs(v[2]) for v in sheet.values())
    by_y = defaultdict(list)
    for v in sheet.values():
        by_y[round(v[1], 12)].append(v[3])
    uy_spread = max(max(v) - min(v) for v in by_y.values())
    worst = max(ux, uy_spread)
    if worst > AFFINE_TOL:
        fail(f"field is not one-dimensional (hourglass-free premise fails): {worst:.3e} m")
    return worst


def adopted_head(frame, water_cfg):
    """Recover the imposed free-surface level from saved stage state.

    Pore pressure in the saved top element is a transient unknown, so it is
    not the external pressure used for the exposed-face resultant.  The frame's
    START_PROFILE and elapsed fraction are the persisted source of that load.
    """
    lines = (frame / "fixed_stress.csv").read_text(encoding="utf-8-sig").splitlines()
    phase = next((line.split(",") for line in lines if line.startswith("PHASE,")), None)
    start = next((line.split(",", 1)[1] for line in lines if line.startswith("START_PROFILE,")), None)
    if phase is None or start is None:
        fail(f"{frame}: lacks PHASE or START_PROFILE")
    stage, elapsed, duration = int(phase[1]), float(phase[2]), float(phase[3])
    if duration <= 0 or elapsed < -1e-12 or elapsed > duration + 1e-12:
        fail(f"{frame}: invalid elapsed stage fraction")
    plans = {int(r[0]): r for r in water_cfg["stagePlans"]}
    target = plans.get(stage, [None, None, ""])[2] or start
    modes = {int(r[0]): str(r[2]).upper() for r in water_cfg.get("timeOptions", [])}
    fraction = min(1.0, max(0.0, elapsed / duration)) if modes.get(stage) == "LINEAR" else 1.0
    profiles = defaultdict(list)
    for profile, x, head, *_ in water_cfg["profiles"]:
        profiles[str(profile)].append((float(x), float(head)))
    def level(name, x):
        points = sorted(profiles.get(str(name), []))
        if not points:
            fail(f"{frame}: profile {name!r} is not in fixture")
        if x <= points[0][0]: return points[0][1]
        if x >= points[-1][0]: return points[-1][1]
        for (x0, h0), (x1, h1) in zip(points, points[1:]):
            if x0 <= x <= x1:
                return h0 + (h1 - h0) * (x - x0) / (x1 - x0)
        raise AssertionError("profile bracket")
    # The issue fixture is horizontally level; retain the spread in the report
    # as a guard that this special analytic resultant is not reused on a slope.
    heads = [(1.0 - fraction) * level(start, x) + fraction * level(target, x)
             for x in (0.0, 1.0)]
    return float(np.mean(heads)), max(heads) - min(heads), fraction, start, target


def compare_repeat_states(records, sheets):
    groups = defaultdict(list)
    for rec, field in zip(records, sheets):
        groups[str(rec["id"])].append(field)
    for state, fields in groups.items():
        baseline = fields[0]
        for later in fields[1:]:
            delta = max(float(np.max(np.abs(baseline[n] - later[n]))) for n in baseline)
            if delta > 1e-12:
                fail(f"repeated state {state} changed after frame reload: {delta:.3e}")


def check(fixture_path, actual_path, records_root=None):
    fixture, nodes, elements, material, coupled, water_cfg = fixture_model(fixture_path)
    actual = json.loads(actual_path.read_text(encoding="utf-8-sig"))
    records = actual.get("records")
    if not isinstance(records, list) or not records:
        fail("actual output has no records")
    if records_root is not None:
        # Relocate the archived raw frames without changing their provenance
        # paths in frame_forces.json or requiring the author's PC directory.
        root = records_root.resolve()
        for record in records:
            parts = str(record['frame']).replace('\\', '/').split('/')
            if 'targeted6' not in parts:
                fail('record lacks the targeted6 archive prefix')
            frame = root.joinpath(*parts[parts.index('targeted6'):]).resolve()
            if not frame.is_relative_to(root):
                fail('frame path escapes the extracted record directory')
            record['frame'] = str(frame)
    count = 2 * len(nodes)
    fields = []
    report = []
    gamma_w = float(water_cfg["gamma"])
    top = max(y for _, y, _, _ in nodes.values())
    support_nodes = [nid for nid, (_, y, _, cy) in nodes.items() if y == 0.0 and cy]
    for rec in records:
        frame = Path(rec["frame"])
        for filename in ("dof.csv", "gauss.csv", "water.csv"):
            if not (frame / filename).is_file():
                fail(f"{rec['id']}: saved frame lacks {filename}")
        field = sheet_field(rec, nodes)
        fields.append(field)
        locked = read_dof(frame / "dof.csv", count)
        stress = read_stress(frame / "gauss.csv")
        water = read_water(frame / "water.csv")
        internal = reconstruct_internal(nodes, elements, material, coupled, stress, water)
        maximum_load_error = 0.0
        maximum_reaction_error = 0.0
        for nid, values in field.items():
            base = 2 * (nid - 1)
            maximum_load_error = max(maximum_load_error,
                                     float(np.max(np.abs(values[4:6] - locked[base:base + 2]))))
            expected_reaction = internal[base:base + 2] - locked[base:base + 2]
            maximum_reaction_error = max(maximum_reaction_error,
                                         float(np.max(np.abs(values[6:8] - expected_reaction))))
        if maximum_load_error > TOL:
            fail(f"{rec['id']}: displayed G/H differs from locked_sw+locked_ap by {maximum_load_error:.3e} kN")
        affine_error = assert_one_dimensional_field(field)
        if maximum_reaction_error > TOL:
            fail(f"{rec['id']}: displayed I/J differs from independent effective-stress reaction by {maximum_reaction_error:.3e} kN")
        head, head_spread, fraction, start_profile, target_profile = adopted_head(frame, water_cfg)
        actual_ry = sum(field[n][7] for n in support_nodes)
        expected_ry = float(material[4]) * top + gamma_w * (head - top)
        body_error = abs(actual_ry - expected_ry)
        if body_error > TOL:
            fail(f"{rec['id']}: support Ry {actual_ry:.12g} != body+water {expected_ry:.12g} kN")
        report.append({"id": rec["id"], "frame": str(frame),
                       "max_locked_force_error_kN": maximum_load_error,
                       "max_reaction_error_kN": maximum_reaction_error,
                       "affine_displacement_error_m": affine_error,
                       "adopted_head_m": head, "head_spread_m": head_spread,
                       "water_fraction": fraction, "start_profile": start_profile,
                       "target_profile": target_profile,
                       "support_Ry_kN": actual_ry, "expected_Ry_kN": expected_ry,
                       "support_Ry_error_kN": body_error})
    compare_repeat_states(records, fields)
    return {"status": "PASS", "fixture": str(fixture_path),
            "actual_book_sha256": actual.get("book_sha256"), "frames": report,
            "tolerance_kN": TOL,
            "basis": "G/H=accepted locked self+applied loads; I/J=-B^T(sigma-alpha*pI)-Qp-load"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fixture", type=Path, required=True)
    ap.add_argument("--actual", type=Path, required=True)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--records-root", type=Path, help="Root of the extracted records.zip")
    args = ap.parse_args()
    try:
        result = check(args.fixture, args.actual, args.records_root)
    except Exception as exc:
        result = {"status": "FAIL", "error": str(exc)}
        args.out.write_text(json.dumps(result, indent=2), encoding="utf-8")
        print(json.dumps(result, indent=2), file=sys.stderr)
        return 1
    args.out.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
