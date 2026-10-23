"""Compare three independent Mandel checker reports at their common accepted times.

This is deliberately a report combiner: pressure errors are those produced by
``check_fixed_stress_mandel.py`` against its analytical series at reconstructed
physical Gauss points.  It does not interpolate either numerical history or
replace the prescribed-displacement check with a computed settlement check.
"""
import argparse
import json
import math
import sys
from pathlib import Path


TIME_REL_TOL = 1.0e-9
TIME_ABS_TOL = 1.0e-7


def read_json(path):
    value = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if not isinstance(value, dict) or not isinstance(value.get("frames"), list):
        raise ValueError(f"{path}: expected a Mandel checker report with frames")
    return value


def read_policy(path):
    value = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    required = {"fine_max_pressure_error_kPa", "fine_error_ratio_to_coarse_max",
                "prescribed_displacement_error_m", "cryer_peak_min_kPa"}
    if not isinstance(value, dict) or not required.issubset(value):
        raise ValueError(f"{path}: incomplete Mandel acceptance policy")
    return value


def transient_frames(report, name):
    frames = []
    for frame in report["frames"]:
        try:
            t = float(frame["time_s"])
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"{name}: invalid frame time") from exc
        error = frame.get("max_pressure_error_kPa")
        if error is None:
            continue  # The discrete t=0 pressure projection is outside this check.
        try:
            error = float(error)
            uy = float(frame["top_uy_prescribed_error_m"])
            centre = float(frame["centre_excess_pressure_kPa"])
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"{name}: incomplete transient frame at t={t}") from exc
        if not all(math.isfinite(x) for x in (t, error, uy, centre)):
            raise ValueError(f"{name}: non-finite transient value at t={t}")
        frames.append({"time_s": t, "pressure_error_kPa": error,
                       "uy_error_m": uy, "centre_excess_pressure_kPa": centre})
    if not frames:
        raise ValueError(f"{name}: no transient frames")
    return sorted(frames, key=lambda frame: frame["time_s"])


def full_history_frames(report, name):
    """Read centre pressure and imposed displacement at every saved time, including t=0."""
    frames = []
    for frame in report["frames"]:
        try:
            t = float(frame["time_s"])
            uy = float(frame["top_uy_prescribed_error_m"])
            centre = float(frame["centre_excess_pressure_kPa"])
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"{name}: incomplete history frame") from exc
        if not all(math.isfinite(x) for x in (t, uy, centre)):
            raise ValueError(f"{name}: non-finite history value at t={t}")
        row = {"time_s": t, "uy_error_m": uy,
               "centre_excess_pressure_kPa": centre}
        # Older reports predate the independent all-node calculation.  Keep
        # them usable for pressure/Cryer comparison, but require this data for
        # the new early-run acceptance check below.
        ux = frame.get("all_node_ux_error_m")
        uy_node = frame.get("all_node_uy_error_m")
        if ux is not None or uy_node is not None:
            try:
                ux, uy_node = float(ux), float(uy_node)
            except (TypeError, ValueError) as exc:
                raise ValueError(f"{name}: incomplete all-node displacement frame at t={t}") from exc
            if not all(math.isfinite(x) for x in (ux, uy_node)):
                raise ValueError(f"{name}: non-finite all-node displacement error at t={t}")
            row["all_node_ux_error_m"] = ux
            row["all_node_uy_error_m"] = uy_node
        frames.append(row)
    if not frames:
        raise ValueError(f"{name}: no saved frames")
    return sorted(frames, key=lambda frame: frame["time_s"])


def same_time(a, b):
    return abs(a - b) <= max(TIME_ABS_TOL, TIME_REL_TOL * max(1.0, abs(a), abs(b)))


def common_triples(coarse, fine, early):
    """Return matching accepted times without interpolation or nearest-time reuse."""
    triples = []
    fine_index = early_index = 0
    for c in coarse:
        while fine_index < len(fine) and fine[fine_index]["time_s"] < c["time_s"] - TIME_ABS_TOL:
            fine_index += 1
        while early_index < len(early) and early[early_index]["time_s"] < c["time_s"] - TIME_ABS_TOL:
            early_index += 1
        if fine_index >= len(fine) or early_index >= len(early):
            break
        f, e = fine[fine_index], early[early_index]
        if same_time(c["time_s"], f["time_s"]) and same_time(c["time_s"], e["time_s"]):
            triples.append((c, f, e))
            fine_index += 1
            early_index += 1
    if not triples:
        raise ValueError("no common transient accepted times across coarse, fine, and early reports")
    return triples


def history_cryer(frames, name, minimum_rise):
    """Use the complete original history, rather than the short early refinement."""
    initial = frames[0]["centre_excess_pressure_kPa"]
    peak_index = max(range(len(frames)), key=lambda i: frames[i]["centre_excess_pressure_kPa"])
    peak = frames[peak_index]["centre_excess_pressure_kPa"]
    later = frames[peak_index + 1:]
    has_decline = bool(later) and min(row["centre_excess_pressure_kPa"] for row in later) < peak
    return {
        "source": name,
        "initial_kPa": initial,
        "peak_kPa": peak,
        "peak_time_s": frames[peak_index]["time_s"],
        "peak_rise_kPa": peak - initial,
        "subsequent_decline": has_decline,
        "passes": peak - initial > minimum_rise and has_decline,
    }


def max_or_none(values):
    return max(values) if values else None


def all_node_summary(frames):
    """Report components and combined norm bound without inventing missing old data."""
    available = [row for row in frames if "all_node_ux_error_m" in row]
    if not available:
        return {"available_frames": 0, "total_frames": len(frames),
                "max_ux_error_m": None, "max_uy_error_m": None, "max_component_error_m": None}
    return {
        "available_frames": len(available), "total_frames": len(frames),
        "max_ux_error_m": max(row["all_node_ux_error_m"] for row in available),
        "max_uy_error_m": max(row["all_node_uy_error_m"] for row in available),
        "max_component_error_m": max(max(row["all_node_ux_error_m"], row["all_node_uy_error_m"])
                                     for row in available),
    }


def compare(coarse_report, fine_report, early_report, policy):
    pressure_limit = float(policy["fine_max_pressure_error_kPa"])
    ratio_limit = float(policy["fine_error_ratio_to_coarse_max"])
    uy_limit = float(policy["prescribed_displacement_error_m"])
    cryer_rise = float(policy["cryer_peak_min_kPa"])
    all_node_limit = float(policy.get("all_node_displacement_error_m", 1.0e-5))
    if min(pressure_limit, ratio_limit, uy_limit, cryer_rise, all_node_limit) < 0:
        raise ValueError("policy thresholds must be non-negative")

    coarse = transient_frames(coarse_report, "coarse")
    fine = transient_frames(fine_report, "fine")
    early = transient_frames(early_report, "early")
    coarse_history = full_history_frames(coarse_report, "coarse")
    fine_history = full_history_frames(fine_report, "fine")
    early_history = full_history_frames(early_report, "early")
    triples = common_triples(coarse, fine, early)
    coarse_common = max(row[0]["pressure_error_kPa"] for row in triples)
    fine_common = max(row[1]["pressure_error_kPa"] for row in triples)
    early_common = max(row[2]["pressure_error_kPa"] for row in triples)
    if coarse_common <= 0.0:
        raise ValueError("coarse common-time maximum is zero; error ratios are undefined")
    max_uy = max_or_none(row["uy_error_m"] for frames in (coarse_history, fine_history, early_history) for row in frames)
    history_by_time = [{row["time_s"]: row for row in frames}
                       for frames in (coarse_history, fine_history, early_history)]
    matched_all_node = []
    for column, label in enumerate(("coarse", "fine", "early")):
        rows = []
        for triple in triples:
            source = triple[column]
            match = next((row for t, row in history_by_time[column].items()
                          if same_time(t, source["time_s"])), None)
            if match is not None and "all_node_ux_error_m" in match:
                rows.append(match)
        summary = all_node_summary(rows)
        summary["source"] = label
        summary["common_frames"] = len(triples)
        matched_all_node.append(summary)
    cryer = [history_cryer(coarse_history, "coarse full history", cryer_rise),
             history_cryer(fine_history, "fine full history", cryer_rise)]

    result = {
        "status": "FAIL",
        "scope": (
            "Pressure maxima are analytical checker errors at only exact common accepted times; "
            "there is no time interpolation. The third run establishes temporal refinement over the reported common-time range "
            "only. Cryer uses the complete coarse and fine histories."
        ),
        "precision": {"time_absolute_tolerance_s": TIME_ABS_TOL,
                      "time_relative_tolerance": TIME_REL_TOL},
        "policy": {"max_pressure_error_kPa": pressure_limit,
                   "max_early_to_coarse_ratio": ratio_limit,
                   "max_imposed_uy_error_m": uy_limit,
                   "max_all_node_displacement_error_m": all_node_limit,
                   "min_cryer_rise_kPa": cryer_rise},
        "common_times": {"count": len(triples),
                         "first_s": triples[0][0]["time_s"],
                         "last_s": triples[-1][0]["time_s"]},
        "common_time_pressure_error_kPa": {
            "coarse_max": coarse_common, "fine_max": fine_common, "early_max": early_common,
            "fine_to_coarse_ratio": fine_common / coarse_common,
            "early_to_coarse_ratio": early_common / coarse_common,
            "early_to_fine_ratio": early_common / fine_common if fine_common > 0.0 else None,
        },
        "imposed_top_uy": {"max_error_m": max_uy, "passes": max_uy < uy_limit},
        "computed_all_node_displacement": {
            "all_time": [dict({"source": label}, **all_node_summary(frames))
                         for label, frames in zip(("coarse", "fine", "early"),
                                                  (coarse_history, fine_history, early_history))],
            "matched_transient_common_times": matched_all_node,
            "early_limit_m": all_node_limit,
        },
        "cryer_center_excess_pressure": cryer,
    }
    pressure_ok = early_common < pressure_limit
    refinement_ok = early_common / coarse_common < ratio_limit
    cryer_ok = all(item["passes"] for item in cryer)
    early_all_node = matched_all_node[2]
    all_node_ok = (early_all_node["available_frames"] == early_all_node["common_frames"]
                   and early_all_node["max_component_error_m"] < all_node_limit)
    result["checks"] = {"third_run_pressure_limit": pressure_ok,
                        "third_run_to_coarse_ratio": refinement_ok,
                        "imposed_uy": max_uy < uy_limit,
                        "early_matched_all_node_displacement": all_node_ok,
                        "coarse_and_fine_cryer": cryer_ok}
    if all(result["checks"].values()):
        result["status"] = "PASS"
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--coarse", required=True, help="coarse Mandel checker JSON")
    parser.add_argument("--fine", required=True, help="spatially refined Mandel checker JSON")
    parser.add_argument("--early", required=True, help="short time-refined Mandel checker JSON")
    parser.add_argument("--policy", required=True, help="acceptance-policy JSON")
    parser.add_argument("--out", required=True, help="comparison JSON output")
    args = parser.parse_args(argv)
    try:
        result = compare(read_json(args.coarse), read_json(args.fine), read_json(args.early), read_policy(args.policy))
    except (OSError, ValueError, KeyError, json.JSONDecodeError) as exc:
        result = {"status": "ERROR", "error": str(exc)}
    Path(args.out).write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, indent=2))
    return 0 if result["status"] == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
