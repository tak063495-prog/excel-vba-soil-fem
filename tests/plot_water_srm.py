"""Plot a verified water-SRM result without inventing a plastic band.

Usage:
  python tests/plot_water_srm.py --run tests/tmp/groundwater/acceptance20/wet_srm \
      --out tests/tmp/water_display/wet_srm_model_result.png

The result folder must contain result.json; the fixture is selected by the
fixture_sha256 when it is available, or by --fixture.  No Excel calculation is
performed by this script.
"""
from __future__ import annotations
import argparse, hashlib, json
from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
from matplotlib.colors import Normalize

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_FIXTURE = ROOT / "tests/fixtures/groundwater_slope_low_srm.json"

def read(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", type=Path, required=True)
    ap.add_argument("--fixture", type=Path, default=DEFAULT_FIXTURE)
    ap.add_argument("--deformed", action="store_true")
    ap.add_argument("--out", type=Path, required=True)
    args = ap.parse_args()
    run = args.run.resolve(); fixture_path = args.fixture.resolve()
    result = read(run / "result.json")
    fixture = read(fixture_path)
    expected_hash = result.get("fixture_sha256")
    actual_hash = hashlib.sha256(fixture_path.read_bytes()).hexdigest().upper()
    if expected_hash and expected_hash.upper() != actual_hash:
        raise SystemExit(f"fixture hash mismatch: result={expected_hash}, fixture={actual_hash}")

    nodes = {int(r[0]): (float(r[1]), float(r[2])) for r in fixture["nodes"]}
    disp = [float(v) for v in result["disp"]]
    if len(disp) != 2 * len(nodes):
        raise SystemExit("result displacement length does not match fixture nodes")
    u = {n: (disp[2*(n-1)], disp[2*(n-1)+1]) for n in nodes}
    perimeter = (0, 4, 1, 5, 2, 6, 3, 7)
    original, deformed, mags = [], [], []
    for row in fixture["elements"]:
        ids = [int(v) for v in row[1:9]]
        original.append([nodes[ids[i]] for i in perimeter])
        deformed.append([(nodes[n][0] + u[n][0], nodes[n][1] + u[n][1]) for n in (ids[i] for i in perimeter)])
        mags.append(sum((u[n][0]**2 + u[n][1]**2) ** .5 for n in ids) / len(ids))

    water = fixture.get("water", {})
    profiles = water.get("profiles", [])
    # For the horizontal verified example, draw the supplied profile itself;
    # for a polyline, use all profile points without interpolation.
    profile_id = water.get("defaultProfile", "")
    profile = [(float(r[1]), float(r[2])) for r in profiles if str(r[0]) == str(profile_id)]
    mode = str(water.get("mode", "NONE")); flow = str(result.get("flow", ""))
    metrics = result.get("metrics", [])
    vmax = max(mags) if mags else 0.0
    norm = Normalize(vmin=0.0, vmax=vmax or 1.0)
    fig, (ax0, ax1) = plt.subplots(1, 2, figsize=(12, 5.3), constrained_layout=True)
    ax0.add_collection(PolyCollection(original, facecolors="#e8ecee", edgecolors="#7b858b", linewidths=.45))
    ax0.plot([], [], color="#7b858b", label="original Q8 mesh")
    if profile: ax0.plot(*zip(*profile), "--", color="#147eac", label=f"water profile {profile_id}")
    ax0.set_title(f"model / water input\nmode={mode}")
    ax0.legend(fontsize=8, loc="upper right"); ax0.set_aspect("equal")

    ax1.add_collection(PolyCollection(deformed if args.deformed else original, array=mags, cmap="viridis", norm=norm,
                                      edgecolors="#7b3030", linewidths=.45))
    if profile: ax1.plot(*zip(*profile), "--", color="#147eac", label="input water level")
    ax1.set_title(f"saved |u| on reference mesh (no deformation)\nflow={flow}" if not args.deformed else f"saved displacement, scale 1\nflow={flow}")
    ax1.legend(fontsize=8, loc="upper right"); ax1.set_aspect("equal")
    cb = fig.colorbar(ax1.collections[0], ax=ax1, shrink=.82)
    cb.set_label("mean nodal displacement |u| [m]")
    for ax in (ax0, ax1): ax.set_xlabel("X [m]"); ax.set_ylabel("elevation Y [m]")
    fs = ""
    if len(metrics) >= 8:
        fs = f" | accepted Fs {float(metrics[4]):.4f}"
        if float(metrics[6]) > 0:
            fs += f" / failed Fs {float(metrics[6]):.4f}"
    title = f"verified water SRM | status={result.get('status','')}" + fs
    fig.suptitle(title, fontsize=10)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(args.out, dpi=180)
    plt.close(fig)
    print(args.out)

if __name__ == "__main__": main()
