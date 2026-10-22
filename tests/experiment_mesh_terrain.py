"""Deterministic Python model of MeshAdapt.BuildTerrain's Q8 grid assembly.

The legacy path is intentionally modelled with an edge dictionary.  The
structured path uses horizontal/vertical edge arrays, while assigning IDs in
the same first-seen order as the legacy loop.  This makes the comparison
about storage/lookup strategy rather than a changed mesh numbering scheme.
"""
from __future__ import annotations

import json
import os
import time
import random
from typing import Callable


def _vertices(nx: int, ny: int, height: Callable[[float], float]):
    out = []
    for row in range(ny + 1):
        for col in range(nx + 1):
            x = col / nx
            out.append((x, height(x) * row / ny))
    return out


def legacy_mesh(nx: int, ny: int, height: Callable[[float], float]):
    nodes = _vertices(nx, ny, height)
    edges = {}
    conn = []
    for row in range(ny):
        for col in range(nx):
            v1 = row * (nx + 1) + col
            v2 = v1 + 1
            v4 = v1 + nx + 1
            v3 = v4 + 1
            vertices = (v1, v2, v3, v4)
            mids = []
            for a, b in zip(vertices, vertices[1:] + vertices[:1]):
                key = (min(a, b), max(a, b))
                if key not in edges:
                    edges[key] = len(nodes)
                    nodes.append(((nodes[a][0] + nodes[b][0]) / 2,
                                  (nodes[a][1] + nodes[b][1]) / 2))
                mids.append(edges[key])
            conn.append(vertices + tuple(mids))
    return nodes, conn


def structured_mesh(nx: int, ny: int, height: Callable[[float], float]):
    nodes = _vertices(nx, ny, height)
    # Arrays are direct structured storage; IDs are assigned on first use to
    # retain BuildTerrain's exact node order.
    horizontal = [[None] * nx for _ in range(ny + 1)]
    vertical = [[None] * (nx + 1) for _ in range(ny)]
    conn = []
    for row in range(ny):
        for col in range(nx):
            v1 = row * (nx + 1) + col
            v2 = v1 + 1
            v4 = v1 + nx + 1
            v3 = v4 + 1
            mids = []
            for kind, i, j, a, b in (
                ("h", row, col, v1, v2), ("v", row, col + 1, v2, v3),
                ("h", row + 1, col, v3, v4), ("v", row, col, v4, v1)):
                table = horizontal if kind == "h" else vertical
                if table[i][j] is None:
                    table[i][j] = len(nodes)
                    nodes.append(((nodes[a][0] + nodes[b][0]) / 2,
                                  (nodes[a][1] + nodes[b][1]) / 2))
                mids.append(table[i][j])
            conn.append((v1, v2, v3, v4, *mids))
    return nodes, conn


def _case(name, nx, ny, height):
    old_nodes, old_conn = legacy_mesh(nx, ny, height)
    new_nodes, new_conn = structured_mesh(nx, ny, height)
    assert old_nodes == new_nodes and old_conn == new_conn
    ne = nx * ny
    actual = 3 * nx * ny + 2 * nx + 2 * ny + 1
    legacy_capacity = ne * 30 + 100
    return {"name": name, "nx": nx, "ny": ny, "nodes": len(old_nodes),
            "elements": ne, "exact_match": True,
            "legacy_xx_capacity": legacy_capacity,
            "actual_node_count": actual,
            "capacity_over_allocation": legacy_capacity - actual}


def _on_segment(a, b, p, tol, mesh_scale):
    ax, ay = a
    bx, by = b
    px, py = p
    cross = (bx - ax) * (py - ay) - (by - ay) * (px - ax)
    dot = (px - ax) * (px - bx) + (py - ay) * (py - by)
    return abs(cross) <= tol * mesh_scale and dot <= tol * tol


def _on_segment_prefilter(a, b, p, tol, mesh_scale):
    length = ((b[0] - a[0]) ** 2 + (b[1] - a[1]) ** 2) ** 0.5
    padding = tol + (tol * mesh_scale / length if length else 0.0)
    lo = min(a[0], b[0]) - padding
    hi = max(a[0], b[0]) + padding
    return lo <= p[0] <= hi and _on_segment(a, b, p, tol, mesh_scale)


def _prefilter_experiment():
    tol, mesh_scale = 1.0e-9, 10.0
    segments = [((0.0, 0.0), (1.0, 0.0)),
                ((2.0, -1.0), (2.0, 3.0)),
                ((1.0, 1.0), (1.0 + 1.0e-12, 1.0)),
                ((-3.0, 2.0), (4.0, -1.0))]
    rng = random.Random(20261022)
    points = [(x, y) for x, y in [
        (0.0, 0.0), (1.0, 0.0), (2.0, -1.0), (2.0, 3.0),
        (1.0 + tol, 0.0), (1.0 - tol, 0.0), (2.0 + tol, 1.0),
        (-1.0e-8, 0.0), (1.0 + 1.0e-8, 0.0)]]
    points.extend((rng.uniform(-5, 5), rng.uniform(-5, 5)) for _ in range(1000))
    original = []
    filtered = []
    candidates = 0
    for a, b in segments:
        for p in points:
            hit = _on_segment(a, b, p, tol, mesh_scale)
            original.append(hit)
            length = ((b[0] - a[0]) ** 2 + (b[1] - a[1]) ** 2) ** 0.5
            padding = tol + tol * mesh_scale / length
            in_x = min(a[0], b[0]) - padding <= p[0] <= max(a[0], b[0]) + padding
            candidates += in_x
            filtered.append(_on_segment_prefilter(a, b, p, tol, mesh_scale))
    assert all(not want or got for want, got in zip(original, filtered))
    assert original == filtered
    return {"passed": True, "segments": len(segments), "points_per_segment": len(points),
            "original_matches": sum(original), "filtered_matches": sum(filtered),
            "x_prefilter_candidates": candidates,
            "x_prefilter_reduction": 1.0 - candidates / len(original),
            "tolerance": tol, "mesh_scale": mesh_scale,
            "note": "Filtered scan invokes exact OnSegment after conservative x bounds."}


def main():
    cases = [("flat", 1, 1, lambda x: 1.0),
             ("slope", 3, 2, lambda x: 1.0 + 0.4 * x),
             ("kink", 5, 4, lambda x: 1.0 + (0.7 * x if x <= .4 else .28 + .25 * x)),
             ("large_slope", 40, 30, lambda x: 1.0 + .6 * x)]
    comparisons = [_case(*case) for case in cases]
    prefilter = _prefilter_experiment()
    timings = []
    for name, nx, ny, height in cases:
        loops = 1 if nx * ny > 500 else 30
        t0 = time.perf_counter()
        for _ in range(loops): legacy_mesh(nx, ny, height)
        legacy_ms = (time.perf_counter() - t0) * 1000 / loops
        t0 = time.perf_counter()
        for _ in range(loops): structured_mesh(nx, ny, height)
        structured_ms = (time.perf_counter() - t0) * 1000 / loops
        timings.append({"name": name, "legacy_ms": legacy_ms,
                        "structured_ms": structured_ms, "loops": loops})
    report = {"description": "BuildTerrain legacy dictionary versus direct structured edge arrays",
              "numbering_preserved": all(x["exact_match"] for x in comparisons),
              "on_segment_prefilter": prefilter,
              "comparisons": comparisons, "timings": timings,
              "memory_model": {"legacy_xx_and_nodeRows_capacity": "ne*30+100",
                               "structured_node_count": "3*nx*ny+2*nx+2*ny+1",
                               "note": "Python timings are directional only; they do not claim VBA speedup."}}
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    target = os.path.join(root, "tests", "tmp", "mesh_light", "python_experiment.json")
    os.makedirs(os.path.dirname(target), exist_ok=True)
    with open(target, "w", encoding="utf-8", newline="\n") as f:
        json.dump(report, f, indent=2)
        f.write("\n")
    print(json.dumps(report, indent=2))


if __name__ == "__main__":
    main()
