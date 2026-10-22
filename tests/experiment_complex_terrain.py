"""Exact orthogonal non-Delaunay mesh experiment.

All boundary vertex X/Y coordinates form shared grid lines. Axis-aligned
material masks and holes are preserved exactly; diagonals are rejected.
Independent Q8 derivatives check all 25 reference points per element.
"""
from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Iterable


Point = tuple[float, float]
Loop = tuple[Point, ...]


def area(poly: Iterable[Point]) -> float:
    p = tuple(poly)
    return abs(sum(p[i][0] * p[(i + 1) % len(p)][1]
                   - p[(i + 1) % len(p)][0] * p[i][1]
                   for i in range(len(p))) / 2.0)


def inside(p: Point, poly: Loop) -> bool:
    x, y = p
    hit = False
    for a, b in zip(poly, poly[1:] + poly[:1]):
        if (a[1] > y) != (b[1] > y):
            cross = (b[0] - a[0]) * (y - a[1]) / (b[1] - a[1]) + a[0]
            if x < cross:
                hit = not hit
    return hit


def q8_jacobian_min(nodes, cells):
    import numpy as np
    from check_groundwater import shape
    xy=np.array(nodes)
    return min(float(np.linalg.det(xy[list(c)].T@shape(r,s)[1].T))
               for c in cells for r in (-1,-.5,0,.5,1) for s in (-1,-.5,0,.5,1))


def cartesian_orthogonal(outer: Loop, holes: tuple[Loop, ...], material_regions: dict[str, Loop], h: float):
    """Exact conforming path for orthogonal (axis-aligned) geometry.

    All polygon edges must be horizontal or vertical.  Every vertex coordinate
    is a grid line and intermediate lines are inserted at ``h``.  Cells are
    retained only from the inside classification; because every boundary is
    a grid line this does not create cut cells or hanging nodes.
    """
    loops = (outer,) + holes + tuple(material_regions.values())
    if any(a[0] != b[0] and a[1] != b[1]
           for loop in loops for a, b in zip(loop, loop[1:] + loop[:1])):
        raise ValueError("斜辺を含むため直交格子を適用できません")
    minx = min(x for loop in loops for x, _ in loop)
    maxx = max(x for loop in loops for x, _ in loop)
    miny = min(y for loop in loops for _, y in loop)
    maxy = max(y for loop in loops for _, y in loop)
    xs = {x for loop in loops for x, _ in loop}
    ys = {y for loop in loops for _, y in loop}
    x = minx
    while x < maxx - 1e-12:
        xs.add(x)
        x += h
    y = miny
    while y < maxy - 1e-12:
        ys.add(y)
        y += h
    xs, ys = sorted(xs | {maxx}), sorted(ys | {maxy})
    used: dict[Point, int] = {}
    cells: list[tuple[int, int, int, int]] = []
    labels: list[str] = []
    def node(p: Point) -> int:
        if p not in used:
            used[p] = len(used)
        return used[p]
    for xa, xb in zip(xs, xs[1:]):
        for ya, yb in zip(ys, ys[1:]):
            c = ((xa + xb) / 2.0, (ya + yb) / 2.0)
            if not inside(c, outer) or any(inside(c, hole) for hole in holes):
                continue
            label = "SOIL"
            for name, poly in material_regions.items():
                if inside(c, poly):
                    label = name
                    break
            cells.append((node((xa, ya)), node((xb, ya)), node((xb, yb)), node((xa, yb))))
            labels.append(label)
    nodes = [None] * len(used)
    for p, i in used.items():
        nodes[i] = p
    edge_nodes: dict[tuple[int, int], int] = {}
    q8: list[tuple[int, ...]] = []
    for c in cells:
        mids = []
        for a, b in zip(c, c[1:] + c[:1]):
            key = (min(a, b), max(a, b))
            if key not in edge_nodes:
                pa, pb = nodes[a], nodes[b]
                edge_nodes[key] = len(nodes)
                nodes.append(((pa[0] + pb[0]) / 2, (pa[1] + pb[1]) / 2))
            mids.append(edge_nodes[key])
        q8.append(c + tuple(mids))
    return nodes, q8, labels


def run_exact_case(name: str, outer: Loop, holes: tuple[Loop, ...], regions: dict[str, Loop], h: float,
                   expected_area: dict[str, float]):
    try:
        nodes, q8, labels = cartesian_orthogonal(outer, holes, regions, h)
        cells = q8
        by_material = {m: sum(area([nodes[i] for i in c[:4]]) for c, mm in zip(cells, labels) if mm == m)
                       for m in sorted(set(labels))}
        if by_material != expected_area:
            raise ValueError(f"材料面積不一致: actual={by_material} expected={expected_area}")
        if set(range(len(nodes))) != {i for c in q8 for i in c}:
            # Every corner and midpoint must occur in an element.
            raise ValueError("孤立節点または未使用節点があります")
        jac = q8_jacobian_min(nodes, cells)
        if jac <= 0:
            raise ValueError(f"non-positive Q8 Jacobian: {jac}")
        return {"name": name, "status": "PASS", "nodes": len(nodes), "elements": len(cells),
                "areas_by_material": by_material, "min_jacobian": jac,
                "hole_area": sum(area(hole) for hole in holes)}
    except ValueError as exc:
        return {"name": name, "status": "REJECT", "reason": str(exc)}


def main() -> None:
    # stepped cut/fill with a wall strip; the wall is a separate material region
    stepped = ((0, 0), (20, 0), (20, 8), (14, 8), (14, 6), (10, 6),
               (10, 10), (4, 10), (4, 7), (0, 7))
    wall = ((9, 0), (10, 0), (10, 6), (9, 6))
    # rectangular tunnel void and lining ring (two aligned rectangles)
    tunnel_outer = ((0, 0), (24, 0), (24, 12), (0, 12))
    void = ((8, 3), (16, 3), (16, 7), (8, 7))
    lining = ((7, 2), (17, 2), (17, 8), (7, 8))
    cases = [
        run_exact_case("stepped_cut_fill_wall_exact", stepped, (), {"WALL": wall}, 1.0,
                       {"SOIL": 154.0, "WALL": 6.0}),
        run_exact_case("tunnel_void_lining_exact", tunnel_outer, (void,), {"LINING": lining}, 1.0,
                       {"LINING": 28.0, "SOIL": 228.0}),
    ]
    assert all(x["status"] == "PASS" for x in cases), cases
    out = Path(__file__).parent / "tmp" / "complex_terrain" / "vertical_slab_results.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps({"algorithm": "exact_orthogonal_q8", "cases": cases},
                              indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps({"algorithm": "exact_orthogonal_q8", "cases": cases}, ensure_ascii=False))


if __name__ == "__main__":
    main()
