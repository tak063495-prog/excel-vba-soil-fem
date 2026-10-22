"""Independent check for the sparse boundary-edge tables used by Mesh.PINPUT.

The reference implementation intentionally mirrors the old dense tables.  The
candidate implementation stores only directed edges that occur in the input
boundaries.  This is a Python-only equivalence and memory-scaling check; it
does not run Excel or claim a measured Excel peak-memory reduction.
"""
from __future__ import annotations

import tracemalloc
import argparse
import csv
import json
from pathlib import Path


def dense_edges(loops, spacings):
    n = max((p for loop in loops for p in loop), default=0)
    imno = [[0.0] * (n + 1) for _ in range(n + 1)]
    deno = [[0] * (n + 1) for _ in range(n + 1)]
    sms = [[0] * (n + 1) for _ in range(n + 1)]
    return _populate(loops, spacings, imno, deno, sms)


def sparse_edges(loops, spacings):
    return _populate(loops, spacings, {}, {}, {})


def _get(table, a, b):
    if isinstance(table, list):
        return table[a][b]
    return table.get((a, b), 0)


def _set(table, a, b, value):
    if isinstance(table, list):
        table[a][b] = value
    else:
        table[(a, b)] = value


def _populate(loops, spacings, imno, deno, sms):
    # PINPUT first chooses the smaller spacing when a directed edge is repeated.
    for loop, spacing in zip(loops, spacings):
        for a, b in zip(loop, loop[1:] + loop[:1]):
            old = _get(imno, a, b)
            if old == 0 or spacing < old:
                _set(imno, a, b, spacing)
            old = _get(imno, b, a)
            if old == 0 or spacing < old:
                _set(imno, b, a, spacing)
            ab, ba = _get(imno, a, b), _get(imno, b, a)
            if ab > ba:
                _set(imno, a, b, ba)
            else:
                _set(imno, b, a, ab)

    inserted = 0
    for loop, spacing in zip(loops, spacings):
        for a, b in zip(loop, loop[1:] + loop[:1]):
            edge_length = abs(b - a)  # synthetic length; only subdivision count matters
            count = max(int(edge_length / _get(imno, a, b)) - 1, 0)
            if _get(deno, a, b) == 0 and count >= 1:
                inserted += count
                _set(sms, a, b, count)
                _set(sms, b, a, count)
                _set(deno, a, b, inserted - count + 1)
                _set(deno, b, a, inserted)
    return imno, deno, sms, inserted


def _logical(table):
    if isinstance(table, list):
        return {(a, b): value for a, row in enumerate(table) for b, value in enumerate(row) if value}
    return {key: value for key, value in table.items() if value}


def run():
    # Same edge, reverse edge, and a tighter spacing on the repeated edge.
    loops = [[1, 4, 7], [7, 4, 1], [2, 5, 8]]
    spacings = [1.0, 0.5, 2.0]
    dense = dense_edges(loops, spacings)
    sparse = sparse_edges(loops, spacings)
    assert _logical(dense[0]) == _logical(sparse[0])
    assert _logical(dense[1]) == _logical(sparse[1])
    assert _logical(dense[2]) == _logical(sparse[2])
    assert dense[3] == sparse[3]

    # An interior-only point is never an edge key; repeated reads return zero.
    assert _get(sparse[0], 99, 100) == 0
    assert _get(sparse[1], 99, 100) == 0
    assert _get(sparse[2], 99, 100) == 0
    # Boundary closure includes the last-to-first edge exactly once.
    closed = sparse_edges([[10, 11, 12]], [1.0])
    assert _get(closed[0], 12, 10) == 1.0

    # Report a reproducible scaling estimate for the three dense tables.  The
    # estimate excludes Python object overhead and is intentionally conservative.
    print("case, dense_bytes_estimate, sparse_edge_values")
    for n in (200, 1000, 10000):
        dense_bytes = n * n * (8 + 4 + 4)
        sparse_values = 4 * n  # four directed boundary edges per input node
        print(f"{n},{dense_bytes},{sparse_values}")

    tracemalloc.start()
    sparse_edges([list(range(1, 10001))], [1.0])
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    print(f"sparse_python_peak_bytes,{peak}")
    print("PASS")


def compare_excel(reference, candidate, report):
    a, b = Path(reference), Path(candidate)
    def load(root):
        assert json.loads((root/'result.json').read_text(encoding='utf-8-sig'))['status']=='PASS'
        with (root/'memory.csv').open(encoding='utf-8-sig', newline='') as f:
            return list(csv.DictReader(f))
    old, new = load(a), load(b)
    for filename in ['memory.csv.nodes.csv', 'memory.csv.boundary.csv']:
        assert (a/filename).read_bytes()==(b/filename).read_bytes(), filename
    rows=[]
    for case, n in [('polygon',200),('polygon',1000),('polygon',1800),('shared_reverse_subdivided',6)]:
        def values(samples):
            r={x['label']:x for x in samples if x['case']==case and int(x['n'])==n}
            assert len(r)==4, (case,n,r.keys())
            start=next(v for k,v in r.items() if k.startswith('before_') and k!='before_prepare')
            end=r['after_tables']
            return dict(private_commit_before=int(start['pagefile_bytes']),private_commit_at_tables=int(end['pagefile_bytes']),private_commit_increase=int(end['pagefile_bytes'])-int(start['pagefile_bytes']),working_set_at_tables=int(end['working_set_bytes']),pinput_seconds=float(r['after_pinput']['timer'])-float(start['timer']))
        x,y=values(old),values(new)
        rows.append(dict(case=case,input_points=n,old=x,new=y,private_commit_increase_saved_bytes=x['private_commit_increase']-y['private_commit_increase']))
    # A large-boundary allocation probe must demonstrate a meaningful real saving.
    assert rows[2]['private_commit_increase_saved_bytes']>30_000_000,rows[2]
    result=dict(status='PASS',boundary_outputs_exact=True,cases=rows,measurement='GetProcessMemoryInfo PagefileUsage (private committed memory) and WorkingSetSize sampled inside PINPUT before allocation and while tables are live; whole Excel process, not isolated allocator or lifetime peak.',limitations='One sequential run per book/case. Regular polygons do not add subdivision nodes; the separate shared-edge case exercises reverse numbering and different subdivision spacings. Timing includes instrumentation and is not a full-meshing speed benchmark.')
    Path(report).write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False,indent=2))


if __name__ == "__main__":
    p=argparse.ArgumentParser();p.add_argument('--compare',nargs=2);p.add_argument('--report');args=p.parse_args()
    if args.compare:
        assert args.report,'--report required'
        compare_excel(*args.compare,args.report)
    else:run()
