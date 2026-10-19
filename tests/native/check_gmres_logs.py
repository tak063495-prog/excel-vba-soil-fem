"""Independently check raw GMRES logs; allow only documented diagnostic differences."""
import argparse
import collections
import hashlib
import json
import re
from pathlib import Path


def inspect(record):
    data = (Path(record['RunOutputDirectory']) / 'solver_events.csv').read_bytes()
    assert len(data) == record['SolverEventsBytes']
    assert hashlib.sha256(data).hexdigest().upper() == record['SolverEventsSha256']
    kept, events, changed = [], collections.Counter(), []
    for line in data.splitlines(keepends=True):
        fields = line.split(b',', 6)
        if len(fields) < 7 or fields[5] == b'event':
            kept.append(line)
            continue
        event = fields[5]
        events[event.decode('ascii')] += 1
        if event.startswith(b'GMRES_'):
            continue
        if event == b'REGULARIZED_CORRECTION_REQUEST':
            normalized = re.sub(rb'(?<=;linear_relres=)[^;\r\n]*', b'<gmres_residual>', line)
            if normalized != line:
                changed.append(line.decode('latin1').rstrip())
            line = normalized
        kept.append(line)
    digest = hashlib.sha256(b''.join(kept)).hexdigest().upper()
    assert digest == record['DecisionEventsSha256']
    gmres = sum(v for k, v in events.items() if k.startswith('GMRES_'))
    assert gmres == record['GmresEventRows']
    assert sum(events.values()) - gmres == record['NonGmresSolverEventsRows']
    return kept, dict(event_counts=dict(events), decision_sha256=digest,
                      normalized_linear_residual_rows=len(changed),
                      first_normalized_input=changed[0] if changed else None)


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('directory'); ap.add_argument('--out', required=True)
    args = ap.parse_args()
    reports = []
    for i in range(1, 4):
        pair = [json.loads((Path(args.directory) / f'pair{i}_enabled{v}.json').read_text(encoding='utf-8-sig')) for v in (0, 1)]
        a, ra = inspect(pair[0]); b, rb = inspect(pair[1])
        assert a == b, 'An undocumented non-GMRES event changed'
        reports.append(dict(pair=i, reference=ra, candidate=rb))
    result = dict(status='PASS', pairs=reports,
                  rule='Exclude GMRES_* rows and only linear_relres in REGULARIZED_CORRECTION_REQUEST; all remaining bytes must match.')
    Path(args.out).write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8')
    print('PASS: raw log hashes, event counts and all remaining event bytes, 3 pairs')
