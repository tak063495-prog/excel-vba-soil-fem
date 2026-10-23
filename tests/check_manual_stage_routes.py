"""Prepare archived input tables and check the live ALL/result-table route."""
from pathlib import Path
import argparse
import gzip
import hashlib
import json

ROOT = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser()
parser.add_argument('--root', type=Path, default=ROOT/'tests/tmp/issue35_routes')
parser.add_argument('--prepare', action='store_true')
parser.add_argument('--road', default='road7')
parser.add_argument('--tunnel', default='tunnel2')
args = parser.parse_args()
args.root.mkdir(parents=True, exist_ok=True)
report = {'status': 'PASS', 'scope': 'Live Excel public UI route; historical figure data remain separately identified', 'cases': []}
for name, archive, last_stage, folder in [
    ('road', 'ex08_road_result_20261021.json.gz', 11, args.road),
    ('tunnel', 'ex10_tunnel_result_20261021.json.gz', 9, args.tunnel),
]:
    source = ROOT/'docs/manuals/examples'/archive
    old = json.loads(gzip.decompress(source.read_bytes()))['tables']
    if args.prepare:
        data = {k: old[k] for k in ['材料データ', '節点データ', '要素データ', 'ステージ', '載荷']}
        data['ステージ'] = [r for r in data['ステージ'] if r[1] != 'SRM']
        (args.root/f'{name}_input.json').write_text(json.dumps(data, ensure_ascii=False)+'\n', encoding='utf8')
        continue
    result_path = args.root/folder/'result.json'
    r = json.loads(result_path.read_text(encoding='utf-8-sig'))
    assert r['status'] == r['compile'] == 'PASS' and r['application_code_unchanged']
    assert r['save_route'].endswith('D8=ALL')
    stages = list(range(1, last_stage+1))
    for route in r['routes']:
        assert route['stage_ids'] == stages
        assert route['original'] in ['結果節点', '結果要素']
    rows = r['tables']['結果節点'][1:]
    by_key = {tuple(row[:2]): row for row in rows}
    assert len(rows) == len(by_key) == (len(old['節点データ'])-1)*last_stage
    expected_stages = [1, 4, 6, 11] if name == 'road' else stages
    assert [v['stage'] for v in r['filters']] == expected_stages
    point = 19 if name == 'road' else 351
    for v in r['filters']:
        assert v['node'] == point
        row = by_key[(v['stage'], point)]
        assert v['values_A_B_F_J'] == [row[i] for i in (0, 1, 5, 9)]
    connectivity = {int(v[0]): set(v[1:9]) for v in old['要素データ'][1:]}
    active = {stage: {int(v[1]) for v in r['tables']['結果要素'][1:] if v[0] == stage} for stage in stages}
    for stage in stages:
        assert active[stage] == {int(v[1]) for v in old['結果要素'][1:] if v[0] == stage}
        assert any(point in connectivity[e] for e in active[stage])
    baseline = {tuple(v[:2]): v for v in old['結果節点'][1:] if v[0] <= last_stage}
    deltas = {key: max(abs(v[c]-baseline[tuple(v[:2])][c]) for v in rows) for key, c in [('ux_m', 4), ('uy_m', 5), ('Rx_kN', 8), ('Ry_kN', 9)]}
    # The archive predates solver/material fixes: report differences, never
    # assert identical numerical results across these versions.
    values = {str(s): {'uy_m': by_key[(s, point)][5], 'additional_settlement_mm': -(by_key[(s, point)][5]-by_key[(1, point)][5])*1000} for s in expected_stages}
    report['cases'].append({'name': name, 'normal_stage_ids': stages, 'node_rows': len(rows),
                            'element_rows': len(r['tables']['結果要素'])-1,
                            'filtered_node': point, 'filtered_values': values,
                            'max_absolute_delta_from_historical_nodes': deltas,
                            'historical_archive_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
                            'current_book_sha256': r['source_sha256'].lower(),
                            'result_sha256': hashlib.sha256(result_path.read_bytes()).hexdigest(),
                            'seconds': r['seconds'], 'filter_checks': len(r['filters'])})
if not args.prepare:
    (args.root/'checks.json').write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n', encoding='utf8')
    print(json.dumps(report, ensure_ascii=False, indent=2))
