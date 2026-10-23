"""Rebuild field-oriented plots from published, accepted calculation records."""
from pathlib import Path
import gzip
import hashlib
import json
import zipfile

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'docs/manual-field-evidence/20261024'
IMAGES = ROOT / 'docs/manuals/images/practice'
OUT.mkdir(parents=True, exist_ok=True)
IMAGES.mkdir(parents=True, exist_ok=True)
plt.rcParams.update({'font.family': FontProperties(fname='C:/Windows/Fonts/meiryo.ttc').get_name(),
                     'font.size': 10, 'axes.spines.top': False, 'axes.spines.right': False})
sources = {}


def source(path):
    p = ROOT / path
    sources[path] = hashlib.sha256(p.read_bytes()).hexdigest()
    return p


def records(name):
    p = source('docs/manuals/examples/' + name)
    r = json.loads(gzip.decompress(p.read_bytes()))
    assert r['status'] == 'PASS'
    return r['tables']


def node(table, stage, number):
    rows = [r for r in table['結果節点'][1:] if r[0] == stage and r[1] == number]
    assert len(rows) == 1, (stage, number)
    return rows[0]


def active_adjacency(table, stages, points):
    """Count result elements incident to each point at each stage.

    ResultElement is the authoritative active set for a saved stage; using the
    input connectivity only maps an active element to its incident nodes.  This
    deliberately leaves gaps during DEATH/late BIRTH intervals instead of
    connecting displacements across an inactive period.
    """
    elements = {int(r[0]): [int(x) for x in r[1:9]]
                for r in table['要素データ'][1:]}
    result_elements = table['結果要素'][1:]
    out = {}
    for p in points:
        out[str(p)] = [sum(1 for r in result_elements
                            if int(r[0]) == stage
                            and p in elements.get(int(r[1]), []))
                       for stage in stages]
    return out


def contiguous_segments(valid):
    segments = []
    start = None
    for i, ok in enumerate(valid, 1):
        if ok and start is None:
            start = i
        if start is not None and (not ok or i == len(valid)):
            end = i if ok else i - 1
            segments.append((start, end))
            start = None
    return segments


def replay_active_elements(table, stages):
    """Replay stage material operations and return active element IDs by stage."""
    elements = {int(r[0]): int(r[9]) for r in table['要素データ'][1:]}
    materials = {int(r[0]): str(r[9]).upper() == 'ON' for r in table['材料データ'][1:]}
    active = {eid for eid, mat in elements.items() if materials.get(mat, False)}
    commands = {int(r[0]): r for r in table['ステージ'][1:]}
    out = {}
    result_elements = table['結果要素'][1:]
    for stage in stages:
        row = commands[stage]
        kind, material, parameter = str(row[1]).upper(), row[2], row[3]
        if kind == 'BIRTH' and material is not None:
            active.update(eid for eid, mat in elements.items() if mat == int(material))
        elif kind == 'DEATH' and material is not None:
            active.difference_update(eid for eid, mat in elements.items() if mat == int(material))
        elif kind == 'MATSET' and material is not None and parameter is not None:
            # MATSET copies properties into the target material slot; it does
            # not change element MatNo/group membership or active flags.
            assert int(material) in materials and int(parameter) in materials
        expected = {int(r[1]) for r in result_elements if int(r[0]) == stage}
        assert active == expected, (stage, sorted(active ^ expected))
        out[str(stage)] = sorted(active)
    return out


def save(fig, name):
    fig.savefig(IMAGES / name, dpi=180)
    plt.close(fig)


road = records('ex08_road_result_20261021.json.gz')
road_stages = list(range(1, 12))
road_adjacency = active_adjacency(road, road_stages, [20, 19])
road_replay = replay_active_elements(road, road_stages)
assert road_adjacency['20'] == [2, 2, 4, 4, 2, 0, 0, 0, 2, 4, 4]
assert all(v > 0 for v in road_adjacency['19'])
road_rows = []
road_reference_node = 19
road_reference = node(road, 1, road_reference_node)[5]
for row in road['ステージ結果'][1:]:
    assert row[3] == 'PASS'
    sid = int(row[0])
    r = node(road, sid, road_reference_node)
    old_r = node(road, sid, 20)
    road_rows.append({'stage': sid, 'kind': row[1], 'reaction_y_kN': row[7],
                      'monitor_node': road_reference_node,
                      'monitor_node_adjacency': road_adjacency[str(road_reference_node)][sid - 1],
                      'point19_additional_settlement_mm': -(r[5] - road_reference)*1000,
                      'point20_adjacency': road_adjacency['20'][sid - 1],
                      'point20_additional_settlement_mm': -(old_r[5] - node(road, 1, 20)[5])*1000 if road_adjacency['20'][sid - 1] else None})
expected = {1: 1440, 4: 1728, 6: 1424, 11: 1842}
for row in road_rows:
    if row['stage'] in expected:
        assert abs(row['reaction_y_kN'] - expected[row['stage']]) < 1e-6
fig, axs = plt.subplots(2, 1, figsize=(9, 5.4), sharex=True, layout='constrained')
xs = [r['stage'] for r in road_rows]
axs[0].plot(xs, [r['reaction_y_kN'] for r in road_rows], 'o-', color='#356b68')
axs[0].set(ylabel='鉛直反力 [kN/m]', title='道路拡幅：初回盛土 → 掘削 → 別材料で再盛土')
for sid, ry in expected.items():
    axs[0].annotate(f'{ry}', (sid, ry), xytext=(0, 8), textcoords='offset points', ha='center')
axs[1].plot(xs, [r['point19_additional_settlement_mm'] for r in road_rows], 'o-', color='#b76e32')
axs[1].set(ylabel='恒久監視点19の追加沈下 [mm]', xlabel='工程ID（工程1の自重後を変位の基準）', xticks=xs)
for ax in axs:
    ax.axvspan(4.5, 6.5, color='#cfdde6', alpha=.4)
    ax.axvspan(8.5, 11.3, color='#e6d3ab', alpha=.35)
    ax.grid(alpha=.2)
save(fig, 'road_stage_monitoring.png')

tunnel = records('ex10_tunnel_result_20261021.json.gz')
points = {'crown': 330, 'invert': 114, 'left': 438, 'right': 222, 'surface': 351}
coords = {'crown': (0, -20), 'invert': (0, -28), 'left': (-4, -24), 'right': (4, -24), 'surface': (0, 0)}
for key, number in points.items():
    assert tuple(node(tunnel, 1, number)[2:4]) == coords[key]
def du(stage, key, component, reference_stage=1):
    number = points[key]
    return node(tunnel, stage, number)[component] - node(tunnel, reference_stage, number)[component]
tunnel_stages = list(range(1, 10))
tunnel_adjacency = active_adjacency(tunnel, tunnel_stages, list(points.values()))
tunnel_replay = replay_active_elements(tunnel, tunnel_stages)
assert tunnel_adjacency['330'] == [4, 2, 0, 0, 2, 2, 2, 2, 2]
assert tunnel_adjacency['114'] == [4, 4, 4, 4, 4, 2, 0, 0, 2]
assert tunnel_adjacency['351'] == [2, 2, 2, 2, 2, 2, 2, 2, 2]
tunnel_vertical_valid = [a > 0 and b > 0 for a, b in zip(tunnel_adjacency['330'], tunnel_adjacency['114'])]
tunnel_horizontal_valid = [a > 0 and b > 0 for a, b in zip(tunnel_adjacency['438'], tunnel_adjacency['222'])]
tunnel_vertical_segments = contiguous_segments(tunnel_vertical_valid)
tunnel_horizontal_segments = contiguous_segments(tunnel_horizontal_valid)
tunnel_vertical_reference = {}
for first, last in tunnel_vertical_segments:
    for sid in range(first, last + 1):
        tunnel_vertical_reference[sid] = first
tunnel_horizontal_reference = {}
for first, last in tunnel_horizontal_segments:
    for sid in range(first, last + 1):
        tunnel_horizontal_reference[sid] = first
tunnel_rows = []
for sid in range(1, 10):
    vref = tunnel_vertical_reference.get(sid)
    href = tunnel_horizontal_reference.get(sid)
    vertical = None if vref is None else -(du(sid, 'crown', 5, vref)-du(sid, 'invert', 5, vref))*1000
    horizontal = None if href is None else -(du(sid, 'right', 4, href)-du(sid, 'left', 4, href))*1000
    tunnel_rows.append({'stage': sid,
                        'vertical_closure_mm': vertical,
                        'horizontal_closure_mm': horizontal,
                        'vertical_valid': bool(vref is not None),
                        'vertical_reference_stage': vref,
                        'horizontal_valid': bool(href is not None),
                        'horizontal_reference_stage': href,
                        'surface_settlement_mm': -du(sid, 'surface', 5)*1000})
check = json.loads(source('docs/manuals/construction-evidence/20261020/tunnel10/construction_displacements.json').read_text())
raw_vertical = -(du(9, 'crown', 5, 1)-du(9, 'invert', 5, 1))*1000
raw_horizontal = -(du(9, 'right', 4, 1)-du(9, 'left', 4, 1))*1000
raw_surface = -du(9, 'surface', 5, 1)*1000
for key in ('vertical_closure_mm', 'horizontal_closure_mm'):
    assert abs({'vertical_closure_mm': raw_vertical,
                'horizontal_closure_mm': raw_horizontal}[key] - check[key]) < 1e-9
assert abs(raw_surface-check['surface_settlement_increment_mm']) < 1e-9
fig, ax = plt.subplots(figsize=(9, 3.0), layout='constrained')
for key, label, color in [('vertical_closure_mm','内空高さの変化（有効区間ごとに基準更新）','#356b68'),
                          ('horizontal_closure_mm','内空幅の減少（負：拡大）','#b76e32'),
                          ('surface_settlement_mm','地表中央の追加沈下','#5964a0')]:
    ax.plot(range(1, 10), [float('nan') if r[key] is None else r[key] for r in tunnel_rows],
            'o-', label=label, color=color)
ax.axvspan(1.5, 5.5, alpha=.2, color='#cfdde6')
ax.axvspan(5.5, 9.3, alpha=.2, color='#e6d3ab')
ax.set(xticks=range(1, 10), xlabel='工程ID（1：初期自重、2～5：上半、6～9：下半）',
       ylabel='基準工程からの変化 [mm]', title='トンネル：有効区間ごとに基準を更新して追跡')
for sid in sorted(set(r['vertical_reference_stage'] for r in tunnel_rows if r['vertical_reference_stage'] is not None)):
    ax.axvline(sid, color='#356b68', lw=.7, ls=':', alpha=.6)
    ax.annotate(f'基準{sid}', (sid, 0), xytext=(3, 7), textcoords='offset points', fontsize=8, color='#356b68')
ax.legend(loc='upper center', bbox_to_anchor=(.5, -.24), ncol=3, frameon=False, fontsize=9)
ax.grid(alpha=.2)
save(fig, 'tunnel_stage_monitoring.png')

archive = source('docs/groundwater-evidence/20261022/records.zip')
boxes = []
with zipfile.ZipFile(archive) as z:
    for name, theory in [('box', 102), ('box_inside', 107)]:
        r = json.loads(z.read(name + '/result.json'))
        assert r['status'].startswith('PASS|')
        assert abs(r['metrics'][2]-theory) < 1e-6
        boxes.append({'case': name, 'weight_kN': 192, 'external_uplift_kN': 90,
                      'internal_downward_water_kN': theory-102, 'theory_reaction_kN': theory,
                      'calculated_reaction_kN': r['metrics'][2],
                      'absolute_error_kN': abs(r['metrics'][2]-theory)})
fig, ax = plt.subplots(figsize=(9, 3.3), layout='constrained')
for i, r in enumerate(boxes):
    vals = [192, -90, r['internal_downward_water_kN'], r['calculated_reaction_kN']]
    positions = [j+i*.34 for j in range(4)]
    ax.bar(positions, vals, width=.3, color=['#356b68','#b76e32'][i], label=['内面DRY','内水位1.5 m'][i])
    for x, y in zip(positions, vals):
        ax.text(x,y+(3 if y>=0 else -10),f'{y:.0f}',ha='center',fontsize=9)
ax.set_xticks([j+.17 for j in range(4)], ['構造体自重','外水圧（上向き）','内水（下向き）','計算支持反力'])
ax.set(ylabel='鉛直力 [kN/m]', title='地下ボックス：自重 − 外水圧 ＋ 内水荷重 ＝ 支持反力')
ax.axhline(0, color='#555', lw=.8)
ax.legend(frameon=False, loc='upper right')
save(fig, 'box_water_balance.png')

cutoff = next(r for r in json.loads(source('docs/groundwater-evidence/20261022/independent.json').read_text()) if r['case']=='cutoff')
assert abs(cutoff['inflow_m3_s']-cutoff['outflow_m3_s']) < 1e-15
result = {'status':'PASS', 'scope':'Independent reaggregation of published accepted results; active adjacency is derived from stage ResultElement and input connectivity',
          'sources_sha256':sources, 'road':road_rows,
          'road_replay_active_elements': road_replay,
          'road_monitoring': {'reference_node': road_reference_node,
                              'reference_coordinates': [4, 0],
                              'reference_reason': 'node19 is active in all 11 stages and is near the former node20',
                              'node20_adjacency_by_stage': road_adjacency['20'],
                              'reference_node_adjacency_by_stage': road_adjacency['19']},
          'tunnel': tunnel_rows,
          'tunnel_raw_stage1_reference_final': {'vertical_closure_mm': raw_vertical,
                                                'horizontal_closure_mm': raw_horizontal,
                                                'surface_settlement_mm': raw_surface},
          'tunnel_points':points,
          'tunnel_adjacency_by_stage': tunnel_adjacency,
          'tunnel_replay_active_elements': tunnel_replay,
          'tunnel_closure_reference_segments': {'vertical': [list(s) for s in tunnel_vertical_segments],
                                                 'horizontal': [list(s) for s in tunnel_horizontal_segments]},
          'tunnel_closure_validity': {'vertical': tunnel_vertical_valid,
                                      'horizontal': tunnel_horizontal_valid},
          'box':boxes, 'cutoff':cutoff}
(OUT/'field_examples.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print('PASS: road11, tunnel9, box2, cutoff balance; three figures')
