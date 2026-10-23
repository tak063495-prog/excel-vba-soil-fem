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


def save(fig, name):
    fig.savefig(IMAGES / name, dpi=180)
    plt.close(fig)


road = records('ex08_road_result_20261021.json.gz')
road_rows = []
for row in road['ステージ結果'][1:]:
    assert row[3] == 'PASS'
    sid = int(row[0])
    r = node(road, sid, 20)
    road_rows.append({'stage': sid, 'kind': row[1], 'reaction_y_kN': row[7],
                      'point20_additional_settlement_mm': -(r[5] - node(road, 1, 20)[5])*1000})
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
axs[1].plot(xs, [r['point20_additional_settlement_mm'] for r in road_rows], 'o-', color='#b76e32')
axs[1].set(ylabel='節点20の追加沈下 [mm]', xlabel='工程ID（工程1の自重後を変位の基準）', xticks=xs)
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
def du(stage, key, component):
    number = points[key]
    return node(tunnel, stage, number)[component] - node(tunnel, 1, number)[component]
tunnel_rows = []
for sid in range(1, 10):
    tunnel_rows.append({'stage': sid,
                        'vertical_closure_mm': -(du(sid, 'crown', 5)-du(sid, 'invert', 5))*1000,
                        'horizontal_closure_mm': -(du(sid, 'right', 4)-du(sid, 'left', 4))*1000,
                        'surface_settlement_mm': -du(sid, 'surface', 5)*1000})
check = json.loads(source('docs/manuals/construction-evidence/20261020/tunnel10/construction_displacements.json').read_text())
for key in ('vertical_closure_mm', 'horizontal_closure_mm'):
    assert abs(tunnel_rows[-1][key] - check[key]) < 1e-9
assert abs(tunnel_rows[-1]['surface_settlement_mm']-check['surface_settlement_increment_mm']) < 1e-9
fig, ax = plt.subplots(figsize=(9, 3.6), layout='constrained')
for key, label, color in [('vertical_closure_mm','内空高さの減少','#356b68'),
                          ('horizontal_closure_mm','内空幅の減少（負：拡大）','#b76e32'),
                          ('surface_settlement_mm','地表中央の追加沈下','#5964a0')]:
    ax.plot(range(1, 10), [r[key] for r in tunnel_rows], 'o-', label=label, color=color)
ax.axvspan(1.5, 5.5, alpha=.2, color='#cfdde6')
ax.axvspan(5.5, 9.3, alpha=.2, color='#e6d3ab')
ax.set(xticks=range(1, 10), xlabel='工程ID（1：初期自重、2～5：上半、6～9：下半）',
       ylabel='工程1からの変化 [mm]', title='トンネル：同じ測点の相対変位を施工順に追う')
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
result = {'status':'PASS', 'scope':'Independent reaggregation of published accepted results; figures retain original model/stage/point definitions',
          'sources_sha256':sources, 'road':road_rows, 'tunnel':tunnel_rows,
          'tunnel_points':points, 'box':boxes, 'cutoff':cutoff}
(OUT/'field_examples.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
print('PASS: road11, tunnel9, box2, cutoff balance; three figures')
