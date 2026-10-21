"""Independent comparison and figures for the taller road exercise."""
from pathlib import Path
import gzip, json, sys
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
from check_manual_case import N
root=Path(__file__).resolve().parents[1]
folder=Path(sys.argv[1]) if len(sys.argv)>1 else root/'tests/tmp/manual_applied/road_taller'
def read(p):
    if not p.exists():p=p.with_suffix(p.suffix+'.gz')
    return json.loads(gzip.decompress(p.read_bytes()) if p.suffix=='.gz' else p.read_text(encoding='utf-8-sig'))
base=read(root/'docs/manuals/construction-evidence/20261021/manual-review/road_ui05/result.json.gz')
data=read(folder/'result.json');checks=read(folder/'independent_checks.json');balance=read(folder/'balance_checks.json')
assert data['status']==base['status']=='PASS' and data['application_code_unchanged']
fixture=read(root/'tests/manual_cases/road_widening_taller_shape.json');old=read(root/'tests/manual_cases/road_widening_shape.json')
expected=json.loads(json.dumps(old));expected['description']=fixture['description'];expected['expected_region_areas'][-1]=9
for row in expected['points']:
    if row[0] in (17,18):row[2]=5
assert expected==fixture
for name in ('材料データ','ステージ','載荷'):assert base['tables'][name]==data['tables'][name],name
def settings(d):return {r[4]:r[2] for r in d['tables']['設定'][1:] if r[4]}
u,v=settings(base),settings(data)
differences={k:[u[k],v.get(k)] for k in u if u[k]!=v.get(k)}
assert set(differences)<= {'MESH_GEN_LAST_METHOD','MESH_GEN_LAST_COMPARE'},differences
assert checks['deformed_geometry_valid'] and checks['fixed_disp_max']==0
assert abs(balance['material_areas_m2']['8']-9)<1e-9
for i,area in enumerate(fixture['expected_region_areas'],1):assert abs(balance['material_areas_m2'][str(i)]-area)<1e-8
for node in data['tables']['節点データ'][1:]:
    x,y,fix_x,fix_y=node[1:5]
    assert int(fix_x or 0)==int(abs(y+6)<1e-9 or abs(x)<1e-9 or abs(x-12)<1e-9)
    assert int(fix_y or 0)==int(abs(y+6)<1e-9)
assert len(balance['stage_balance'])==11
assert all(row[3]=='PASS' and row[4]==1 for row in data['tables']['ステージ結果'][1:])
assert abs(balance['stage_balance'][0]['reaction_y_kN']-1440)<1e-8
assert abs(balance['stage_balance'][-1]['reaction_y_kN']-1899)<1e-8
assert max(abs(r['global_equilibrium_closure_y_kN']) for r in balance['stage_balance'])<1e-7
def metrics(d):
    t=d['tables'];rows=[r for r in t['結果節点'][1:] if r[0]==11]
    stage=t['ステージ結果'][-1]
    return dict(nodes=d['nodes'],elements=d['elements'],reaction_y_kN=stage[7],max_displacement_m=max(float(np.hypot(r[4],r[5])) for r in rows),seconds=d['seconds'])
summary=dict(status='PASS',baseline=metrics(base),taller=metrics(data),expected_added_area_m2=3,expected_added_weight_kN=57,observed_added_reaction_kN=checks['reaction_y_kN']-metrics(base)['reaction_y_kN'],last_relative_residual=checks['relative_residual'],min_reference_jacobian_11x11=checks['dense_11x11_jacobian']['reference'],min_deformed_jacobian_11x11=checks['dense_11x11_jacobian']['deformed'],fixed_displacement_max=checks['fixed_disp_max'],max_global_closure_kN=max(abs(r['global_equilibrium_closure_y_kN']) for r in balance['stage_balance']),settings_differences=differences,settings_unchanged_except_mesh_metadata=True,quality_note=v.get('MESH_GEN_LAST_COMPARE'),limitations='Angle quality WARN remains: 99 pass / 4 warn / 0 fail; no SRM or mesh-convergence claim. Runtime is a single observation, not a performance comparison.')
if '--check-only' in sys.argv:
    print('PASS published taller road: 11 stages, 1899 kN, +57 kN, material areas/supports/settings');sys.exit()
(folder/'exercise_summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
fig,axes=plt.subplots(1,2,figsize=(10,4),sharex=True,sharey=True)
for ax,d,title in zip(axes,[base,data],['Before: 328 nodes / 95 Q8','After: 352 nodes / 103 Q8']):
    nodes={int(r[0]):r for r in d['tables']['節点データ'][1:]};elements=d['tables']['要素データ'][1:]
    t=np.linspace(-1,1,9);edge=[(r,-1) for r in t]+[(1,s) for s in t]+[(r,1) for r in t[::-1]]+[(-1,s) for s in t[::-1]]
    polys=[np.array([N(r,s)@np.array([[nodes[int(i)][1],nodes[int(i)][2]] for i in e[1:9]]) for r,s in edge]) for e in elements]
    ax.add_collection(PolyCollection(polys,array=np.array([e[9] for e in elements]),cmap='Pastel1',clim=(1,9),edgecolors='#777777',linewidths=.45))
    for row in nodes.values():
        if row[3] or row[4]:ax.plot(row[1],row[2],'x',color='#ab3838',ms=2)
    ax.autoscale();ax.set_aspect('equal');ax.set_title(title,fontsize=10);ax.set_xlabel('X / m');ax.grid(alpha=.15)
axes[0].set_ylabel('Y / m');fig.tight_layout();fig.savefig(folder/'geometry_before_after.png',dpi=180);plt.close(fig)
print('PASS taller road: 11 stages, 1899 kN, +57 kN; Double/convergence settings unchanged')
