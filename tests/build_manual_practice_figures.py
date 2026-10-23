"""Plot archived road-widening results; do not rerun or alter the model."""
import json
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties

root = Path(__file__).resolve().parents[1]
data = json.loads((root / 'docs/manuals/construction-evidence/20261021/manual-applied/exercise_summary.json').read_text(encoding='utf-8'))
assert data['status'] == 'PASS'
assert data['observed_added_reaction_kN'] == data['expected_added_weight_kN'] == 57
font = FontProperties(fname='C:/Windows/Fonts/meiryo.ttc')
plt.rcParams['font.family'] = font.get_name()
fig, axes = plt.subplots(1, 2, figsize=(9, 3.5), layout='constrained')
for ax, key, factor, title, unit in [
    (axes[0], 'reaction_y_kN', 1, '最終鉛直反力', 'kN/m'),
    (axes[1], 'max_displacement_m', 1000, '最大変位ノルム', 'mm'),
]:
    values = [data[case][key]*factor for case in ('baseline', 'taller')]
    ax.bar(['基準', '天端 +1 m'], values, color=['#356b68', '#cf8a3c'], width=.5)
    ax.set_title(title)
    ax.set_ylabel(unit)
    ax.set_ylim(0, max(values)*1.18)
    ax.grid(axis='y', alpha=.2)
    ax.set_axisbelow(True)
    for i, value in enumerate(values):
        ax.text(i, value+max(values)*.025, f'{value:.3f}' if factor == 1000 else f'{value:.0f}', ha='center', fontsize=11)
    ax.spines[['top', 'right']].set_visible(False)
out = root / 'docs/manuals/images/practice/road_comparison.png'
out.parent.mkdir(parents=True, exist_ok=True)
fig.savefig(out, dpi=180)
plt.close(fig)
print(out)
