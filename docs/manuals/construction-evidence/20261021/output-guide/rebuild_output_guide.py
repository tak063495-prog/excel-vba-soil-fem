from pathlib import Path
import json,csv,hashlib,gzip,io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
plt.rcParams['font.family']='Yu Gothic';plt.rcParams['axes.unicode_minus']=False
root=Path.cwd();out=Path(__file__).resolve().parent
road=json.loads((root/'docs/manuals/construction-evidence/20261021/road01/display_stage_checks.json').read_text(encoding='utf8'))
with (out/'road_stage_summary.csv').open('w',encoding='utf8',newline='') as f:
    w=csv.writer(f);w.writerow(['stage','max_displacement_norm_m','reaction_y_kN','accepted_increment_relative_residual'])
    for x in road:w.writerow([x['stage'],x['max_displacement_m'],x['reaction_y_kN'],x['increment_log_relative_residual']])
fig,ax=plt.subplots(1,2,figsize=(10,3.8));stage=[x['stage'] for x in road]
ax[0].plot(stage,[1000*x['max_displacement_m'] for x in road],'o-');ax[0].set_ylabel('全節点の最大変位ノルム [mm]')
ax[1].plot(stage,[x['reaction_y_kN'] for x in road],'s-',color='#c57a20');ax[1].set_ylabel('支持点の鉛直反力 [kN]')
for a in ax:a.set_xlabel('施工工程');a.set_xticks(stage);a.grid(alpha=.25)
fig.tight_layout();fig.savefig(out/'output_road_summary.png',dpi=190);plt.close(fig)
base=root/'docs/manuals/construction-evidence/20261020/wall_davis08'
def read_csv(name):return list(csv.DictReader(io.StringIO(gzip.decompress((base/(name+'.gz')).read_bytes()).decode('utf8'))))
perf=read_csv('perf_summary.csv')
inc=read_csv('increment_summary.csv')
trials={}
for x in inc:
    if int(x['srm_stage'])==4:trials[int(x['trial_id'])]=float(x['fs'])
assert len(perf)==len({int(x['trial_id']) for x in inc})
alltrials=sorted({int(x['trial_id']) for x in inc});rows=[]
for trial,p in zip(alltrials,perf):
    if trial in trials:
        fs=trials[trial];assert abs(fs-float(p['fs']))<=.00051
        rows.append(dict(trial_id=trial,fs=fs,numerical_status=p['numerical_status'],umax_m=float(p['umax']),relative_residual=float(p['relres']),elapsed_s=float(p['elapsed_s'])))
assert len(rows)==12
with (out/'srm_trial_summary.csv').open('w',encoding='utf8',newline='') as f:
    w=csv.DictWriter(f,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)
fig,ax=plt.subplots(1,2,figsize=(10,3.9))
for status,marker,col,label in [('CONVERGED','o','#197d65','収束'),('FAILED','x','#cf653d','非収束（途中値）')]:
    rr=[r for r in rows if r['numerical_status']==status]
    ax[0].scatter([r['fs'] for r in rr],[r['umax_m']*1000 for r in rr],marker=marker,color=col,label=label)
    ax[1].scatter([r['fs'] for r in rr],[r['relative_residual'] for r in rr],marker=marker,color=col,label=label)
for a in ax:a.set_xlabel('SRM試行のFs');a.grid(alpha=.25);a.axvspan(1.7,1.7125,color='#e2b267',alpha=.18);a.legend(fontsize=8)
ax[0].set_ylabel('CSVのumax [mm]');ax[1].set_ylabel('CSVの相対残差 relres');ax[1].set_yscale('log')
fig.tight_layout();fig.savefig(out/'output_srm_summary.png',dpi=190);plt.close(fig)
print('PASS road4 stages, SRM12 trials matched by trial_id and srm_stage=4')
