from pathlib import Path
import argparse, json, csv, math
ROOT=Path(__file__).resolve().parent
parser=argparse.ArgumentParser()
parser.add_argument('--output',type=Path,default=ROOT.parents[1]/'outputs/20261009_literature_validation')
parser.add_argument('--fixtures',type=Path,default=ROOT/'fixtures.json')
args=parser.parse_args()
OUT=args.output
fixtures={f['name']:f for f in json.loads(args.fixtures.read_text(encoding='utf-8-sig'))['fixtures']}
def equivalent_yield_strength(fs,phi):
 # For these single SOIL-material fixtures only: psi0=0, ReduceStrength=True.
 # beta=cos(atan(tan(phi0)/Fs)); therefore Fs/beta=sqrt(Fs²+tan²(phi0)).
 # This equates c and tan(phi), NOT the flow rule or the physical FOS.
 return math.sqrt(fs*fs+math.tan(math.radians(phi))**2)
rows=[]
trial_rows=[]
published_displacement={.8:.379,1.:.381,1.2:.422,1.3:.453,1.35:.544,1.4:1.476}
for path in sorted((OUT/'results').glob('*.json')):
 d=json.loads(path.read_text(encoding='utf-8-sig'))
 if 'metrics' not in d:continue
 m=dict(zip(d['metric_names'],d['metrics']));ref=d['reference']
 row={'case':d['case'],'status':d['status'],'flow_policy':d['flow_policy'],'elements':m['elements'],'nodes':m['nodes'],'seconds':d['elapsed_seconds'],'fos_pass':m['fos_pass'],'fos_fail':m['fos_fail'],'fos_mid':m['fos_mid'],'fos_bracket':m['fos_bracket'],'relative_residual':m['relative_residual'],'reaction_x':m['reaction_x'],'reaction_y':m['reaction_y']}
 if ref['source_id']=='elastic_validation':
  actual={int(v[0]):v[1:] for v in zip(*[iter(d['displacements_flat'])]*3)}
  err=max(abs(actual[n][i]-[ux,uy][i]) for n,ux,uy in ref['expected_displacements'] for i in range(2))
  scale=max(abs(x) for n,ux,uy in ref['expected_displacements'] for x in [ux,uy])
  actual_s={(int(e),int(g)):[sx,sy,st] for e,g,sx,sy,st in zip(*[iter(d['stresses_flat'])]*5)}
  # Stored FEM stresses use compression-positive tensor signs. A conventional
  # positive applied horizontal traction therefore has negative stored tau_xy.
  es=max(abs(actual_s[(e,g)][i]-[sx,sy,-tau][i]) for e,g,sx,sy,tau in ref['expected_stress'] for i in range(3))
  ss=max(abs(v) for e,g,sx,sy,tau in ref['expected_stress'] for v in [sx,sy,tau])
  reaction_err=max(abs(m['reaction_x']-ref['expected_reaction_x']),abs(m['reaction_y']-ref['expected_reaction_y']))
  row.update(max_displacement_abs_error=err,max_displacement_relative_error=err/max(scale,1e-30),max_stress_abs_error=es,max_stress_relative_error=es/max(ss,1e-30),reaction_abs_error=reaction_err,analytical_check_pass=bool(d['status'].startswith('PASS|') and err<1e-8*max(1,scale) and es<1e-7*max(1,ss) and reaction_err<1e-7))
 else:
  fixture=fixtures[d['fixture']]
  phi=fixture['materials'][0][5]
  disp=list(zip(*[iter(d['displacements_flat'])]*3));umax=max(math.hypot(x,y) for n,x,y in disp)
  component=max(abs(v) for n,x,y in disp for v in [x,y])
  row.update(umax=component,umax_vector=umax,umax_over_slope_height=component/ref['H'],umax_vector_over_slope_height=umax/ref['H'])
  weight=fixture['checks']['area_sum']*fixture['materials'][0][4]*fixture['materials'][0][3]
  row.update(expected_self_weight=weight,vertical_reaction_balance_error=abs(m['reaction_y']-weight))
  if ref.get('fos_reported'):row.update(reference_fos=ref['fos_reported'],mid_difference_percent=(m['fos_mid']/ref['fos_reported']-1)*100)
  if ref.get('fos_reference_range'):row.update(reference_min=ref['fos_reference_range'][0],reference_max=ref['fos_reference_range'][1],range_overlap=bool(m['fos_bracket'] and m['fos_pass']<=ref['fos_reference_range'][1] and m['fos_fail']>=ref['fos_reference_range'][0]))
  if d['flow_policy']=='DAVIS' and m['fos_bracket']:
   row.update(equivalent_yield_strength_pass=equivalent_yield_strength(m['fos_pass'],phi),equivalent_yield_strength_fail=equivalent_yield_strength(m['fos_fail'],phi),equivalent_yield_strength_mid=equivalent_yield_strength(m['fos_mid'],phi))
  log=OUT/'cases'/(d['case']+'_out')/'perf_summary.csv'
  if log.exists():
   with log.open(encoding='utf-8-sig',newline='') as f:
    trials=list(csv.DictReader(f))
   for t in trials:
    # Native CSV rounds Fs to 3 decimals. Exact bracket endpoints come from
    # the result JSON, so do not infer finer search resolution from this field.
    t={'case':d['case'],'flow_policy':d['flow_policy'],**t}
    t['displacement_measure']='maximum absolute free-DOF component (native log)'
    t['equivalent_yield_strength_factor']=equivalent_yield_strength(float(t['fs']),phi) if d['flow_policy']=='DAVIS' else float(t['fs'])
    t['slope_height']=ref['H']
    t['umax_over_slope_height']=float(t['umax'])/ref['H']
    if ref['source_id']=='griffiths_lane_1999_ex1':
     fs=float(t['fs']);normalized=100000*float(t['umax'])/(20*ref['H']**2)
     t['normalized_displacement']=normalized
     if fs in published_displacement:
      t['published_normalized_displacement']=published_displacement[fs]
      t['normalized_displacement_difference_percent']=(normalized/published_displacement[fs]-1)*100
    trial_rows.append(t)
   limits=[t for t in trials if t['mechanical_status']=='LIMIT_STATE']
   row['limit_state_trial_count']=len(limits)
   row['converged_limit_state_trial_count']=sum(t['numerical_status']=='CONVERGED' for t in limits)
 rows.append(row)
fields=list(dict.fromkeys(k for row in rows for k in row))
with (OUT/'comparison.csv').open('w',encoding='utf-8-sig',newline='') as f:
 w=csv.DictWriter(f,fieldnames=fields);w.writeheader();w.writerows(rows)
(OUT/'comparison.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
if trial_rows:
 trial_fields=list(dict.fromkeys(k for row in trial_rows for k in row))
 with (OUT/'srm_trials.csv').open('w',encoding='utf-8-sig',newline='') as f:
  w=csv.DictWriter(f,fieldnames=trial_fields);w.writeheader();w.writerows(trial_rows)
for row in rows:print(row['case'],row.get('analytical_check_pass',row.get('fos_mid')),row.get('max_displacement_abs_error',''),row.get('max_stress_abs_error',''))
