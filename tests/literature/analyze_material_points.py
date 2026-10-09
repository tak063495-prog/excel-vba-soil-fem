import json,math,argparse
from pathlib import Path
parser=argparse.ArgumentParser();parser.add_argument('--output',type=Path,required=True);args=parser.parse_args()
d=json.loads((args.output/'material_points_raw.json').read_text(encoding='utf-8-sig'));rows=[]
E=100000.;nu=.3;c=10.;K=E/(3*(1-2*nu));G=E/(2*(1+nu))
for r in d['records']:
    v=dict(zip(d['value_names'],r['values']));sin=math.sin(math.radians(r['phi']));cos=math.cos(math.radians(r['phi']))
    if r['gxy']:
        tau=min(G*r['gxy'],c);expected=[0.,0.,tau,0.];plastic=G*r['gxy']>c
    else:
        # At zero dilation, mean stress p remains elastic. The symmetric edge
        # return has sigma_x=sigma_z and q=sigma_y-sigma_x. The MC surface gives
        # q=(2*c*cos(phi)+2*p*sin(phi))/(1-sin(phi)/3).
        p=K*r['ey'];qt=2*G*r['ey'];qy=(2*c*cos+2*p*sin)/(1-sin/3)
        q=min(qt,qy);expected=[p-q/3,p+2*q/3,0.,p-q/3];plastic=qt>qy
    actual=[v[k] for k in ['sx','sy','tau','sz']];error=max(abs(a-b) for a,b in zip(actual,expected))
    volume=v['epx']+v['epy']+v['epz']
    passed=bool(v['ok'] and v['converged'] and v['failure_code']==0 and v['elastic']==(not plastic) and error<1e-6 and abs(volume)<1e-10)
    result=dict(case=r['case'],phi=r['phi'],expected_stress=expected,actual_stress=actual,stress_max_abs_error=error,plastic_volume_strain=volume,expected_plastic=plastic,ok=passed)
    if 't00' in v:
        a,b,cross,d0=[v[k] for k in ['t00','t01','t10','t11']]
        result['tangent_3x3']=[[v[f't{i}{j}'] for j in range(3)] for i in range(3)]
        result['tangent_asymmetry_abs']=max(abs(v[f't{i}{j}']-v[f't{j}{i}']) for i in range(3) for j in range(3))
        result['symmetrized_normal_block_min_eigenvalue']=(a+d0-math.sqrt((a-d0)**2+(b+cross)**2))/2
    result['tangent_probes']=[]
    for probe in r.get('tangent_probes',[]):
        base=dict(zip(d['value_names'],probe['base_values']))
        returned=[[base[f't{i}{j}'] for j in range(3)] for i in range(3)]
        fd=[[probe['fd_columns'][j][i] for j in range(3)] for i in range(3)]
        norm=math.sqrt(sum(x*x for row in fd for x in row))
        gap=math.sqrt(sum((fd[i][j]-returned[i][j])**2 for i in range(3) for j in range(3)))
        branches=[]
        for b0 in probe.get('branch_probes',[]):
            plus=dict(zip(d['value_names'],b0['plus_values']))
            minus=dict(zip(d['value_names'],b0['minus_values']))
            keys=['sx','sy','tau'];h=probe['h']
            forward=[(plus[k]-base[k])/h for k in keys]
            backward=[(base[k]-minus[k])/h for k in keys]
            branches.append(dict(column=b0['column'],plus_elastic=plus['elastic'],minus_elastic=minus['elastic'],plus_message=plus['failure_message'],minus_message=minus['failure_message'],forward_derivative=forward,backward_derivative=backward))
        result['tangent_probes'].append(dict(name=probe['name'],probe_ok=probe['ok'],h=probe['h'],returned=returned,finite_difference=fd,relative_frobenius_gap=gap/max(norm,1e-30),branches=branches))
    rows.append(result)
independent_ok=all(r['ok'] for r in rows)
self_failures=[list(v) for v in zip(*[iter(d['self_test_table_flat'])]*12) if v[8]=='FAIL']
report=dict(source_sha256=d.get('source_sha256'),ok=independent_ok and d['built_in_material_self_test'] and d['built_in_edge_self_test'],ok_scope='Independent stress/volume checks and built-in self-tests only; tangent probes are diagnostics, not an acceptance test.',independent_ok=independent_ok,records=rows,built_in_material_self_test=d['built_in_material_self_test'],built_in_edge_self_test=d['built_in_edge_self_test'],selftest_inputs_trial_repair=d.get('selftest_inputs_trial_repair',False),built_in_failure_rows=self_failures)
(args.output/'material_point_checks.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
for r in rows:print(r['case'],r['ok'],r['stress_max_abs_error'])
if not report['ok']:raise SystemExit(1)
