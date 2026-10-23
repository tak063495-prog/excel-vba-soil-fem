"""Compare actual accepted, restored and following-WAIT hydraulic states."""
from pathlib import Path
import argparse,json,math

def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def close(a,b,tol=1e-9):
 if isinstance(a,bool) or isinstance(b,bool):
  assert type(a) is type(b) and a==b,(a,b);return 0.
 if isinstance(a,(int,float)) and isinstance(b,(int,float)):
  assert math.isfinite(a) and math.isfinite(b),(a,b)
  d=abs(a-b);assert d<=tol,(d,tol,a,b);return d
 if isinstance(a,list):
  assert isinstance(b,list) and len(a)==len(b) and len(a)>0
  return max(close(x,y,tol) for x,y in zip(a,b))
 assert isinstance(a,str) and isinstance(b,str) and a==b,(a,b);return 0.
FIELDS={'profile':0,'external_load':1e-9,'boundary_locked':0,'boundary_pressure':1e-8,'heads':1e-9,'flow':1e-12,'flow_reaction':1e-12,'gp_pressure':1e-8}
def state(a,b):
 for k in FIELDS:
  if k=='profile':assert isinstance(a[k],str) and a[k] and isinstance(b[k],str) and b[k]
  else:assert isinstance(a[k],list) and a[k] and isinstance(b[k],list) and b[k],k
 return {k:close(a[k],b[k],tol) for k,tol in FIELDS.items()}
def check(root,cases):
 report={'status':'PASS','cases':[]}
 for suffix in cases:
  for checkpoint in ['mid','completed']:
   p=root/f'issue50_{suffix}_{checkpoint}';a=read(p/'result.json');b=read(p/'resume.json')
   assert a['status'].startswith('PASS|') and b['status'].startswith('PASS|')
   if a['requested_native']=='OFF':assert 'VBA(OFF_SELECTED)' in a['native'],a['native']
   else:assert a['requested_native']=='AUTO' and a['native'].startswith('NATIVE('),a['native']
   du=close(a['disp'],b['disp'],1e-9);metrics=close(a['metrics'],b['metrics'],1e-6)
   fresh=read(p/'water_observables.json');resumed=read(p/'resume_water_observables.json')
   final=state(fresh,resumed)
   loaded=resumed['loaded'];assert loaded['stage']==3 and loaded['duration']==40
   if checkpoint=='mid':assert 0<loaded['elapsed']<40
   else:assert loaded['elapsed']==40
   matching=[v for v in fresh['frames'] if v['stage']==3 and v['elapsed']==loaded['elapsed']]
   assert len(matching)==1
   load_delta=state(matching[0],loaded)
   wait=[]
   for v in resumed['frames']:
    if v['stage']!=4:continue
    candidates=[w for w in fresh['frames'] if w['stage']==4 and w['elapsed']==v['elapsed']]
    assert len(candidates)==1;wait.append({'elapsed':v['elapsed'],'delta':state(candidates[0],v)})
   assert len(wait)==2
   model=read(Path('tests/fixtures/issue50')/f'issue50_{suffix}.json')
   expected_profile='W' if suffix in ['explicit_hold','explicit_blank_mode_hold'] else 'W2'
   assert fresh['profile']==resumed['profile']==loaded['profile']==expected_profile
   levels=[r[2] for r in model['water']['profiles'] if r[0]==expected_profile];assert levels and max(levels)==min(levels)
   width=max(n[1] for n in model['nodes'])-min(n[1] for n in model['nodes']);top=max(n[2] for n in model['nodes'])
   expected_force=-model['water']['gamma']*(levels[0]-top)*width
   close(sum(fresh['external_load'][1::2]),expected_force,1e-8)
   close(sum(resumed['external_load'][1::2]),expected_force,1e-8)
   close(a['metrics'][2],a['metrics'][0]-expected_force,1e-6)
   # Distinguish a genuine named-profile change from a shared stale answer.
   if suffix=='default_start_no_detail':assert sum(resumed['external_load'][1::2])< -25
   report['cases'].append({'case':suffix,'checkpoint':checkpoint,'native':a['requested_native'],'book_sha256':a['book_sha256'],'status':'PASS','dofs':len(a['disp']),'max_u_m':du,'max_metrics_delta':metrics,'restored_elapsed_s':loaded['elapsed'],'restored_delta':load_delta,'final_delta':final,'following_WAIT':wait,'final_profile':resumed['profile'],'final_surface_force_y_kN':sum(resumed['external_load'][1::2]),'fresh_seconds':a['seconds'],'resume_seconds':b['seconds']})
 if 'default_start_no_detail' in cases and 'explicit_start' in cases:
  a=read(root/'issue50_default_start_no_detail_completed/result.json');b=read(root/'issue50_explicit_start_completed/result.json')
  report['default_vs_explicit_start']={'max_u_m':close(a['disp'],b['disp'],1e-12),
   'state':state(read(root/'issue50_default_start_no_detail_completed/water_observables.json'),read(root/'issue50_explicit_start_completed/water_observables.json'))}
 return report
if __name__=='__main__':
 ap=argparse.ArgumentParser();ap.add_argument('--root',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);ap.add_argument('--cases',nargs='+',default=['default_start_no_detail','explicit_hold','explicit_start','explicit_linear','explicit_blank_mode_hold','water_series']);a=ap.parse_args();r=check(a.root,a.cases);a.out.write_text(json.dumps(r,ensure_ascii=False,indent=2)+'\n',encoding='utf8');print('PASS',len(r['cases']),'resume cases including immediate load and following WAIT')
