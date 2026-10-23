"""Additional cold-process, direct-view and observer-control checks for #50."""
from pathlib import Path
import argparse,json
from check_issue50_resume import read,close,state

def decode(v):
 assert len(v)==16
 return dict(zip(['stage','elapsed','duration','time','accepted','profile','first','last','fraction','external_load','boundary_locked','boundary_pressure','heads','flow','flow_reaction','gp_pressure'],v))
def check(root):
 out={'status':'PASS','cold':[],'view':[]}
 for label in ['mid','completed','old07']:
  p=root/('cold_'+label);a=read(root/'matrix_auto/issue50_default_start_no_detail_completed/result.json');r=read(p/'continued.json')
  assert r['status'].startswith('PASS|') and r['fresh_excel_process'] and not r['preloaded_frame']
  final=read(root/'matrix_auto/issue50_default_start_no_detail_completed/water_observables.json')
  loaded=decode(read(p/'water_loaded.json'));current=decode(read(p/'water_state.json'))
  matches=[v for v in final['frames'] if v['stage']==3 and v['elapsed']==loaded['elapsed']];assert len(matches)==1
  delta=state(matches[0],loaded);end=state(final,current)
  rows=read(root/'matrix_auto/issue50_default_start_no_detail_completed/結果節点.json')
  nodes={row[1]:row[4:6] for row in rows[1:] if row[0]==4}
  du=max(close(nodes[row[0]],row[1:],1e-9) for row in r['nodes'])
  metrics=close(a['metrics'],r['metrics'],1e-6)
  out['cold'].append({'case':label,'status':'PASS','loaded_elapsed_s':loaded['elapsed'],'loaded_delta':delta,'final_delta':end,'u_max_m':du,'metrics_max_delta':metrics,'seconds':r['seconds']})
 for label in ['mid','completed']:
  p=root/('view_'+label);current=decode(read(p/'water_state.json'));loaded=decode(read(p/'water_loaded.json'))
  a=read(root/'matrix_auto/issue50_default_start_no_detail_completed/water_observables.json')
  matches=[v for v in a['frames'] if v['stage']==3 and v['elapsed']==current['elapsed']];assert len(matches)==1
  out['view'].append({'case':label,'status':'PASS','view_delta':state(matches[0],current),'loaded_delta':state(matches[0],loaded)})
 for phase in ['result','resume']:
  a=read(root/'matrix_auto/issue50_default_start_no_detail_completed'/(phase+'.json'));b=read(root/'no_observer'/(phase+'.json'))
  out['observer_'+phase]={'u_max_m':close(a['disp'],b['disp'],0),'stress_max_kPa':close(a['stress'],b['stress'],0),'metrics_max_delta':close(a['metrics'],b['metrics'],0)}
 p=root/'series_fractional';a=read(p/'water_observables.json');b=read(p/'resume_water_observables.json');v=b['loaded']
 assert v['stage']==3 and v['elapsed']==10 and close(v['selectors'][2],.5,1e-15)==0
 expected=next(vv for vv in a['frames'] if vv['stage']==3 and vv['elapsed']==10)
 out['fractional_series']={'status':'PASS','loaded_delta':state(expected,v),'final_delta':state(a,b),'fraction':v['selectors'][2]}
 return out
if __name__=='__main__':
 ap=argparse.ArgumentParser();ap.add_argument('--root',type=Path,required=True);ap.add_argument('--out',type=Path,required=True);a=ap.parse_args();r=check(a.root);a.out.write_text(json.dumps(r,ensure_ascii=False,indent=2)+'\n',encoding='utf8');print('PASS cold3/view2/fractionalseries/observer-free control')
