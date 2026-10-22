"""Compare complete groundwater regression vectors, intervals and force balances."""
import argparse,json
from pathlib import Path
import numpy as np

def read(path):return json.loads(Path(path).read_text(encoding='utf-8-sig'))
def records(value):return value if isinstance(value,list) else [value]
def main():
 ap=argparse.ArgumentParser();ap.add_argument('root',type=Path);a=ap.parse_args();out=[]
 for flow in ('INCONSISTENT','DAVIS'):
  for suite in ('smoke','smokesrm'):
   ref=records(read(a.root/f'dry_reference_{flow}_{suite}.json'));new=records(read(a.root/f'dry_candidate_{flow}_{suite}.json'));assert len(ref)==len(new)
   for x,y in zip(ref,new):
    assert x['Case']==y['Case'] and '|PASS|' in x['Run'] and '|PASS|' in y['Run']
    dx=np.array(x['Disp']);dy=np.array(y['Disp']);assert np.array_equal(dx,dy),(flow,suite,'displacement')
    assert x['Bracket']==y['Bracket'],(flow,suite,'Fs interval')
    rx=float(x['Run'].split('|')[3]);ry=float(y['Run'].split('|')[3]);assert rx==ry,(flow,suite,'residual')
    out.append(dict(flow=flow,suite=suite,case=x['Case'],all_dofs=len(dx),max_du=0,bracket=y['Bracket'],residual=ry,reference_seconds=x['WallSeconds'],candidate_seconds=y['WallSeconds']))
 x=read(a.root/'wet_fixed_OFF/result.json');y=read(a.root/'wet_fixed_AUTO/result.json');delta=float(np.max(np.abs(np.array(x['disp'])-y['disp'])))
 assert x['status'].startswith('PASS|') and y['status'].startswith('PASS|') and delta<1e-10
 assert x['metrics'][4:8]==y['metrics'][4:8] and max(x['metrics'][3],y['metrics'][3])<1e-6
 out.append(dict(case='wet_OFF_AUTO',all_dofs=len(x['disp']),max_du=delta,OFF_seconds=x['seconds'],AUTO_seconds=y['seconds'],OFF_residual=x['metrics'][3],AUTO_residual=y['metrics'][3]))
 # Partly filled cavity: 90 kN exterior buoyancy minus 5 kN interior water.
 r=read(a.root/'box_inside/result.json');tab=read(a.root/'box_inside/水圧結果.json');fx=fy=moment=0.
 for row in tab[35:]:
  if len(row)>29 and isinstance(row[23],(int,float)):fx+=row[27];fy+=row[28];moment+=row[29]
 assert abs(fx)<1e-8 and abs(fy-85)<1e-8 and abs(moment-127.5)<1e-8 and abs(r['metrics'][2]-107)<1e-7
 out.append(dict(case='inside_water',Fy=fy,M=moment,Ry=r['metrics'][2]))
 for name in ('stages','initial_birth','all_birth','wet_srm'):
  x=read(a.root/name/'result.json');y=read(a.root/name/'resume.json')
  assert x['status'].startswith('PASS|') and y['status'].startswith('PASS|') and np.max(np.abs(np.array(x['disp'])-y['disp']))<1e-9
  assert np.max(np.abs(np.array(x['metrics'][4:8])-y['metrics'][4:8]))<1e-10
  assert y['metrics'][3]<1e-7
  out.append(dict(case=name,seconds=x['seconds'],resume_seconds=y['seconds'],metrics=x['metrics'],resume_metrics=y['metrics']))
 for name,count in (('initial_birth',8),('all_birth',0)):
  tab=read(a.root/name/'水圧確認.json');actual=sum(isinstance(r[0],(int,float)) for r in tab[35:]);assert actual==count,(name,actual,count)
 text=json.dumps(out,ensure_ascii=False,indent=2);(a.root/'comparison.json').write_text(text,encoding='utf-8');print(text)
if __name__=='__main__':main()
