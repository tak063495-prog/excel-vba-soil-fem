"""Independent checks for the manual's prescribed polyline-water example."""
import argparse,hashlib,json,math,gzip
from pathlib import Path

def read(p):
 if p.exists():return json.loads(p.read_text(encoding='utf-8-sig'))
 with gzip.open(str(p)+'.gz','rt',encoding='utf-8-sig') as stream:return json.load(stream)
def main():
 ap=argparse.ArgumentParser();ap.add_argument('run',type=Path);ap.add_argument('--report',type=Path,required=True);a=ap.parse_args()
 f=Path('tests/fixtures/water_slope_polyline.json');fix=read(f);r=read(a.run/'result.json');rr=read(a.run/'resume.json')
 assert hashlib.sha256(f.read_bytes()).hexdigest().upper()==r['fixture_sha256']
 assert r['status'].startswith('PASS|') and rr['status'].startswith('PASS|')
 pts=[(row[1],row[2]) for row in fix['water']['profiles']]
 def level(x):
  for (x0,y0),(x1,y1) in zip(pts,pts[1:]):
   if x0-1e-8<=x<=x1+1e-8:return y0+(y1-y0)*(x-x0)/(x1-x0)
  raise AssertionError(x)
 # Values below are calculated independently at the exported GP coordinates.
 rows=[row for row in read(a.run/'水圧結果.json')[35:] if isinstance(row[0],(int,float)) and isinstance(row[1],(int,float)) and isinstance(row[5],(int,float))]
 assert len(rows)==160
 ep=max(abs(row[5]-fix['water']['gamma']*max(level(row[3])-row[4],0)) for row in rows)
 es=max(abs(row[c]-row[c+1]-row[5]) for row in rows for c in (6,8,10))
 assert ep<1e-7 and es<1e-10
 ns=[row[:10] for row in read(a.run/'結果節点.json') if isinstance(row[0],(int,float))];last=max(row[0] for row in ns);final=[row for row in ns if row[0]==last];assert len(final)==len(fix['nodes'])
 byid={int(row[1]):row for row in final}
 for n in fix['nodes']:
  z=byid[n[0]];assert abs(z[2]-n[1])<1e-10 and abs(z[3]-n[2])<1e-10
  if n[3]:assert abs(z[4])<1e-14
  if n[4]:assert abs(z[5])<1e-14
 du=max(abs(x-y) for x,y in zip(r['disp'],rr['disp']));assert du<1e-9
 maxu=max(math.hypot(z[4],z[5]) for z in final)
 weight=sum(abs(sum((fix['nodes'][el[i+1]-1][1]*fix['nodes'][el[(i+1)%4+1]-1][2]-fix['nodes'][el[(i+1)%4+1]-1][1]*fix['nodes'][el[i+1]-1][2]) for i in range(4)))/2 for el in fix['elements'])*20
 ry=sum(z[9] for z in final);assert abs(ry-weight)<1e-6 and abs(r['metrics'][2]-ry)<1e-6
 assert r['metrics'][3]<1e-7 and r['metrics'][4:]==[1,1,0,0]
 report=dict(status='PASS',fixture_sha256=r['fixture_sha256'],source_book_sha256=r['book_sha256'],nodes=len(final),elements=len(fix['elements']),GP_rows=len(rows),pressure_max_abs_error_kPa=ep,total_effective_pressure_max_abs_error_kPa=es,weight_kN=weight,reactionY_kN=ry,max_displacement_norm_m=maxu,relative_residual=r['metrics'][3],fixed_Fs=1.0,seconds=r['seconds'],resume_max_du_m=du,resume_relative_residual=rr['metrics'][3],profile=pts,scope='Prescribed polyline pore pressure; no flow/free-surface solve or Fs-bracket search.')
 a.report.parent.mkdir(parents=True,exist_ok=True);a.report.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');print(json.dumps(report))
if __name__=='__main__':main()
