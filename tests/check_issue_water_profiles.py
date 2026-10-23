"""Independent analytic checks for the XLSM water-boundary issue fixtures."""
import argparse,json,pathlib,math

GOOD=('issue44_external_linear','issue44_external_layers','issue45_equivalent')
REJECT=('issue44_internal_head','issue45_start_only','issue45_hidden_knot','issue45_numeric_head')
def main():
 p=argparse.ArgumentParser();p.add_argument('--root',required=True);p.add_argument('--out',required=True);a=p.parse_args();root=pathlib.Path(a.root); report={'cases':{},'limitations':[]}
 for name in GOOD:
  d=root/name; r=d/'result.json'
  if not r.exists(): raise SystemExit('missing result: '+name)
  j=json.loads(r.read_text(encoding='utf-8'))
  if not str(j.get('status','')).startswith('PASS|'): raise SystemExit('nonpass: '+name)
  m=j.get('metrics',[]); gp=0; maxabs=0.0
  wr=d/'水圧結果.json'
  if not wr.exists(): raise SystemExit('missing water export: '+name)
  if wr.exists():
   errp=errq=errh=errr=0.0; gp_rows=[]; node_rows=[]; face_rows=[]
   for row in json.loads(wr.read_text(encoding='utf-8'))[35:]:
    if len(row)>13 and isinstance(row[0],(int,float)):
     vals=[row[i] for i in (3,4,5,12,13) if row[i] is not None]
     if len(vals)<3 or not all(isinstance(v,(int,float)) and math.isfinite(v) for v in vals): raise SystemExit('nonfinite GP water result: '+name)
     gp+=1; gp_rows.append(row); maxabs=max(maxabs,max(abs(float(v)) for v in vals))
     x,y,pv=float(row[3]),float(row[4]),float(row[5]); qx=float(row[12] or 0); qy=float(row[13] or 0)
     if name=='issue44_external_linear': h=3-.5*x; q=5e-6
     elif name=='issue44_external_layers': h=3-(2/3)*x if x<=1 else 7/3-(1/3)*(x-1); q=6.666666666666667e-6
     else: continue
     errp=max(errp,abs(pv-10*(h-y)));errq=max(errq,abs(qx-q),abs(qy))
    if len(row)>38 and isinstance(row[31],(int,float)): node_rows.append(row); errh=max(errh,abs(float(row[33])-(3-.5*float(row[37])) if name=='issue44_external_linear' else abs(float(row[33])-(3-(2/3)*float(row[37]) if float(row[37])<=1 else 7/3-(1/3)*(float(row[37])-1)))))
    if len(row)>44 and isinstance(row[40],(int,float)): face_rows.append(row)
   if name.startswith('issue44') and (len(gp_rows)!=8 or len(node_rows)!=13): raise SystemExit(f'export count mismatch {name}: gp={len(gp_rows)} nodes={len(node_rows)}')
   if errh>1e-9: raise SystemExit(f'analytic head mismatch {name}: {errh}')
   if name.startswith('issue44'):
    if len(face_rows)!=8: raise SystemExit('missing element faces: '+name)
    left=sum(float(r[35]) for r in node_rows if abs(float(r[37]))<1e-9)
    right=sum(float(r[35]) for r in node_rows if abs(float(r[37])-2)<1e-9)
    total=sum(float(r[35]) for r in node_rows)
    errr=max(abs(left-q),abs(right+q),abs(total))
    errface=max(abs(float(r[43])-(q if int(r[42])==2 else -q if int(r[42])==4 else 0)) for r in face_rows)
    if max(errr,errface)>1e-12: raise SystemExit(f'boundary/face mismatch {name}: reaction={errr},face={errface}')
   else:
    if gp!=4: raise SystemExit('missing hydro GP: '+name)
    for r in gp_rows:
     if abs(float(r[5])-10*(2-float(r[4])))>1e-8: raise SystemExit('hydro pressure mismatch')
    surface=[r for r in json.loads(wr.read_text(encoding='utf-8'))[35:] if r[22] is not None]
    if len(surface)!=1 or max(abs(float(surface[0][27])),abs(float(surface[0][28])+10),abs(float(surface[0][29])+5))>1e-8: raise SystemExit('hydro boundary resultant mismatch')
   if errp>1e-8 or errq>1e-12: raise SystemExit(f'analytic water mismatch {name}: p={errp} q={errq}')
  report['cases'][name]={'status':'PASS','metrics':m,'gp_rows_checked':gp,'node_rows_checked':len(node_rows),'face_rows_checked':len(face_rows),'max_abs_water_value':maxabs,'analytic_max_pressure_error':errp,'analytic_max_flux_error':errq,'analytic_max_head_error':errh,'boundary_reaction_error':errr}
  if name=='issue45_equivalent' and (len(m)<3 or abs(float(m[2])-30)>1e-8): raise SystemExit('issue45 reaction is not 30 kN')
 for name in REJECT:
  d=root/name; found=[]
  for n in ('failure.json','expected_failure.json','input_check.json'):
   if (d/n).exists(): found.append(n)
  if not found: raise SystemExit('missing rejection evidence: '+name)
  diagnostics=[]
  for f in found:
   x=json.loads((d/f).read_text(encoding='utf-8')); diagnostics.append(x.get('actual_error',x.get('expected_error',x.get('status',''))))
  if not any(str(x) for x in diagnostics): raise SystemExit('rejection evidence has no diagnostic: '+name)
  report['cases'][name]={'status':'REJECTED','evidence':found,'diagnostics':diagnostics}
 report['limitations']=[]
 pathlib.Path(a.out).write_text(json.dumps(report,indent=2),encoding='utf-8');print('PASS')
if __name__=='__main__': main()
