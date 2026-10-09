"""Make standalone plots from our mesh and completed native trial data."""
from pathlib import Path
import argparse, csv, json, math
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, black, white
parser=argparse.ArgumentParser()
parser.add_argument('--output',type=Path,required=True)
parser.add_argument('--fixtures',type=Path,default=Path(__file__).with_name('fixtures.json'))
args=parser.parse_args();dest=args.output/'figures';dest.mkdir(exist_ok=True)
fixtures=json.loads(args.fixtures.read_text(encoding='utf-8-sig'))['fixtures']
c=canvas.Canvas(str(dest/'case_geometry.pdf'),pagesize=(850,420))
for name in ['griffiths_1999_coarse','pruska_h7_phi10']:
 f=next(f for f in fixtures if f['name']==name);nodes={r[0]:r for r in f['nodes']};scale=17.;ox=65;oy=85
 c.setFont('Helvetica-Bold',17);c.drawString(45,385,name+' / Q8 mesh')
 c.setFont('Helvetica',11);c.drawString(45,363,f"Nodes: {len(nodes)}    Elements: {len(f['elements'])}    Area: {f['checks']['area_sum']:.0f} m2    thickness: 1 m")
 c.setStrokeColor(HexColor('#64748b'));c.setLineWidth(.35)
 for element in f['elements']:
  path=c.beginPath()
  for i,n in enumerate(element[1:5]):
   x=ox+scale*nodes[n][1];y=oy+scale*nodes[n][2]
   if i==0:path.moveTo(x,y)
   else:path.lineTo(x,y)
  path.close();c.drawPath(path)
 c.setFillColor(HexColor('#b91c1c'))
 for n,x,y,cx,cy,*rest in f['nodes']:
  if cx or cy:c.circle(ox+scale*x,oy+scale*y,1.25,stroke=0,fill=1)
 c.setFillColor(black);c.setFont('Helvetica',11)
 if name.startswith('griffiths'):
  for x,text in [(0,'0'),(12,'12'),(32,'32 m')]:c.drawCentredString(ox+scale*x,oy-20,text)
  c.drawString(35,oy+scale*10,'10 m')
  c.drawString(45,48,'Left: ux=0. Base: ux=uy=0. Slope and top: free. H=10 m; dry self-weight.')
 else:
  for x,text in [(0,'0'),(15,'15'),(25,'25'),(40,'40 m')]:c.drawCentredString(ox+scale*x,oy-20,text)
  c.drawString(30,oy+scale*8,'8 m');c.drawString(755,oy+scale*15,'15 m')
  c.drawString(45,48,'Sides: ux=0. Base: ux=uy=0. Slope height h=7 m; dry self-weight.')
 c.setFont('Helvetica',9);c.drawString(45,27,'Red dots indicate restrained nodes. Mesh and boundary assumptions are ours; source drawings are not reproduced.')
 c.showPage()
c.save()

if (args.output/'srm_trials.csv').exists():
 trials=list(csv.DictReader((args.output/'srm_trials.csv').open(encoding='utf-8-sig')))
 c=canvas.Canvas(str(dest/'griffiths_displacement_curve.pdf'),pagesize=(850,480))
 c.setFont('Helvetica-Bold',17);c.drawString(50,440,'Griffiths Example 1: displacement before failure')
 c.setFont('Helvetica',9);c.drawString(50,420,'Native: max absolute free-DOF component. Paper: reported maximum nodal displacement; component/norm unspecified.')
 ox,oy,w,h=90,95,580,285
 def X(x):return ox+(x-.8)/.6*w
 def Y(y):return oy+(math.log10(y)+.5)*h
 c.setLineWidth(.5);c.setFont('Helvetica',10)
 for x in [.8,.9,1.,1.1,1.2,1.3,1.4]:
  c.setStrokeColor(HexColor('#d1d5db'));c.line(X(x),oy,X(x),oy+h)
  c.setFillColor(black);c.drawCentredString(X(x),oy-17,f'{x:.1f}')
 for y in [.316,1,3.162]:
  c.setStrokeColor(HexColor('#d1d5db'));c.line(ox,Y(y),ox+w,Y(y))
  c.setFillColor(black);c.drawRightString(ox-10,Y(y)-3,f'{y:g}')
 c.setStrokeColor(black);c.line(ox,oy,ox+w,oy);c.line(ox,oy,ox,oy+h)
 c.drawCentredString(ox+w/2,oy-40,'Yield-strength reduction factor (DAVIS shifted; flow rule still differs)')
 c.saveState();c.translate(28,oy+h/2);c.rotate(90);c.drawCentredString(0,0,'Reported normalized displacement, logarithmic scale');c.restoreState()
 curves=[('Paper: psi=0',[(.8,.379),(1.,.381),(1.2,.422),(1.3,.453),(1.35,.544)],'#111827')]
 for case,label,color in [('griffiths_1999_coarse_davis','DAVIS, 200 elements','#0369a1'),('griffiths_1999_fine_davis','DAVIS, 450 elements','#b45309'),('griffiths_1999_coarse','INCONSISTENT, 200','#a21caf')]:
  points={float(t['equivalent_yield_strength_factor']):float(t['normalized_displacement']) for t in trials if t['case']==case and t['numerical_status']=='CONVERGED' and 'normalized_displacement' in t and t['normalized_displacement']}
  if points:curves.append((label,sorted(points.items()),color))
 for k,(label,points,color) in enumerate(curves):
  c.setStrokeColor(HexColor(color));c.setFillColor(HexColor(color));c.setLineWidth(1.5)
  for a,b in zip(points,points[1:]):c.line(X(a[0]),Y(a[1]),X(b[0]),Y(b[1]))
  for x,y in points:c.circle(X(x),Y(y),2.5,stroke=1,fill=1)
  c.line(685,365-k*27,710,365-k*27);c.setFont('Helvetica',9);c.drawString(715,362-k*27,label)
 c.setFillColor(black);c.setFont('Helvetica',9)
 c.drawString(50,36,'Source: Griffiths & Lane (1999), Table 2. DAVIS: Fy=sqrt(Fs^2+tan(phi0)^2) for psi0=0; Fy is not a corrected FOS.')
 c.drawString(50,20,'Successful full-load trials only. Paper failure at Fs=1.40; nonconverged trial displacements are excluded.')
 c.save()
print(dest)
