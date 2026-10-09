"""Independent geometries, consistent traction loads and exact elastic solutions."""
import json, math
from pathlib import Path
ROOT=Path(__file__).resolve().parent
GAUSS=[(-1/math.sqrt(3),-1/math.sqrt(3)),(1/math.sqrt(3),-1/math.sqrt(3)),(1/math.sqrt(3),1/math.sqrt(3)),(-1/math.sqrt(3),1/math.sqrt(3))]
def q8_mesh(us,vs,mapping,boundary):
 ids={};nodes=[];elements=[];gp=[];edges={};area=0.;min_det=float('inf')
 def node(u,v):
  x,y=mapping(u,v);key=(round(x,11),round(y,11))
  if key not in ids:
   ids[key]=len(nodes)+1;cx,cy=boundary(x,y);nodes.append([ids[key],x,y,cx,cy,0.,0.,0.,0.])
  return ids[key]
 for j in range(len(vs)-1):
  for i in range(len(us)-1):
   u0,u1=us[i:i+2];v0,v1=vs[j:j+2]
   uv=[(u0,v0),(u1,v0),(u1,v1),(u0,v1),((u0+u1)/2,v0),(u1,(v0+v1)/2),((u0+u1)/2,v1),(u0,(v0+v1)/2)]
   n=[node(u,v) for u,v in uv];e=len(elements)+1;elements.append([e,*n,1]);p=[nodes[k-1][1:3] for k in n[:4]]
   a=sum(p[k][0]*p[(k+1)%4][1]-p[(k+1)%4][0]*p[k][1] for k in range(4))/2;assert a>0;area+=a
   for k in range(4):
    key=tuple(sorted((n[k],n[(k+1)%4])))
    if key in edges:assert edges[key][0]==n[4+k]
    edges[key]=(n[4+k],edges.get(key,(0,0))[1]+1)
   for k,(r,s) in enumerate(GAUSS):
    N=[(1-r)*(1-s)/4,(1+r)*(1-s)/4,(1+r)*(1+s)/4,(1-r)*(1+s)/4]
    dr=[-(1-s)/4,(1-s)/4,(1+s)/4,-(1+s)/4];ds=[-(1-r)/4,-(1+r)/4,(1+r)/4,(1-r)/4]
    xr=sum(dr[l]*p[l][0] for l in range(4));xs=sum(ds[l]*p[l][0] for l in range(4));yr=sum(dr[l]*p[l][1] for l in range(4));ys=sum(ds[l]*p[l][1] for l in range(4))
    det=xr*ys-xs*yr;assert det>0;min_det=min(min_det,det)
    gp.append([e,k+1,sum(N[l]*p[l][0] for l in range(4)),sum(N[l]*p[l][1] for l in range(4)),det])
 assert all(c<=2 for _,c in edges.values())
 return nodes,elements,{'area_sum':area,'node_count':len(nodes),'element_count':len(elements),'min_gp_determinant':min_det,'gp_coordinates':gp,'shared_edges':sum(c==2 for _,c in edges.values())}
def material(E,nu,gamma,c,phi,kind='SOIL'):return [1,E,nu,1,gamma,phi,c,0,1 if kind=='SOIL' else 0,'ON',kind,None,None]
def case(name,nodes,elements,checks,mat,stages,loads,reference):
 checks['surface_force_sum']=[sum(r[4] for r in loads if r[2]=='X'),sum(r[4] for r in loads if r[2]=='Y')];checks['body_force_sum_y']=-checks['area_sum']*mat[4]*mat[3]
 return {'name':name,'nodes':nodes,'elements':elements,'checks':checks,'materials':[mat],'stages':stages,'loads':loads,'reference':reference}
def main():
 fixtures=[]
 for nx,ny,label in [(20,10,'coarse'),(30,15,'fine')]:
  n,e,c=q8_mesh([i/nx for i in range(nx+1)],[i/ny for i in range(ny+1)],lambda u,v:(u*(32-20*v),10*v),lambda x,y:(int(abs(x)<1e-9 or abs(y)<1e-9),int(abs(y)<1e-9)));assert abs(c['area_sum']-220)<1e-9
  fixtures.append(case('griffiths_1999_'+label,n,e,c,material(100000,.3,20,10,20),[[1,'GRAVITY',None,None,1,'自重'],[2,'SRM',None,None,1,'強度低減']],[],{'source_id':'griffiths_lane_1999_ex1','fos_reported':1.4,'published_pass_fail':[1.35,1.4],'lem_reference':1.38,'H':10,'normalised_displacement_formula':'E*umax/(gamma*H^2)','explicit_assumptions':['H=10m and gamma=20kN/m3 preserve c/(gamma H)=0.05','E=100000kPa, nu=.3 are nominal values suggested on p390','stress-free gravity turn-on; psi=0 non-associated']}))
 for h,phi,refs in [(7,10,[1.21,1.22,1.31]),(7,20,[1.64,1.65,1.71]),(10.5,10,[.83,.85,.91])]:
  us=[i*15/9 for i in range(10)]+[15+i*10/6 for i in range(1,7)]+[25+i*15/9 for i in range(1,10)]
  def top(x):return 8 if x<=15 else 8+h*(x-15)/10 if x<25 else 8+h
  n,e,c=q8_mesh(us,[i/10 for i in range(11)],lambda u,v:(u,v*top(u)),lambda x,y:(int(abs(x)<1e-9 or abs(x-40)<1e-9 or abs(y)<1e-9),int(abs(y)<1e-9)));assert abs(c['area_sum']-(320+20*h))<1e-8
  fixtures.append(case(f'pruska_h{h:g}_phi{phi}',n,e,c,material(5000,.3,24,20,phi),[[1,'GRAVITY',None,None,1,'自重'],[2,'SRM',None,None,1,'強度低減']],[],{'source_id':'pruska_homogeneous','fos_reference_range':[min(refs),max(refs)],'code_values':dict(zip(['ZSoil','Plaxis','GeoFEM'],refs)),'H':h,'explicit_assumptions':['ux fixed at sides and ux/uy fixed at base','stress-free gravity turn-on instead of imposed K0 stresses','15m toe bench, 10m slope run, 15m crest bench; RS2 problem56 confirms dimensions']}))
 for name,q,gamma,tau,mode in [('confined_pressure',100,0,0,'confined'),('confined_gravity',0,20,0,'confined'),('confined_combined',100,20,0,'confined'),('unconfined_pressure',100,0,0,'unconfined'),('simple_shear',0,0,50,'shear')]:
  def boundary(x,y):
   if mode=='confined':return int(abs(x)<1e-9 or abs(x-4)<1e-9),int(abs(y)<1e-9)
   if mode=='unconfined':return int(abs(x)<1e-9 and abs(y)<1e-9),int(abs(y)<1e-9)
   return int(abs(y)<1e-9),1
  n,e,c=q8_mesh([0,1,2,3,4],[0,1,2],lambda u,v:(u,v),boundary);acc={};loads=[]
  for element in e:
   ids=element[1:9]
   if abs(n[ids[2]-1][2]-2)>1e-9:continue
   edge=[ids[3],ids[6],ids[2]];L=n[edge[2]-1][1]-n[edge[0]-1][1]
   for index,w in zip(edge,[L/6,2*L/3,L/6]):acc[index]=acc.get(index,0)+w
  assert abs(sum(acc.values())-4)<1e-12
  for index,w in sorted(acc.items()):
   x,y=n[index-1][1:3]
   if q:loads.append([1,'BOX','Y','FORCE',-q*w,x,y,x,y,'Q8一致節点力'])
   if tau:loads.append([1,'BOX','X','FORCE',tau*w,x,y,x,y,'Q8一致節点力'])
  E=10000.;nu=.25;H=2.;M=E*(1-nu)/((1+nu)*(1-2*nu));G=E/(2*(1+nu));disp=[];stress=[]
  for row in n:
   x,y=row[1:3]
   if mode=='confined':ux,uy=0.,-(q*y+gamma*(H*y-y*y/2))/M
   elif mode=='unconfined':ux,uy=nu*(1+nu)*q*x/E,-(1-nu*nu)*q*y/E
   else:ux,uy=tau*y/G,0.
   disp.append([row[0],ux,uy])
   if row[3]:assert abs(ux)<1e-12
   if row[4]:assert abs(uy)<1e-12
  for el,g,x,y,det in c['gp_coordinates']:
   sy=q+gamma*(H-y);stress.append([el,g,nu/(1-nu)*sy if mode=='confined' else 0.,sy if mode!='shear' else 0.,tau])
  ref={'source_id':'elastic_validation','q':q,'gamma':gamma,'tau':tau,'E':E,'nu':nu,'H':H,'W':4,'confined_modulus':M,'shear_modulus':G,'stress_convention':'normal compression positive; expected tau follows applied positive traction, negated when comparing native compression-positive tensor','expected_displacements':disp,'expected_stress':stress,'expected_reaction_x':-tau*4,'expected_reaction_y':q*4+gamma*8}
  stages=[[1,'GRAVITY',None,None,1,'自重']]
  if loads:stages.append([2,'LOAD',1,None,1,'一致節点力'])
  f=case(name,n,e,c,material(E,nu,gamma,1e9,0,'STRUCT'),stages,loads,ref)
  assert abs(c['area_sum']-8)<1e-12 and abs(c['surface_force_sum'][0]-tau*4)<1e-10 and abs(c['surface_force_sum'][1]+q*4)<1e-10
  fixtures.append(f)
 p=ROOT/'fixtures.json';p.write_text(json.dumps({'schema':'literature_fem_fixtures_v3','fixtures':fixtures},ensure_ascii=False,indent=2),encoding='utf-8')
 for f in fixtures:print(f['name'],f['checks']['node_count'],f['checks']['element_count'],'area',round(f['checks']['area_sum'],6),'Fx/Fy',f['checks']['surface_force_sum'])
 assert len(json.loads(p.read_text(encoding='utf-8'))['fixtures'])==10
if __name__=='__main__':main()
