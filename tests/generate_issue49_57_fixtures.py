import copy,json
from pathlib import Path
D=Path('tests/fixtures/issue49_57');D.mkdir(exist_ok=True)
def base(n):return json.loads(Path(f'tests/fixtures/fixed_stress/{n}.json').read_text(encoding='utf-8-sig'))
def save(n,m):
 m['name']=n;(D/(n+'.json')).write_text(json.dumps(m,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
for mode in ('START','LINEAR','SERIES','DOWN'):
 m=base('joint_transient');m['jointPairs'][0][5:]=['PRESCRIBED','*'];m['water']['timeOptions'][0][2]='START' if mode=='START' else 'LINEAR'
 if mode=='DOWN':
  for p in m['water']['profiles']:
   if p[0]=='W':p[2]=2.2
   else:p[2]=2.
 if mode=='SERIES':m['water']['timeOptions'][0][7]='JCURVE';m['water']['timeCurves']=[['JCURVE',0,'W'],['JCURVE',43200,'W2'],['JCURVE',86400,'W2']]
 save('joint_'+mode.lower(),m)
for mode in ('HOLD','START','LINEAR','SERIES'):
 m=base('water_external_linear');m['water']['timeOptions'][0][2]=mode if mode!='SERIES' else 'LINEAR';m['stages'].append([4,'WAIT','',1,1,'after water']);m['water']['stagePlans'].append([4,'INHERIT','',20,10,'',200,1e-8,''])
 if mode=='SERIES':m['water']['timeOptions'][0][7]='CURVE';m['water']['timeCurves']=[['CURVE',0,'W'],['CURVE',20,'W2'],['CURVE',40,'W2']]
 save('water_'+mode.lower(),m)
for mode in ('INSTANT','LINEAR','DISP'):
 m=base('named_unload');m['stages']=[[1,'GRAVITY','',1,1,'initial'],[2,'LOAD',1,1,1,'load A'],[3,'LOAD',1,1,1,'load B, inherit A'],[4,'WAIT','',1,1,'hold both'],[5,'UNLOAD',1,1,1,'explicit A zero'],[6,'WAIT','',1,1,'hold B']]
 m['loads']=[[1,'40','Y','FORCE',-2,None,None,None,None,'','A'],[1,'39','Y','FORCE',-4,None,None,None,None,'','B'],[1,'40','Y','FORCE',0,None,None,None,None,'','U']]
 if mode=='DISP':m['loads'][0][3:5]=['DISP',-0.0013]
 m['water']['stagePlans']=[[2,'FIXED_STRESS','',0,0,'',200,1e-8,''],[3,'INHERIT','',43200 if mode=='LINEAR' else 0,10800 if mode=='LINEAR' else 0,'',200,1e-8,''],[4,'INHERIT','',43200,10800,'',200,1e-8,''],[5,'INHERIT','',0,0,'',200,1e-8,''],[6,'INHERIT','',43200,10800,'',200,1e-8,'']]
 m['water']['timeOptions']=[[2,'S','HOLD','INSTANT',0,0,0,'','A',''],[3,'S','HOLD','LINEAR' if mode=='LINEAR' else 'INSTANT',0,0,0,'','B',''],[5,'S','HOLD','INSTANT',0,0,0,'','U','']]
 if mode=='DISP':
  # Drain during the imposed-displacement ramp; an instantaneous, localized
  # constraint produces a negative-pressure candidate outside this model.
  m['water']['stagePlans'][0][3:5]=[86400,21600]
  m['water']['timeOptions'][0][3]='LINEAR'
  m['loads'][1][4]=-1.
  m['water']['stagePlans'][1][3:5]=[86400,21600]
  m['water']['timeOptions'][1][3]='LINEAR'
 save('named_'+mode.lower(),m)
# Explicit shared-face NOFLOW, plus exterior and separated/continuous JOINT.
m=base('column_smoke');m['water']['drainSets']=[['D',1,3,'NOFLOW','',None,'internal']];m['water']['stagePlans'][0][5]='D';m['expectedAnalysisError']='NOFLOW';save('noflow_internal',m)
m=base('column_smoke');m['water']['drainSets']=[['D',8,3,'HEAD','W',None,'drained top'],['D',1,1,'NOFLOW','',None,'base']];m['water']['stagePlans'][0][5]='D';save('noflow_exterior',m)
for mode in ('PRESCRIBED','CONTINUOUS'):
 m=base('joint_transient');m['materials'][1]=copy.deepcopy(m['materials'][0]);m['materials'][1][0]=2;m['materials'][1][8]=0
 m['water']['materials'][1]=[2,'EFFECTIVE',1e-5,'*','soil'];m['water']['coupledMaterials'].append([2,1,1e-6,1e-5,1e-5,'second soil',.35]);m['jointPairs'][0][5:]=[mode,'*' if mode=='PRESCRIBED' else '']
 m['stages']=[[1,'GRAVITY','',1,1,'gravity'],[2,'WAIT','',1,1,'coupled hold']];m['water']['stagePlans']=[[2,'FIXED_STRESS','',86400,21600,'D',200,1e-8,'']];m['water']['timeOptions']=[];m['water']['drainSets']=[['D',1,1,'HEAD','W',None,'base'],['D',1,3,'NOFLOW','',None,'joint face']]
 if mode=='CONTINUOUS':m['expectedAnalysisError']='NOFLOW'
 save('noflow_joint_'+mode.lower(),m)
# Dry manual-input topology guards and a conforming patch control.
material=[[1,1000,0,1,0,0,1e6,0,0,1,'SOIL',0,0]]
xy=[(0,0),(1,0),(1,2),(0,2),(.5,0),(1,1),(.5,2),(0,1),(2,0),(2,1),(1.5,0),(2,.5),(1.5,1),(1,.5),(2,2),(2,1.5),(1.5,2),(1,1.5)]
conn=[[1,1,2,3,4,5,6,7,8,1],[2,2,9,10,6,11,12,13,14,1],[3,6,10,15,3,13,16,17,18,1]]
def dry(xy,conn):return {'nodes':[[i+1,x,y,1,1,.001*x,0,0,0] for i,(x,y) in enumerate(xy)],'elements':conn,'materials':material,'stages':[[1,'GRAVITY','',1,1,'patch']],'loads':[],'water':{'mode':'NONE','gamma':9.81,'defaultProfile':'','auto':'OFF','profiles':[],'materials':[],'boundaries':[]}}
m=dry(xy,conn);m['expectedAnalysisError']='hanging node';save('hanging_q8',m)
# Two valid straight T6s; gravity makes nonaffine displacement and must reject RESET before mutation.
xy=[(0,0),(1,0),(1,1),(.5,0),(1,.5),(.5,.5),(0,1),(.5,1),(0,.5)];conn=[[1,1,2,3,3,4,5,3,6,1,'T6'],[2,1,3,7,7,6,8,7,9,1,'T6']]
m=dry(xy,conn);m['materials'][0][1:5]=[10000,.3,1,18];m['nodes']=[[i+1,x,y,1,1 if y==0 else 0,0,0,0,0] for i,(x,y) in enumerate(xy)];m['stages'].append([2,'RESET_STRESS','',1,1,'reset']);m['expectedAnalysisError']='曲線T6';save('reset_t6_nonaffine',m)
m=copy.deepcopy(m);m.pop('expectedAnalysisError');m['materials'][0][4]=0;save('reset_t6_zero',m)
# Affine prescribed displacement is compatible and retains midpoint locations.
m=copy.deepcopy(m);m['nodes']=[[i+1,x,y,1,1,0,0,0,0] for i,(x,y) in enumerate(xy)]
m['stages'][0]=[1,'LOAD',1,1,1,'affine prescribed displacement']
m['loads']=[[1,str(i+1),direction,'DISP',value,None,None,None,None,''] for i,(x,y) in enumerate(xy) for direction,value in [('X',-.0001*x),('Y',-.0002*y)]]
save('reset_t6_affine',m)
# Four conforming Q8 cells, with a genuine free interior patch test.
points={};cells=[]
def node(x,y):
 if (x,y) not in points:points[x,y]=len(points)+1
 return points[x,y]
for iy in range(2):
 for ix in range(2):
  coords=[(ix,iy),(ix+1,iy),(ix+1,iy+1),(ix,iy+1),(ix+.5,iy),(ix+1,iy+.5),(ix+.5,iy+1),(ix,iy+.5)]
  cells.append([len(cells)+1]+[node(*p) for p in coords]+[1])
xy=list(points);m=dry(xy,cells);m['nodes']=[[i+1,x,y,int(x in (0,2) or y in (0,2)),int(x in (0,2) or y in (0,2)),0,0,0,0] for i,(x,y) in enumerate(xy)]
m['stages']=[[1,'LOAD',1,1,1,'constant strain patch']];m['loads']=[[1,str(i+1),direction,'DISP',value,None,None,None,None,''] for i,(x,y) in enumerate(xy) if x in (0,2) or y in (0,2) for direction,value in [('X',.001*x),('Y',0)]]
save('conforming_q8',m)
# Native vertex states must deliberately fall back to the same VBA return.
xy=[(0,0),(1,0),(1,1),(0,1),(.5,0),(1,.5),(.5,1),(0,.5)]
m=dry(xy,[[1,1,2,3,4,5,6,7,8,1]]);m['materials'][0][1:8]=[1000,.3,1,0,30,0,30]
m['nodes']=[[i+1,x,y,1,1,0,0,0,0] for i,(x,y) in enumerate(xy)];m['stages']=[[1,'LOAD',1,1,1,'tensile apex']]
m['loads']=[[1,str(i+1),direction,'DISP',value,None,None,None,None,''] for i,(x,y) in enumerate(xy) for direction,value in [('X',.01*x),('Y',.01*y)]]
save('apex_native',m)
print('generated',len(list(D.glob('*.json'))),'issue cases')
