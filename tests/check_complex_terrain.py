"""Independent geometry, Q8, material and conformity checks of real Excel output."""
from pathlib import Path
from collections import defaultdict
import argparse,json,math
import numpy as np
from check_groundwater import shape

ROOT=Path(__file__).resolve().parents[1]
def load(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def inside(x,y,poly):
    hit=False
    for a,b in zip(poly,poly[1:]+poly[:1]):
        if (a[1]>y)!=(b[1]>y) and x<a[0]+(b[0]-a[0])*(y-a[1])/(b[1]-a[1]):hit=not hit
    return hit
def boundaries(case):
    points={int(p[0]):p[1:] for p in case['points']}
    return [(x['kind'],[points[int(i)] for i in x['ids'].split(',')]) for x in case['loops']]
def material_at(x,y,loops):
    if any(inside(x,y,p) for kind,p in loops if kind=='INNER'):return 0
    for i,(_,p) in enumerate((q for q in loops if q[0]=='OUTER'),1):
        if inside(x,y,p):return i
    return 0
def validate(case,data,plot=None):
    nodes=np.array(data['nodes'],dtype=object).reshape(-1,14)[1:]
    elements=np.array(data['elements'],dtype=object).reshape(-1,10)[1:]
    xy=nodes[:,1:3].astype(float);cn=elements[:,1:9].astype(int)-1;mats=elements[:,9].astype(int)
    assert set(cn.ravel())==set(range(len(xy)))
    assert len(set(map(tuple,np.round(xy,10))))==len(xy),'coincident duplicate nodes'
    loops=boundaries(case);areas=defaultdict(float);edges=defaultdict(list);parent=list(range(len(cn)))
    def root(i):
        while parent[i]!=i:parent[i]=parent[parent[i]];i=parent[i]
        return i
    mindet=math.inf;minangle=180.;maxangle=0.;aspect=0.
    deriv=[shape(r,s)[1] for r in (-1,-.5,0,.5,1) for s in (-1,-.5,0,.5,1)]
    for e,(c,mat) in enumerate(zip(cn,mats)):
        q=xy[c];corners=q[:4]
        area=sum(corners[j,0]*corners[(j+1)%4,1]-corners[(j+1)%4,0]*corners[j,1] for j in range(4))/2
        assert area>0;areas[str(mat)]+=area
        assert material_at(*corners.mean(axis=0),loops)==mat,'wrong material or filled void'
        for d in deriv:
            det=float(np.linalg.det(q.T@d.T));assert det>0;mindet=min(mindet,det)
        lengths=[]
        for j in range(4):
            a=corners[(j+1)%4]-corners[j];b=corners[(j-1)%4]-corners[j];lengths.append(np.linalg.norm(a))
            angle=math.degrees(math.acos(np.clip(a@b/(np.linalg.norm(a)*np.linalg.norm(b)),-1,1)))
            minangle=min(minangle,angle);maxangle=max(maxangle,angle)
            key=tuple(sorted((c[j],c[(j+1)%4])));edges[key].append((e,c[4+j]))
            assert np.allclose(q[4+j],(corners[j]+corners[(j+1)%4])/2,rtol=0,atol=1e-12)
        aspect=max(aspect,max(lengths)/min(lengths))
    assert minangle>=30-1e-7 and maxangle<=150+1e-7 and aspect<=5+1e-9
    for key,uses in edges.items():
        assert len(uses)<=2,'non-manifold edge'
        if len(uses)==2:
            assert uses[0][1]==uses[1][1],'non-shared Q8 midpoint'
            parent[root(uses[0][0])]=root(uses[1][0])
        else:
            # A crack/hanging edge inside the occupied union must not appear as an exterior edge.
            a,b=xy[list(key)];mid=(a+b)/2;normal=np.array([b[1]-a[1],a[0]-b[0]])
            normal=normal/np.linalg.norm(normal)*1e-7
            assert (material_at(*(mid+normal),loops)==0)!=(material_at(*(mid-normal),loops)==0),'artificial internal boundary'
    assert len({root(e) for e in range(len(cn))})==1,'disconnected model'
    # Every visible part of every input segment must be represented by an edge chain.
    # Parts hidden by an earlier material mask have the same material on both sides.
    ep=np.array([xy[list(key)] for key in edges]);edge_a=ep[:,0];edge_b=ep[:,1]
    for _,poly in loops:
        for aa,bb in zip(poly,poly[1:]+poly[:1]):
            aa=np.array(aa);bb=np.array(bb);v=bb-aa;length=float(np.linalg.norm(v));normal=np.array([-v[1],v[0]])/length*1e-7
            cross_a=np.abs((edge_a[:,0]-aa[0])*v[1]-(edge_a[:,1]-aa[1])*v[0])/length
            cross_b=np.abs((edge_b[:,0]-aa[0])*v[1]-(edge_b[:,1]-aa[1])*v[0])/length
            hits=(cross_a<1e-9)&(cross_b<1e-9)
            intervals=[]
            for start,end in zip((edge_a[hits]-aa)@v/(length*length),(edge_b[hits]-aa)@v/(length*length)):
                lo,hi=sorted((float(start),float(end)))
                if hi>0 and lo<1:intervals.append((max(0,lo),min(1,hi)))
            cuts=sorted(set([0.,1.]+[x for interval in intervals for x in interval]))
            for lo,hi in zip(cuts,cuts[1:]):
                if hi-lo<1e-10:continue
                at=(lo+hi)/2;mid=aa+v*at
                if material_at(*(mid+normal),loops)!=material_at(*(mid-normal),loops):
                    assert any(a-1e-10<=at<=b+1e-10 for a,b in intervals),'lost input boundary segment'
    assert set(areas)==set(case['expected_areas'])
    for mat,value in case['expected_areas'].items():assert math.isclose(areas[mat],value,rel_tol=1e-11,abs_tol=1e-9),(mat,areas,value)
    if plot:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        from matplotlib.collections import PolyCollection
        fig,ax=plt.subplots(figsize=(9,6))
        if case.get('roles'):
            from matplotlib.patches import Patch
            pc=PolyCollection(xy[cn[:,:4]],facecolors=['#e7c18a' if case['roles'].get(str(m))=='STRUCT' else '#bcc6cb' for m in mats],edgecolors='#666666',linewidths=.3)
            ax.legend(handles=[Patch(facecolor='#e7c18a',label='Wall / lining'),Patch(facecolor='#bcc6cb',label='Soil')])
        else:pc=PolyCollection(xy[cn[:,:4]],array=mats.astype(float),cmap='Set2',edgecolors='#666666',linewidths=.3)
        ax.add_collection(pc);ax.autoscale();ax.set_aspect('equal');ax.set_xlabel('x (m)');ax.set_ylabel('y (m)');ax.set_title(case['name']+' - actual Excel Q8 mesh');fig.tight_layout();fig.savefig(plot,dpi=180);plt.close(fig)
    return dict(case=case['name'],nodes=len(xy),elements=len(cn),areas=dict(areas),min_jacobian_25=mindet,min_angle=minangle,max_angle=maxangle,max_aspect=aspect,connected=True,shared_midpoints=True,no_artificial_internal_boundaries=True,input_boundary_chains=True,seconds=data['seconds'],true_circle_area_deficit_percent=case.get('true_circle_area_deficit_percent'))
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('folder');a=p.parse_args();folder=Path(a.folder);results=[]
    for case in load(ROOT/'tests/fixtures/complex_terrain.json'):
        data=load(folder/(case['name']+'.json'))
        if case['reject']:
            assert data['status']=='EXPECTED_REJECT' and data['mesh_unchanged'];results.append(dict(case=case['name'],status='EXPECTED_REJECT',mesh_unchanged=True))
        else:results.append(validate(case,data,folder/(case['name']+'.png')))
    report=dict(status='PASS',cases=results,method='Independent polygon membership, analytic material areas, full Q8 25-point derivatives, edge/midpoint incidence, connectedness and boundary classification.')
    (folder/'independent.json').write_bytes((json.dumps(report,indent=2)+'\n').encode());print(json.dumps(report,indent=2))
