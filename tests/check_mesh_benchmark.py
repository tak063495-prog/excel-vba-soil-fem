"""Independently check terrain mesh equivalence, area/Jacobians and measured timings."""
from pathlib import Path
import argparse,json,statistics
import numpy as np
from check_groundwater import shape

def load(p): return json.loads(p.read_text(encoding='utf-8-sig'))
def validate(record):
    nodes=np.array(record['nodes'],dtype=object).reshape(-1,14)[1:]
    elements=np.array(record['elements'],dtype=object).reshape(-1,10)[1:]
    xy=nodes[:,1:3].astype(float);cn=elements[:,1:9].astype(int)-1
    assert list(nodes[:,0])==list(range(1,len(nodes)+1))
    assert cn.min()>=0 and cn.max()<len(nodes) and len(set(cn.ravel()))==len(nodes)
    area=0.;mindet=float('inf');minangle=180.;maxangle=0.
    derivs=[shape(r,s)[1] for r in (-1,-.5,0,.5,1) for s in (-1,-.5,0,.5,1)]
    for elem in cn:
        coords=xy[elem];corners=coords[:4]
        ar=sum(np.cross(corners[j],corners[(j+1)%4]) for j in range(4))/2
        assert ar>0;area+=ar
        for d in derivs:
            det=float(np.linalg.det(coords.T@d.T));assert det>0;mindet=min(mindet,det)
        for j in range(4):
            u=corners[(j+1)%4]-corners[j];v=corners[(j-1)%4]-corners[j]
            a=float(np.degrees(np.arccos(np.clip(u@v/(np.linalg.norm(u)*np.linalg.norm(v)),-1,1))))
            minangle=min(minangle,a);maxangle=max(maxangle,a)
    expected={'slope2':116.,'slope1':116.,'slope2_quality':116.,'slope2_delaunay':116.,'flat1':160.,'steep_fallback':104.8,'hole_fallback':152.}[record['case']]
    assert abs(area-expected)<1e-7,(record['case'],area,expected)
    return dict(nodes=len(nodes),elements=len(cn),area_m2=area,min_jacobian=mindet,min_angle_deg=minangle,max_angle_deg=maxangle)

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--reference',required=True);p.add_argument('--candidate',required=True);p.add_argument('--report',required=True);p.add_argument('--direct',action='store_true');a=p.parse_args()
    before,after=Path(a.reference),Path(a.candidate);aa=load(before/'summary.json');bb=load(after/'summary.json');report=[]
    assert {(x['case'],x['repeat']) for x in aa}=={(x['case'],x['repeat']) for x in bb}
    for case in dict.fromkeys(x['case'] for x in aa):
        old=[x for x in aa if x['case']==case];new=[x for x in bb if x['case']==case]
        for b in old:
            x=load(before/f"{case}_{b['repeat']}.json");y=load(after/f"{case}_{b['repeat']}.json")
            if not a.direct or case=='slope2_quality':
                assert x['nodes']==y['nodes'] and x['elements']==y['elements'] and x['result']==y['result'],case
        geometry=validate(load(after/f'{case}_0.json'))
        if a.direct and case!='slope2_quality':
            from experiment_mesh_terrain import structured_mesh
            h=1. if case=='slope2' else .5
            nx=round(20/h);ny=round(8/h)
            height=lambda t:8. if case=='flat1' or t*20<=6 else (8.-(t*20-6)*2/3 if t*20<=12 else 4.)
            target_nodes,target_conn=structured_mesh(nx,ny,height)
            target=np.array(target_nodes);target[:,0]*=20
            current=load(after/f'{case}_0.json')
            current_nodes=np.array(current['nodes'],dtype=object).reshape(-1,14)[1:,1:3].astype(float)
            current_conn=np.array(current['elements'],dtype=object).reshape(-1,10)[1:,1:9].astype(int)-1
            assert np.allclose(current_nodes,target,rtol=0,atol=1e-12) and np.array_equal(current_conn,np.array(target_conn))
        t0=statistics.median(x['seconds'] for x in old);t1=statistics.median(x['seconds'] for x in new)
        report.append(dict(case=case,repeats=len(old),reference_median_seconds=t0,candidate_median_seconds=t1,reduction_percent=100*(1-t1/t0),reference_nodes=old[0]['nodes'],reference_elements=old[0]['elements'],all_mesh_cells_exact=not a.direct or case=='slope2_quality',direct_python_mesh_equivalent=a.direct and case!='slope2_quality',geometry=geometry,reference_profile=old[0]['profile'],candidate_profile=new[0]['profile']))
    result=dict(status='PASS',cases=report,note='Wall time includes shape translation, mesh generation, quality and preview; sequential runs, not a speed guarantee. Direct mode uses a newly generated mesh with spacing/2 Q8 edges, so its mesh may differ. AngleFG times are nested within optimizer time, not additive.')
    Path(a.report).write_bytes((json.dumps(result,ensure_ascii=False,indent=2)+'\n').encode());print(json.dumps(result,ensure_ascii=True,indent=2))
