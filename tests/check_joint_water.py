"""Independent force/head checks on actual Excel JOINT-water output tables."""
import argparse,json,math
from pathlib import Path

def load(path):return json.loads(Path(path).read_text(encoding='utf-8-sig'))
def close(actual,expected,tolerance=1e-6):
    assert math.isfinite(float(actual)),actual
    assert abs(actual-expected)<=tolerance,(actual,expected,tolerance)

def run_check(fixture,folder):
    f=load(fixture);root=Path(folder);result=load(root/'result.json')
    assert result['status'].startswith('PASS|'),result['status']
    assert result['metrics'][3]<=1e-7,result['metrics']
    rows=[r for r in load(root/'接合結果.json')[1:] if isinstance(r[0],(float,int))]
    assert rows,'No JOINT rows exported'
    stage=max(r[0] for r in rows);final=sorted([r for r in rows if r[0]==stage],key=lambda r:r[3])
    assert len(final)==3,final
    exp=f['expected'];gamma=f['water']['gamma']
    for r in final:
        assert r[12]==1,'Final joint inactive'
        close(r[16],r[17]+r[15],1e-10);close(r[6],r[17],1e-10)
        if 'p_kPa' in exp:close(r[15],exp['p_kPa'],1e-10)
        if 'waterline_intercept' in exp:close(r[15],gamma*max(0.,exp['waterline_intercept']+exp['waterline_gradient']*r[4]-r[5]),1e-10)
        if 'contact_state' in exp:assert r[10]==exp['contact_state'],r
        if r[10]==1:close(r[17],0.,1e-10);close(r[18],0.,1e-10)
        else:
            fs=exp.get('fixed_fs',1.);material=f['materials'][2]
            cap=material[6]/fs+max(r[17],0.)*math.tan(math.radians(material[5]))/fs
            close(r[18],cap,1e-7);assert abs(r[7])<=cap+1e-7
            if r[10]==2:close(abs(r[7]),cap,1e-7)
    # Straight 2m interface, thickness=1. Three-point Gauss traction integration.
    weights=[5/9,8/9,5/9]
    effective=sum(w*r[17] for w,r in zip(weights,final));pressure=sum(w*r[15] for w,r in zip(weights,final));total=sum(w*r[16] for w,r in zip(weights,final))
    if 'joint_contact_kN' in exp:close(effective,exp['joint_contact_kN'])
    elif 'contact_normal_kPa' in exp:close(effective,2*exp['contact_normal_kPa'])
    if 'total_normal_kPa' in exp:close(total,2*exp['total_normal_kPa'])
    if 'joint_mean_p_kPa' in exp:close(pressure,2*exp['joint_mean_p_kPa'])
    if 'reaction_y_kN' in exp:close(result['metrics'][2],exp['reaction_y_kN'])
    if 'joint_active_by_stage' in exp:
        for i,active in enumerate(exp['joint_active_by_stage'],1):
            group=[r for r in rows if r[0]==i];assert len(group)==3
            assert all(r[12]==active for r in group),(i,group)
    fluid={}
    if 'joint_fluid_fy_kN' in exp:
        # AUTO=OFF in this fixture: upper-body nodes carry only this JOINT's water load.
        upper_ids=set(f['elements'][1][1:9])
        table=load(root/'水圧確認.json')
        nodal=[r for r in table[35:] if len(r)>20 and isinstance(r[16],(int,float)) and int(r[16]) in upper_ids]
        assert len(nodal)==8
        fx=sum(r[17] for r in nodal);fy=sum(r[18] for r in nodal)
        moment=sum(r[19]*r[18]-r[20]*r[17] for r in nodal)
        close(fx,0.,1e-10);close(fy,exp['joint_fluid_fy_kN'],1e-10);close(moment,exp['joint_fluid_m_kNm'],1e-10)
        close(effective+fy,48.,1e-7)
        fluid={'fx_kN':fx,'fy_kN':fy,'moment_kNm':moment,'GP_pressure_integral_kN':pressure,'contact_plus_actual_fluid_kN':effective+fy}
    max_head=0.;max_q=0.
    if 'head_intercept' in exp:
        table=load(root/'水圧結果.json');gp_rows=[r for r in table[35:] if isinstance(r[0],(float,int))]
        soil_elements={e[0] for e in f['elements'][:2] if f['materials'][e[9]-1][10]=='SOIL'}
        gp_rows=[r for r in gp_rows if int(r[1]) in soil_elements]
        assert len(gp_rows)==4*len(soil_elements),'Missing bulk water data'
        for r in gp_rows:
            h=exp['head_intercept']+exp['head_gradient']*r[3]
            close(r[5],gamma*(h-r[4]),1e-7)
            max_q=max(max_q,abs(r[12]-exp['qx_m_s']),abs(r[13]-exp['qy_m_s']))
        assert max_q<1e-12,max_q
        head_rows=[r for r in table[35:] if len(r)>38 and isinstance(r[32],(float,int))]
        soil_ids={n for e in f['elements'][:2] if f['materials'][e[9]-1][10]=='SOIL' for n in e[1:9]}
        computed=[r for r in head_rows if r[33]!=0 or r[34]]
        assert len(computed)==len(soil_ids),'Missing positive-head solution nodes'
        for r in computed:
            h=exp['head_intercept']+exp['head_gradient']*r[37];max_head=max(max_head,abs(r[33]-h))
        assert max_head<1e-9,max_head
    tangent=[]
    if f.get('checkTangent'):
        tangent=load(root/'tangent.json')
        assert len(tangent)==4
        for row,name in zip(tangent,['STICK','SLIP_POSITIVE','SLIP_NEGATIVE','OPEN']):
            fields=row.split('|');assert fields[0]==name and float(fields[1])<1e-7,row
        assert tangent[1].split('|')[4]!=tangent[2].split('|')[4]
    resume={}
    if (root/'resume.json').exists():
        b=load(root/'resume.json');assert b['status'].startswith('PASS|')
        du=max(abs(x-y) for x,y in zip(result['disp'],b['disp']));ds=max(abs(x-y) for i,(x,y) in enumerate(zip(result['stress'],b['stress'])) if i%17 not in (6,9))
        assert du<1e-9 and ds<1e-6 and b['metrics'][3]<1e-7,(du,ds,b['metrics'])
        restored=[r for r in load(root/'resume_接合結果.json')[1:] if isinstance(r[0],(int,float))];assert len(restored)>=3
        restored=sorted([r for r in restored if r[0]==stage],key=lambda r:r[3]);assert len(restored)==3
        for r,brow in zip(final,restored):
            assert r[10]==brow[10] and r[19]==brow[19]
            for i in range(15,19):close(r[i],brow[i],1e-6)
        resume={'max_du_m':du,'max_physical_stress_error_kPa':ds,'residual':b['metrics'][3],'seconds':b['seconds']}
    return {'status':'PASS','case':f['name'],'seconds':result['seconds'],'effective_contact_kN':effective,'water_support_kN':fluid.get('fy_kN',pressure),'total_contact_kN':fluid.get('contact_plus_actual_fluid_kN',total),'reaction_y_kN':result['metrics'][2],'residual':result['metrics'][3],'Fs':result['metrics'][4:],'max_head_error_m':max_head,'max_flux_error_m_s':max_q,'resume':resume,'tangent':tangent,'fluid_exact':fluid}

def compare(a,b):
    x=load(Path(a)/'result.json');y=load(Path(b)/'result.json')
    delta={key:max(abs(i-j) for i,j in zip(x[key],y[key])) for key in ['disp','stress','metrics']}
    assert all(v==0 for v in delta.values()),delta
    return {'status':'PASS','max_difference':delta,'reference_seconds':x['seconds'],'candidate_seconds':y['seconds']}

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--fixture');p.add_argument('--run');p.add_argument('--compare',nargs=2);p.add_argument('--report',required=True);a=p.parse_args()
    result=compare(*a.compare) if a.compare else run_check(a.fixture,a.run)
    Path(a.report).write_bytes((json.dumps(result,ensure_ascii=False,indent=2)+'\n').encode());print(json.dumps(result,ensure_ascii=False))
