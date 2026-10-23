"""Check actual Excel issue fixtures and checkpoint fields, independently of VBA."""
import argparse, csv, hashlib, json, math
from pathlib import Path

def read(p): return json.loads(p.read_text(encoding='utf-8-sig'))
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def rows(p): return list(csv.DictReader(p.open(encoding='utf-8-sig')))
def close(a,b,tol=1e-9):
    assert len(a)==len(b)
    delta=max((abs(float(x)-float(y)) for x,y in zip(a,b)),default=0.)
    assert delta<=tol, (delta,tol)
    return delta
def water(p):
    lines=p.read_text(encoding='utf-8-sig').splitlines();i=lines.index('element,gauss,pressure_kPa')
    return [float(s.split(',')[2]) for s in lines[i+1:] if s.strip()]
def final_frames(folder):
    roots=sorted(folder.glob('*_out/time_*'))
    return [max((d for d in r.glob('C*') if d.is_dir()),key=lambda d:d.name) for r in roots]

def check_reset_gravity_stress(model,stress):
    """Physical stress, excluding large yield sentinels, before rejected RESET."""
    xy={int(n[0]):n[1:3] for n in model['nodes']};H=max(p[1] for p in xy.values())
    materials={int(m[0]):m for m in model['materials']};a=1/math.sqrt(3)
    worst=0.
    for e,element in enumerate(model['elements']):
        material=materials[int(element[9])];nu,gamma=material[2],material[4];pred=[]
        ys=[xy[int(n)][1] for n in element[1:5]]
        for x,y in ((-a,-a),(a,-a),(-a,a),(a,a)):
            if len(element)>10 and element[10]=='T6':
                r=(1+x)*(1-y)/4;s=(1+y)/2;yg=ys[0]*(1-r-s)+ys[1]*r+ys[2]*s
            else:yg=sum(v*w for v,w in zip(ys,((1-x)*(1-y)/4,(1+x)*(1-y)/4,(1+x)*(1+y)/4,(1-x)*(1+y)/4)))
            sy=gamma*(H-yg);sx=nu/(1-nu)*sy;pred.append([sx,sy,0,sx])
        actual=[[stress[e*68+g*17+j] for j in (0,1,2,16)] for g in range(4)]
        worst=max(worst,close([v for row in sorted(actual,key=lambda r:r[1]) for v in row],
                              [v for row in sorted(pred,key=lambda r:r[1]) for v in row],1e-8))
    return worst

def check(root,fixtures,book,partial=None,reset_run=None):
    reports=[]
    for f in sorted(fixtures.glob('*.json')):
        model=read(f);folder=root/f.stem
        if f.stem=='reset_t6_nonaffine' and reset_run: folder=reset_run
        rejected='expectedAnalysisError' in model
        d=read(folder/('expected_failure.json' if rejected else 'result.json'))
        assert d['fixture_sha256'].lower()==sha(f),f
        assert d['book_sha256'].lower()==sha(book),folder
        assert (d['status'].startswith('PASS|')) != rejected, (folder,d['status'])
        record={'case':f.stem,'status':'EXPECTED_REJECTION' if rejected else 'PASS'}
        if rejected: assert model['expectedAnalysisError'] in d['status']
        if f.stem.startswith('water_'):
            profile='W' if f.stem=='water_hold' else 'W2'
            levels=[float(r[2]) for r in model['water']['profiles'] if r[0]==profile]
            assert levels and max(levels)-min(levels)<1e-12
            top=max(n[2] for n in model['nodes'])
            width=max(n[1] for n in model['nodes'])-min(n[1] for n in model['nodes'])
            expected=d['metrics'][0]+model['water']['gamma']*(levels[0]-top)*width
            record['reaction_analytic_error_kN']=close([d['metrics'][2]],[expected],1e-6)
        if (folder/'resume.json').exists():
            resumed=read(folder/'resume.json');assert resumed['status'].startswith('PASS|')
            record['resume_u_max_m']=close(d['disp'],resumed['disp'])
            record['resume_Fs_max']=close(d['metrics'][4:],resumed['metrics'][4:],1e-10)
            record['resume_reaction_max_kN']=close(d['metrics'][1:3],resumed['metrics'][1:3],1e-6)
            frames=final_frames(folder);assert len(frames)==2,frames
            record['resume_gp_pressure_max_kPa']=close(water(frames[0]/'water.csv'),water(frames[1]/'water.csv'),1e-6)
            for name,fields in [('node.csv',('ux','uy')),('dof.csv',('locked_sw','locked_ap','bound_disp','cond'))]:
                a,b=rows(frames[0]/name),rows(frames[1]/name);assert len(a)==len(b)
                for field in fields: close([r[field] for r in a],[r[field] for r in b],1e-8)
            if f.stem.startswith('named_'):
                expected=-1 if f.stem=='named_disp' else -4
                for frame in frames:
                    dofs=rows(frame/'dof.csv');ap=[float(r['locked_ap']) for r in dofs]
                    close(ap,[expected if i==77 else 0 for i in range(len(ap))],1e-10)
                    if f.stem=='named_disp':
                        # U explicitly changes the A DOF back to FORCE=0.
                        assert int(dofs[79]['cond'])==0
                        close([dofs[79]['bound_disp']],[0],1e-12)
                if f.stem=='named_disp':
                    wait=next(folder.glob('*_out/stg04_WAIT/dof.csv'))
                    dofs=rows(wait);assert int(dofs[79]['cond'])==1
                    close([dofs[79]['bound_disp']],[-.0013],1e-12)
        if f.stem=='conforming_q8':
            output=read(folder/'結果節点.json')
            record['patch_u_max_m']=close([v for r in output[1:] if r[1] for v in r[4:6]],
                [v for r in output[1:] if r[1] for v in (.001*r[2],0)],1e-12)
        if f.stem=='reset_t6_nonaffine':
            close([v for r in d['reference_coordinates'] for v in r[1:]],
                  [v for r in model['nodes'] for v in r[1:3]],1e-12)
            # This two-T6 column represents the quadratic self-weight field
            # exactly: lateral strain is zero, with a fixed bottom.
            E,nu=float(model['materials'][0][1]),float(model['materials'][0][2])
            gamma=float(model['materials'][0][4]);H=max(n[2] for n in model['nodes'])
            md=E*(1-nu)/((1+nu)*(1-2*nu))
            record['preserved_u_max_m']=close([v for r in d['nodes'] for v in r[1:]],
                [v for r in model['nodes'] for v in (0,-gamma/md*(H*r[2]-.5*r[2]**2))],1e-9)
            assert max(abs(v) for r in d['nodes'] for v in r[1:])>1e-8
            record['preserved_stress_analytic_error_kPa']=check_reset_gravity_stress(model,d['stress'])
        if f.stem=='reset_t6_affine':
            output=read(folder/'結果節点.json');final=[r for r in output[1:] if r[0]==2]
            assert len(final)==len(model['nodes'])
            close([v for r in final for v in r[2:6]],
                  [v for r in model['nodes'] for v in (.9999*r[1],.9998*r[2],0,0)],1e-12)
        if f.stem=='apex_native':
            plastic=read(folder/'結果塑性.json')[1:];assert len(plastic)==4
            close([r[9] for r in plastic],[.02]*4,1e-12)
            close([v for r in plastic for v in r[5:9]],[-.01,-.01,0,0]*4,1e-12)
        reports.append(record)
    if partial:
        d,r=read(partial/'result.json'),read(partial/'resume.json')
        assert d['status'].startswith('PASS|') and r['status'].startswith('PASS|')
        u=close(d['disp'],r['disp']);reac=close(d['metrics'][1:3],r['metrics'][1:3],1e-6)
        fs=final_frames(partial);assert len(fs)==2
        p=close(water(fs[0]/'water.csv'),water(fs[1]/'water.csv'),1e-6)
        # Stage 3 resumes from its accepted -2 kN intermediate value but its
        # end target is the B set (-4 kN at original Y-DOF 77).  The following
        # WAIT must inherit that end target; fresh/resume agreement alone could
        # otherwise miss a shared stale target.
        expected=[-4 if i==77 else 0 for i in range(86)]
        for frame in fs:
            applied=[float(row['locked_ap']) for row in rows(frame/'dof.csv')]
            close(applied,expected,1e-10)
        output=list(partial.glob('*_out'));assert len(output)==1,output
        # The resumed stage-4 WAIT retains A=-2 as well as completed B=-4.
        # Stage 6 then retains B alone after the explicit A unload.
        wait4=[float(row['locked_ap']) for row in rows(output[0]/'stg04_WAIT'/'dof.csv')]
        close(wait4,[-4 if i==77 else -2 if i==79 else 0 for i in range(86)],1e-10)
        wait6=[float(row['locked_ap']) for row in rows(output[0]/'stg06_WAIT'/'dof.csv')]
        close(wait6,expected,1e-10)
        reports.append({'case':'named_linear_partial','status':'PASS','resume_u_max_m':u,'reaction_max_kN':reac,'pressure_max_kPa':p,
                        'stage4_B_locked_applied_kN':wait4[77],'stage6_B_locked_applied_kN':wait6[77]})
    return {'status':'PASS','book_sha256':sha(book),'cases':reports}

if __name__=='__main__':
    ap=argparse.ArgumentParser()
    for name in ('root','book','out'): ap.add_argument('--'+name,type=Path,required=True)
    ap.add_argument('--fixtures',type=Path,default=Path('tests/fixtures/issue49_57'))
    ap.add_argument('--partial',type=Path)
    ap.add_argument('--reset-run',type=Path,help='Accepted re-run after fixing the coordinate-collection driver')
    a=ap.parse_args();r=check(a.root,a.fixtures,a.book,a.partial,a.reset_run)
    a.out.write_text(json.dumps(r,indent=2)+'\n',encoding='utf-8');print('PASS',len(r['cases']),'cases')
