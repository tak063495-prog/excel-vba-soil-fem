"""Read-only independent checks for a saved mixed Q8/T6 XLSM.

This deliberately reads worksheet XML from the XLSM ZIP and does not open or
resave the workbook in Excel.  It checks geometry independently of VBA's
metrics.
"""
import argparse, hashlib, json, math, re, zipfile
from xml.etree import ElementTree as ET

NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

def colnum(ref):
    return sum((ord(c)-64)*26**i for i,c in enumerate(reversed(re.match(r'[A-Z]+', ref).group())))

def read_book(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('xl/workbook.xml'))
        relroot = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rels = {x.attrib['Id']: x.attrib['Target'] for x in relroot}
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            sr=ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in sr.findall('m:si',NS): shared.append(''.join(t.text or '' for t in si.iter('{%s}t'%NS['m'])))
        sheets={}
        for sh in root.find('m:sheets',NS):
            target=rels[sh.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]
            target=target.lstrip('/')
            if not target.startswith('xl/'):
                target='xl/'+target
            data={}
            ws=ET.fromstring(z.read(target))
            for row in ws.findall('.//m:sheetData/m:row',NS):
                vals={}
                for cell in row.findall('m:c',NS):
                    ref=cell.attrib['r']; c=colnum(ref)
                    v=cell.find('m:v',NS); value='' if v is None else v.text or ''
                    if cell.attrib.get('t')=='s' and value!='': value=shared[int(value)]
                    elif cell.attrib.get('t')=='inlineStr': value=''.join(t.text or '' for t in cell.iter('{%s}t'%NS['m']))
                    vals[c]=value
                data[int(row.attrib['r'])]=vals
            sheets[sh.attrib['name']]=data
    return sheets

def num(v): return float(v)
def area(a,b,c): return ((b[0]-a[0])*(c[1]-a[1])-(b[1]-a[1])*(c[0]-a[0]))/2
def tri_angle_min(p):
    out=180.0
    for i in range(3):
        a=p[i-1]; b=p[i]; c=p[(i+1)%3]
        u=(a[0]-b[0],a[1]-b[1]); v=(c[0]-b[0],c[1]-b[1])
        cs=max(-1,min(1,(u[0]*v[0]+u[1]*v[1])/(math.hypot(*u)*math.hypot(*v))))
        out=min(out,math.degrees(math.acos(cs)))
    return out

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--book',required=True); ap.add_argument('--out'); a=ap.parse_args()
    sheets=read_book(a.book); nodes={}
    for r,v in sheets['節点データ'].items():
        if r==1 or not v.get(1): continue
        nodes[int(v[1])] = (num(v[2]),num(v[3]))
    elems=[]
    for r,v in sheets['要素データ'].items():
        if r==1 or not v.get(1): continue
        kind=str(v.get(11,'Q8')).upper() or 'Q8'
        ids=[int(v[c]) for c in range(2,10)]
        elems.append((int(v[1]),kind,ids,int(v[10])))
    checks=[]; fail=[]
    def ck(name, ok, detail=''):
        checks.append({'name':name,'ok':bool(ok),'detail':detail})
        if not ok: fail.append(name)
    q8=[e for e in elems if e[1]!='T6']; t6=[e for e in elems if e[1]=='T6']
    ck('element counts',len(q8)==1 and len(t6)==2,f'Q8={len(q8)} T6={len(t6)}')
    areas=[]; real=set()
    for _,kind,ids,_ in elems:
        if kind=='T6':
            p=[nodes[ids[i]] for i in (0,1,2)]; ar=area(*p); ck(f'T6 positive {ids[0]}',ar>0,f'{ar:g}'); ck('T6 angle',tri_angle_min(p)>=10, f'{tri_angle_min(p):g}')
            real.update(ids[i] for i in (0,1,2,4,5,7))
        else:
            p=[nodes[ids[i]] for i in range(4)]; ar=area(p[0],p[1],p[2])+area(p[0],p[2],p[3]); real.update(ids)
        areas.append(ar)
    ck('total area',math.isclose(sum(areas),2.5,rel_tol=1e-10,abs_tol=1e-10),f'{sum(areas):g}')
    ck('node count',len(nodes)==14,str(len(nodes))); ck('no unused nodes',len(real)==len(nodes),f'used={len(real)}')
    mids={}; edge_count={}; boundary_points=set()
    for ident,kind,ids,material in elems:
        ck('known element type',kind in ('Q8','T6'))
        if kind=='T6':
            ck(f'T6 slots {ident}',ids[2]==ids[3]==ids[6] and len(set(ids[i] for i in (0,1,2,4,5,7)))==6)
            edges=((ids[0],ids[1],ids[4]),(ids[1],ids[2],ids[5]),(ids[2],ids[0],ids[7]))
        else:
            ck(f'Q8 unique nodes {ident}',len(set(ids))==8)
            edges=tuple((ids[j],ids[(j+1)%4],ids[j+4]) for j in range(4))
        for a1,b1,m in edges:
            edge=tuple(sorted((a1,b1)))
            ck(f'midpoint coordinates {ident}/{a1}-{b1}',
               max(abs(nodes[m][k]-(nodes[a1][k]+nodes[b1][k])/2) for k in (0,1)) < 1e-12)
            if edge in mids: ck(f'shared midpoint ID {a1}-{b1}',mids[edge]==m)
            mids[edge]=m
            edge_count[edge]=edge_count.get(edge,0)+1
    ck('node IDs',set(nodes)==set(range(1,15)))
    ck('element IDs',{e[0] for e in elems}=={1,2,3})
    ck('all real nodes used',real==set(nodes))
    ck('conforming corner edges',max(edge_count.values(),default=0)<=2)
    ck('outer boundary',sum(n==1 for n in edge_count.values())==6)
    for edge,count in edge_count.items():
        if count==1: boundary_points.update(nodes[n] for n in edge)
    ck('outer contour coordinates',boundary_points=={(0.,0.),(1.,0.),(2.,0.),(3.,.05),(2.,1.),(0.,1.)})
    shared={frozenset(nodes[n] for n in edge) for edge,count in edge_count.items() if count==2}
    ck('two correct interior edges',shared=={frozenset(((1.,0.),(2.,1.))),frozenset(((2.,0.),(2.,1.)))})
    ck('original good Q8 preserved',q8[0][0]==1 and q8[0][2]==list(range(1,9)) and q8[0][3]==1)
    # With verified straight midpoint coordinates, Q8 geometry reduces to a
    # bilinear corner map. Check its determinant independently on 25 points.
    min_j=math.inf
    for _,kind,ids,_ in q8:
        p=[nodes[n] for n in ids[:4]]
        for xi in (-1.,-.5,0.,.5,1.):
            for eta in (-1.,-.5,0.,.5,1.):
                dx=(-(1-eta)/4,(1-eta)/4,(1+eta)/4,-(1+eta)/4)
                de=(-(1-xi)/4,-(1+xi)/4,(1+xi)/4,(1-xi)/4)
                j00=sum(dx[k]*p[k][0] for k in range(4));j01=sum(de[k]*p[k][0] for k in range(4))
                j10=sum(dx[k]*p[k][1] for k in range(4));j11=sum(de[k]*p[k][1] for k in range(4))
                min_j=min(min_j,j00*j11-j01*j10)
    ck('Q8 positive Jacobian on 25 points',min_j>0,str(min_j))
    q8_angles=[]
    for _,kind,ids,_ in q8:
        corners=[nodes[ids[i]] for i in range(4)]
        for j,b0 in enumerate(corners):
            a0=corners[j-1]; c0=corners[(j+1)%4]
            u=(a0[0]-b0[0],a0[1]-b0[1]); v=(c0[0]-b0[0],c0[1]-b0[1])
            q8_angles.append(math.degrees(math.acos(max(-1,min(1,(u[0]*v[0]+u[1]*v[1])/(math.hypot(*u)*math.hypot(*v)))))))
    ck('Q8 angle repaired',max(q8_angles,default=0)<170,f'max={max(q8_angles,default=0):g} deg')
    materials={int(v[1]):v for r,v in sheets['材料データ'].items() if r>1 and v.get(1)}
    weight=sum(ar*num(materials[e[3]][5])*num(materials[e[3]][4]) for ar,e in zip(areas,elems))
    ck('material gamma',all(num(materials[e[3]][5])==18 for e in elems))
    ck('material thickness',all(num(materials[e[3]][4])==1 for e in elems))
    ck('weight',math.isclose(weight,45,rel_tol=1e-10),f'{weight:g}')
    result={'status':'PASS' if not fail else 'FAIL','book':a.book,'sha256':hashlib.sha256(open(a.book,'rb').read()).hexdigest(),'areas_m2':areas,'weight_kn':weight,'q8_min_jacobian':min_j,'checks':checks}
    text=json.dumps(result,ensure_ascii=False,indent=2)
    if a.out: open(a.out,'w',encoding='utf-8').write(text+'\n')
    print(text)
    return 0 if not fail else 1
if __name__=='__main__': raise SystemExit(main())
