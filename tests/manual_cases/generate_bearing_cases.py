"""Generate saved-Q8 bearing-capacity fixtures without opening Excel."""
import json
import math
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent
XCOARSE = [-14,-10,-8,-6,-4,-2,0,.5,1,1.5,2,2.5,3,3.5,4,5,6,8,10,14]
ZCOARSE = [0,.25,.5,1,1.5,2,3,4.5,6,9,12,17]


def mids(values):
    return [v for pair in zip(values[:-1], [(a+b)/2 for a,b in zip(values[:-1], values[1:])]) for v in pair] + [values[-1]]


def surface(x, slope):
    if not slope or x <= -8:
        return 0.0 if slope else 5.0
    if x < 0:
        return (x + 8.0) * 5.0 / 8.0
    return 5.0


def make_case(name, slope, fine):
    xs = mids(XCOARSE) if fine else XCOARSE
    zs = mids(ZCOARSE) if fine else ZCOARSE
    # logical grid nodes are shared by adjacent Q8 elements
    grid = {}
    nodes = []
    def node(ix, iz):
        key = (ix, iz)
        if key in grid: return grid[key]
        x = xs[ix]; z = zs[iz]; y = -12.0 + (1.0 - z / 17.0) * (surface(x, slope) + 12.0)
        side = ix in (0, len(xs)-1)
        bottom = iz == len(zs)-1
        fx = 1 if side or bottom else 0
        fy = 1 if bottom else 0
        row = [float(len(nodes)+1), float(x), float(y), fx, fy, 0, 0, 0, 0]
        nodes.append(row); grid[key] = len(nodes)
        return len(nodes)
    elements = []
    for iz in range(len(zs)-1):
        for ix in range(len(xs)-1):
            a=node(ix,iz); b=node(ix+1,iz); c=node(ix+1,iz+1); d=node(ix,iz+1)
            # bottom/right/top/left midsides; this follows the public Q8 schema.
            # Midside nodes use separate logical keys and are shared by adjacent cells.
            def edge(k, x, z):
                if k not in grid:
                    y=-12.0+(1-z/17.0)*(surface(x,slope)+12.0); grid[k]=len(nodes)+1
                    nodes.append([float(len(nodes)+1),float(x),float(y),0,0,0,0,0,0])
                return grid[k]
            mb=edge(('h',ix,iz), (xs[ix]+xs[ix+1])/2, zs[iz])
            mr=edge(('v',ix+1,iz), xs[ix+1], (zs[iz]+zs[iz+1])/2)
            mt=edge(('h',ix,iz+1), (xs[ix]+xs[ix+1])/2, zs[iz+1])
            ml=edge(('v',ix,iz), xs[ix], (zs[iz]+zs[iz+1])/2)
            # overwrite accidental placeholder only when it is a new, unused key
            # z increases downward, so d-c-b-a is counter-clockwise in X/Y.
            elements.append([float(len(elements)+1),float(d),float(c),float(b),float(a),float(mt),float(mr),float(mb),float(ml),1.0])
    # Remove the harmless logical placeholder created above only if it is not referenced.
    used={int(v) for e in elements for v in e[1:9]}
    nodes=[n for n in nodes if int(n[0]) in used]
    remap={int(n[0]):i+1 for i,n in enumerate(nodes)}
    for i,n in enumerate(nodes,1): n[0]=float(i)
    for e in elements:
        for j in range(1,9): e[j]=float(remap[int(e[j])])
    # restore constraints on generated midsides and all boundary nodes
    for n in nodes:
        x,y=n[1],n[2]
        if abs(y-(-12.0)) < 1e-9: n[3]=1; n[4]=1
        if abs(x-XCOARSE[0]) < 1e-9 or abs(x-XCOARSE[-1]) < 1e-9: n[3]=1
    # Check the Q8 Jacobian at the four Gauss points and an 11x11 diagnostic grid.
    min_det=math.inf; area=0.0
    def q8_det(e, r, s):
        def ns(a,b): return [-.25*(1-a)*(1-b)*(1+a+b), -.25*(1+a)*(1-b)*(1-a+b), -.25*(1+a)*(1+b)*(1-a-b), -.25*(1-a)*(1+b)*(1+a-b), .5*(1-a*a)*(1-b), .5*(1+a)*(1-b*b), .5*(1-a*a)*(1+b), .5*(1-a)*(1-b*b)]
        h=1e-7; n=ns(r,s); nr=ns(r+h,s); nl=ns(r-h,s); nt=ns(r,s+h); nb=ns(r,s-h)
        Ndr=[(nr[k]-nl[k])/(2*h) for k in range(8)]; Nds=[(nt[k]-nb[k])/(2*h) for k in range(8)]
        px=[nodes[int(e[j])-1][1] for j in range(1,9)]; py=[nodes[int(e[j])-1][2] for j in range(1,9)]
        j11=sum(Ndr[k]*px[k] for k in range(8)); j12=sum(Nds[k]*px[k] for k in range(8)); j21=sum(Ndr[k]*py[k] for k in range(8)); j22=sum(Nds[k]*py[k] for k in range(8))
        return j11*j22-j12*j21
    for e in elements:
        p=[nodes[int(e[j])-1][1:3] for j in range(1,5)]
        area += abs((p[0][0]*p[1][1]+p[1][0]*p[2][1]+p[2][0]*p[3][1]+p[3][0]*p[0][1]-p[1][0]*p[0][1]-p[2][0]*p[1][1]-p[3][0]*p[2][1]-p[0][0]*p[3][1])/2)
        for q in (-1/math.sqrt(3),1/math.sqrt(3)):
            for r in (-1/math.sqrt(3),1/math.sqrt(3)): min_det=min(min_det,q8_det(e,r,q))
        for q in range(11):
            for r in range(11): min_det=min(min_det,q8_det(e,-1+q/5,-1+r/5))
    assert min_det > 0 and area > 0 and len({int(n[0]) for n in nodes}) == len(nodes)
    protocol="重力基準の一様追加沈下。単独実行はGRAVITYのみ。run_bearing_cases.pyが先に自重の各footing node変位を取得し、各節点の絶対目標 uy_GRAVITY−追加沈下s に置換してLOADを有効化する。これにより載荷帯の一様な追加沈下を与える。剛体接触/離間/基礎部材は表現しない。"
    loads=[[1.0,"BOX","Y","DISP",-0.02,1.0,5.0,3.0,5.0,"参考値。単独実行はGRAVITYのみ。run_bearing_cases.pyが先に自重の各footing node変位を取得し、各節点の絶対目標 uy_GRAVITY−追加沈下s に置換してLOADを有効化する。"]]
    fixture={"name":name,"materials":[[1,30000,.30,1,18,0,100,0,1,"ON","SOIL",None,None]],"stages":[[1,"GRAVITY",None,None,1,"自重"],[2,"LOAD",1,None,0,"基礎幅2mの重力基準一様追加沈下（wrapperで有効化）"]],"loads":loads,"nodes":nodes,"elements":elements,"metadata":{"mesh_type":"saved_Q8","slope":slope,"fine":fine,"area_m2":area,"min_corner_jacobian":min_det,"protocol":"gravity_referenced_uniform_increment","requires_gravity_reference":True,"footing_min_x":1,"footing_max_x":3,"footing_y":5,"footing_width_m":2,"shape_points":{"x":XCOARSE,"surface":"flat y=5" if not slope else "x<=-8:y=0; -8<x<0:y=(x+8)*5/8; x>=0:y=5","bottom_y":-12,"z":ZCOARSE},"boundaries":"outer contour corresponds to x=-14,14 side and bottom/surface strips; saved nodes/elements are used at runtime","protocol_note":protocol}}
    return fixture


for stem, slope in (("bearing_flat",False),("bearing_slope",True)):
    for fine in (False,True):
        p=OUT/(stem+("_fine" if fine else "")+"_mesh.json")
        p.write_text(json.dumps(make_case(p.stem, slope, fine), ensure_ascii=False, indent=2), encoding="utf-8")
        print(p.name, len(json.loads(p.read_text(encoding="utf-8"))["nodes"]), len(json.loads(p.read_text(encoding="utf-8"))["elements"]))
