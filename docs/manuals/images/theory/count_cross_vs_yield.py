import csv,sys,math,collections
def run(d):
    g=collections.defaultdict(list)
    for r in csv.DictReader(open(d+'/gauss.csv')): g[int(r['elem'])].append(r)
    cum=collections.defaultdict(float)
    for r in csv.DictReader(open(d+'/plastic_cumulative.csv')): cum[int(r['elem'])]+=abs(float(r['mult_cumulative']))
    P={};Y={};GPmax={}
    for e,rs in g.items():
        ex=sum(float(r['ex_pl']) for r in rs)/4;ey=sum(float(r['ey_pl']) for r in rs)/4;gx=sum(float(r['gxy_pl']) for r in rs)/4
        c=(ex+ey)/2;R=math.hypot((ex-ey)/2,gx/2);P[e]=max(abs(c+R),abs(c-R))
        Y[e]=sum(int(r['yielded']) for r in rs)
        GPmax[e]=max(max(abs(float(r[k])) for k in('ex_pl','ey_pl','gxy_pl')) for r in rs)
    m=max(P.values())
    yl=[e for e in g if Y[e]>0]; cr=[e for e in g if P[e]>0]
    print(d,'elements',len(g),'maxP %.3e'%m)
    print(' yielded-GP elements',len(yl),'| nonzero cross',len(cr))
    print(' yielded but cross==0:',sum(1 for e in yl if P[e]==0),'| yielded but cross<5%% of max:',sum(1 for e in yl if 0<P[e]<0.05*m))
    print(' cumulative plastic>0 but cross==0:',sum(1 for e in g if cum[e]>0 and P[e]==0),'| cum>0 but not yielded now:',sum(1 for e in g if cum[e]>0 and Y[e]==0))
    print(' nonzero GP plastic but averaged <10% of GP max (cancellation/partial):',sum(1 for e in g if GPmax[e]>0 and P[e]<0.1*GPmax[e]))
    print(' cross>0 but no yielded GP:',sum(1 for e in cr if Y[e]==0))
for d in sys.argv[1:]: run(d)
