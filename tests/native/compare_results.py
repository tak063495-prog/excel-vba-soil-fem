import json,sys
d=json.load(open(sys.argv[1],encoding='utf-8-sig'))
if isinstance(d,dict): d=[d]
by={}
for r in d: by.setdefault((r['Case'],r['Variant']),{})[r['Mode']]=r
for k,m in by.items():
    if 'OFF' not in m or 'AUTO' not in m: continue
    a=[float(x) for x in m['OFF']['Disp']]; b=[float(x) for x in m['AUTO']['Disp']]
    sc=max(abs(x) for x in a) or 1
    md=max(abs(x-y) for x,y in zip(a,b))
    ra=m['OFF']['Run'].split('|'); rb=m['AUTO']['Run'].split('|')
    same=sum(1 for x,y in zip(a,b) if x==y)
    print(k, 'maxabs',f'{md:.3e}','rel',f'{md/sc:.3e}','bitwise',f'{same}/{len(a)}','t',ra[0],rb[0],'counts',ra[4:8]==rb[4:8],ra[1],rb[1])
