import json,glob,os,sys
base=sys.argv[1];out=[]
for f in sorted(glob.glob(os.path.join(base,'*','result.json'))):
    c=os.path.basename(os.path.dirname(f))
    if 'gravity' in c: continue
    t=json.load(open(f,encoding='utf-8-sig'))['tables']
    ids=[r[1] for r in t['載荷'][1:] if r[3]=='DISP']
    st=max(r[0] for r in t['結果節点'][1:] if r[0] is not None)
    rows={r[1]:r for r in t['結果節点'][1:] if r[0]==st}
    ry=sum(rows[i][9] for i in ids); xs=[rows[i][2] for i in ids]
    sr=t['ステージ結果'][-1]
    pl=None
    out.append(dict(case=c,n=len(ids),xmin=min(xs),xmax=max(xs),y=rows[ids[0]][3],RY=ry,q=-ry/(max(xs)-min(xs)),status=sr[3],maxdisp=sr[5],RYall=sr[7],uy=[rows[i][5] for i in ids][:3],msg=sr[9]))
json.dump(out,open(os.path.join(base,'bq_2210.json'),'w'),ensure_ascii=False,indent=1)
for o in out: print(o['case'],o['n'],round(o['xmin'],3),round(o['xmax'],3),o['y'],round(o['RY'],2),round(o['q'],2),o['status'])
