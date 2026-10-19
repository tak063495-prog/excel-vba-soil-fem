"""Check balanced hierarchical totals and rank exclusive wall time."""
import argparse
import csv
import json
from pathlib import Path

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('csv'); ap.add_argument('--out')
    args=ap.parse_args()
    rows=list(csv.DictReader(Path(args.csv).read_text(encoding='utf-8-sig').splitlines()))
    root=next(row for row in rows if row['procedure']=='P0_RunAnalysis')
    total=float(root['inclusive_seconds'])
    assert total>0 and int(root['calls'])==1
    summed=sum(float(row['exclusive_seconds']) for row in rows)
    assert abs(summed-total)<max(1e-5,total*1e-7), ('Missing/double-counted profile time',total,summed)
    report=[]
    for row in rows:
        exclusive=float(row['exclusive_seconds']); inclusive=float(row['inclusive_seconds'])
        assert -1e-8<=exclusive<=inclusive+1e-8
        report.append(dict(procedure=row['procedure'],calls=int(row['calls']),inclusive_seconds=inclusive,exclusive_seconds=exclusive,exclusive_percent=100*exclusive/total))
    result=dict(total_seconds=total,exclusive_sum_seconds=summed,exclusive_partition_pass=True,ranking=sorted(report,key=lambda row:-row['exclusive_seconds']))
    text=json.dumps(result,ensure_ascii=False,indent=2)
    print(text)
    if args.out: Path(args.out).write_text(text+'\n',encoding='utf-8')

if __name__=='__main__': main()
