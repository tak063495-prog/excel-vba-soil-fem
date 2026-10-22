"""Map printed-page references to matching content in the same chapter."""
from pathlib import Path
import argparse,hashlib,json,re
import pymupdf

def grams(text):
 text=re.sub(r'\s+|[0-9]','',text)
 return set(text[i:i+10] for i in range(len(text)-9))
def chapters(doc):
 rows=[r for r in doc.get_toc() if r[0]==1]
 return [(int(re.match(r'\d+',r[1])[0]),r[2],rows[i+1][2]-1 if i+1<len(rows) else len(doc)) for i,r in enumerate(rows)]
def main():
 ap=argparse.ArgumentParser();ap.add_argument('old',type=Path);ap.add_argument('new',type=Path);ap.add_argument('source',type=Path);ap.add_argument('report',type=Path);ap.add_argument('--apply',action='store_true');a=ap.parse_args()
 if a.apply and a.report.exists():raise RuntimeError('Preserve the existing mapping; do not apply twice.')
 old,new=pymupdf.open(a.old),pymupdf.open(a.new);oc,nc=chapters(old),chapters(new)
 md=a.source.read_text(encoding='utf-8');original=hashlib.sha256(a.source.read_bytes()).hexdigest();ng={i:grams(p.get_text()) for i,p in enumerate(new)};mapping=[]
 for page in sorted(set(map(int,re.findall(r'p\.(\d+)',md)))):
  ch=next(c for c,start,end in oc if start<=page<=end)
  targets=([ch] if len(oc)==len(nc) else ([ch+5 if ch>=23 else ch] if ch!=8 else [8,23,24,25,26,27]))
  ranges=[range(start-1,end) for c,start,end in nc if c in targets]
  og=grams(old[page-1].get_text());score,target=max((len(og&ng[i])/max(1,len(og)),i+1) for ran in ranges for i in ran)
  assert score>.5,(page,score)
  mapping.append(dict(old=page,new=target,chapter=ch,matching_text_fraction=score))
 if a.apply:
  lookup={r['old']:r['new'] for r in mapping}
  a.source.write_text(re.sub(r'p\.(\d+)',lambda x:'p.'+str(lookup[int(x[1])]),md),encoding='utf-8')
 a.report.parent.mkdir(parents=True,exist_ok=True)
 a.report.write_text(json.dumps(dict(status='PASS',source_before_sha256=original,source_after_sha256=hashlib.sha256(a.source.read_bytes()).hexdigest(),mapping=mapping),ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
 print('PASS:',len(mapping),'printed page targets matched within their chapters')
if __name__=='__main__':main()
