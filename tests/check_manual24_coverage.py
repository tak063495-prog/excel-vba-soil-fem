from pathlib import Path
import re,difflib,json
old=Path('docs/manuals/2DSoilFEM_Manual.md').read_text(encoding='utf-8');new=Path('docs/manuals/2DSoilFEM_Manual_20261024.md').read_text(encoding='utf-8')
# separate header numbers from contents
pat=r'(?ms)^## (\d+)\. [^\n]+\n(.*?)(?=^## |\Z)';oo=dict((int(n),t) for n,t in re.findall(pat,old));nn=dict((int(n),t) for n,t in re.findall(pat,new))
count=0
for ch in range(9,28):
 def result(t):
  m=re.search(r'(?ms)^### 確認値\s*\n(.*?)(?=^### |\Z)',t)
  if not m:return []
  return [r for r in m[1].splitlines() if r.startswith('|')]
 a,b=result(oo[ch]),result(nn[ch]);assert a==b,(ch,list(difflib.unified_diff(a,b)));count+=len(a)
heads=lambda s:re.findall(r'^## (.+)$',s,re.M)
assert heads(old)[:30]==heads(new)[:30]
assert heads(new)[-1]=='37. 付録' and len(heads(new))==37
assert [int(x) for x in re.findall(r'^## \d+\. 解析例(\d+)',new,re.M)]==list(range(1,12))+[11]+list(range(12,24))
print('PASS original30chapters plus37appendix,19old confirmation sections,',count,'confirmation table rows unchanged')
Path('tests/tmp/manual24').mkdir(parents=True,exist_ok=True)
Path('tests/tmp/manual24/coverage.json').write_text(json.dumps(dict(status='PASS',chapters=37,example_chapters=19,numbered_examples=23,additional_example='11-B',unchanged_confirmation_table_rows=count,old_source='docs/manuals/2DSoilFEM_Manual.md',new_source='docs/manuals/2DSoilFEM_Manual_20261024.md'),ensure_ascii=False,indent=2),encoding='utf-8')
