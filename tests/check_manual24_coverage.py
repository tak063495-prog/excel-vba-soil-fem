"""Preserve example evidence independently of chapter positions."""
from pathlib import Path
import difflib,json,re
old_path=Path('docs/manuals/2DSoilFEM_Manual.md');new_path=Path('docs/manuals/2DSoilFEM_Manual_20261024.md')
old,new=(p.read_text(encoding='utf8') for p in (old_path,new_path))
def examples(s):
 return {re.match(r'解析例(\d+(?:-B)?)',h)[1]:(int(n),b) for n,h,b in re.findall(r'(?ms)^## (\d+)\. (解析例[^\n]+)\n(.*?)(?=^## |\Z)',s)}
def result(t):
 m=re.search(r'(?ms)^### 確認値\s*\n(.*?)(?=^### |\Z)',t)
 return [r for r in m[1].splitlines() if r.startswith('|')] if m else []
a,b=examples(old),examples(new);count=sections=0
for key,(_,body) in a.items():
 expected,actual=result(body),result(b[key][1])
 # R58-02: preserve the value while correcting its physical interpretation.
 if key=='10':
  expected=[r.replace('|内空高さの減少（天端と下端、半径4.0 mの対向節点間）|23.7871 mm|',
                      '|天端330／下端114のモデル自由度差（工程9−工程1）|23.7871 mm（無効期間をまたぐ参考差分）|') for r in expected]
 assert expected==actual,(key,list(difflib.unified_diff(expected,actual)))
 count+=len(expected);sections+=bool(expected)
chapters=re.findall(r'^## (\d+)\. (.+)$',new,re.M)
assert [int(n) for n,_ in chapters]==list(range(1,38))
assert [n for n,_ in b.values()]==list(range(13,37))
assert list(b)==[str(i) for i in range(1,12)]+['11-B']+[str(i) for i in range(12,24)]
assert chapters[-1][1]=='付録' and new.rfind('### 用語集')>new.rfind('## 37.')
assert not re.search(r'p\.\d+',new),'Use stable chapter references instead of stale printed pages.'
images=lambda text:set(re.findall(r'!\[[^\]]*\]\(([^)]+)\)',text))
baseline_images=json.loads(Path('docs/manual-practice-evidence/20261024/baseline_images.json').read_text(encoding='utf8'))['images']
assert set(baseline_images)<=images(new),'A previously published 1024 image was removed.'
for image in images(new):assert (new_path.parent/image).is_file(),image
for forbidden in ['D14（結果の残し方）','C16:C19','C11:C12=ROLLER','通常画面ではC列にWATERの対象ID','各行はA:Fの6列をすべて入力']:assert forbidden not in new,forbidden
assert '|②|D12|Fs探索の上限' in new and '|④|D8|境界1の格子間隔' in new
assert '水条件 AC12（AC:AL列）' in new and '水条件 V420（V:X列）' in new
assert '第31・' not in new,'Stale time-operation chapter in compound reference'
assert 'Kは載荷セットID' in new and 'accepted/column_smoke/fresh_results.xlsm' in new
report=dict(status='PASS',chapters=37,contiguous_example_chapters=[13,36],numbered_examples=23,additional_example='11-B',confirmation_sections=sections,preserved_numeric_confirmation_rows=count,corrected_definition_rows=1,old_source=str(old_path),new_source=str(new_path))
out=Path('tests/tmp/manual24/coverage.json');out.parent.mkdir(parents=True,exist_ok=True);out.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf8');print('PASS',report)
