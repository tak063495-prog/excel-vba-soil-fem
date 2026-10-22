from pathlib import Path
import re,subprocess,json,urllib.parse,hashlib,pymupdf
p=Path('docs/manuals/2DSoilFEM_Manual_20261024.pdf');d=pymupdf.open(p);tracked=set(subprocess.check_output(['git','ls-tree','-r','--name-only','origin/main'],text=True,encoding='utf-8').splitlines());links=sorted({x['uri'] for pg in d for x in pg.get_links() if x.get('uri')});repo=[]
for u in links:
 m=re.search(r'(?:github.com/tak063495-prog/excel-vba-soil-fem/(?:blob|tree)/|raw.githubusercontent.com/tak063495-prog/excel-vba-soil-fem/)([^/]+)/(.+)',u)
 if not m:continue
 assert m[1]=='main',u
 target=urllib.parse.unquote(m[2].split('#')[0].split('?')[0]);assert target in tracked or any(t.startswith(target.rstrip('/')+'/') for t in tracked),u;repo.append({'url':u,'path':target,'main_object_verified':True})
overflow=[];blank=[];caption_without_image=[];orphans=[]
source=Path('docs/manuals/2DSoilFEM_Manual_20261024.md').read_text(encoding='utf-8');heads={re.sub(r'\s+','',h) for h in re.findall(r'^### (.+)$',source,re.M)}
for i,pg in enumerate(d):
 body=re.sub(r'\s+','',re.sub(r'\d+ / \d+\s*$','',pg.get_text()))
 if any(body.endswith(h) for h in heads):orphans.append(i+1)
 if len(pg.get_text().strip())<10:blank.append(i+1)
 for b in pg.get_text('blocks'):
  if b[6]==0 and (b[0]<-1 or b[1]<-1 or b[2]>pg.rect.width+1 or b[3]>pg.rect.height+1):overflow.append((i+1,b[:5]))
 if re.search(r'(?m)^図\d+(?:-[\dA-Za-z]+)*[　 ]',pg.get_text()) and not pg.get_images():caption_without_image.append(i+1)
assert not blank and not overflow and not orphans,(blank,overflow,orphans)
out={'status':'PASS','pdf_sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'pages':len(d),'main_file_links_checked':repo,'unique_external_links':len(links),'out_of_page_text_blocks':overflow,'orphan_heading_pages':orphans,'blank_pages':blank,'caption_pages_without_image_for_review':caption_without_image,'source_sha256':hashlib.sha256(Path('docs/manuals/2DSoilFEM_Manual_20261024.md').read_bytes()).hexdigest()}
Path('tests/tmp/manual24').mkdir(parents=True,exist_ok=True)
Path('tests/tmp/manual24/pdf_link_bounds_checks.json').write_text(json.dumps(out,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');print('PASS',len(d),'pages;',len(repo),'main file links; caption-only candidates',caption_without_image)
