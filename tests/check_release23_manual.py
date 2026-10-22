"""Check Issue35 corrections, published PDF links, preserved chapters and all pages."""
from pathlib import Path
import ast, hashlib, json, re, subprocess
from urllib.parse import urlsplit
import pymupdf
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
source = ROOT/'docs/manuals/2DSoilFEM_Manual.md'
pdf = source.with_name('2DSoilFEM_Manual_20261023_WATER_02.pdf')
out = ROOT/'tests/tmp/release23/manual_review'
out.mkdir(parents=True, exist_ok=True)
old = subprocess.check_output(['git','show','3a1be89:docs/manuals/2DSoilFEM_Manual.md'],cwd=ROOT).decode('utf-8')
new = source.read_text(encoding='utf-8')
def chapters(s):
    parts = re.split(r'(?m)^(## \d+\. .+)$',s)
    return {int(re.match(r'## (\d+)',parts[i]).group(1)):parts[i+1] for i in range(1,len(parts),2)}
before, after = chapters(old), chapters(new)
preserved = []
for n in range(2,25):
    if n in (2,3,4,22): continue
    assert before[n].strip()==after[n].strip(), n
    preserved.append(n)
for image in re.findall(r'!\[[^\]]*\]\(([^)]+)\)',new):
    if not image.startswith(('https:','http:')): assert (source.parent/image).exists(),image

# Exercise the actual conversion function without executing the PDF builder's main body.
builder = source.with_name('build_manual_pdf_v2.py')
tree = ast.parse(builder.read_text(encoding='utf-8'))
function = next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='public_document_links')
from html import escape,unescape
from urllib.parse import unquote,quote
ns = dict(Path=Path,re=re,escape=escape,unescape=unescape,urlsplit=urlsplit,unquote=unquote,quote=quote,__file__=str(builder),here=str(source.parent))
exec(compile(ast.Module(body=[function],type_ignores=[]),str(builder),'exec'),ns)
convert = ns['public_document_links']
assert 'blob/main/docs/GROUNDWATER.md#input' in convert('<a href="../GROUNDWATER.md#input">guide</a>')
for s in ('<a href="#toc">toc</a>','<a href="https://example.com/a">a</a>','<img src="images/cover.png">'):
    assert convert(s)==s
for href in ('file:///C:/private/a.pdf','../../../private.txt','missing-file.zip'):
    try: convert(f'<a href="{href}">bad</a>')
    except (ValueError,FileNotFoundError): pass
    else: raise AssertionError(href)

with pymupdf.open(pdf) as d:
    toc=d.get_toc(); heads=[(h,p) for lv,h,p in toc if lv==1]
    assert len(heads)==26 and heads[-1][0]=='26. 付録'
    compact=lambda s:re.sub(r'\s+','',s)
    for h,p in heads: assert compact(h) in compact(d[p-1].get_text()),(h,p)
    links=[ln for pg in list(d)[:4] for ln in pg.get_links() if ln.get('kind')==pymupdf.LINK_GOTO]
    assert len(links)>=26
    uris=[]
    for i,p in enumerate(d):
        for ln in p.get_links():
            assert ln['kind'] not in (pymupdf.LINK_LAUNCH,pymupdf.LINK_GOTOR),(i+1,ln)
            if ln.get('uri'):
                uri=ln['uri']; assert urlsplit(uri).scheme in ('https','http','mailto'),(i+1,uri)
                assert not re.search(r'Users[/\\]|Downloads[/\\]|file:',uri,re.I),(i+1,uri)
                uris.append({'page':i+1,'uri':uri})
    p5=d[4].get_text()
    assert '2DSoilFEM_ui_20261023.xlsm' in p5 and 'WATER_02' in p5
    assert any(x['page']==5 and 'SoilFEM_UI_20261023.zip' in x['uri'] for x in uris)
    text='\n'.join(p.get_text() for p in d)
    for term in ('joint_water_wet.xlsm','未計算','TOTAL','EFFECTIVE','112 kN','38 kN'):
        assert term in text,term
    for i,p in enumerate(d): p.get_pixmap(matrix=pymupdf.Matrix(1,1),alpha=False).save(out/f'page_{i+1:03}.png')
    for start in range(0,len(d),20):
        sheet=Image.new('RGB',(1200,1480),'#e8e8e8');draw=ImageDraw.Draw(sheet)
        for j in range(start,min(start+20,len(d))):
            im=Image.open(out/f'page_{j+1:03}.png');im.thumbnail((228,330))
            x=(j-start)%5*240+6;y=(j-start)//5*370+24
            sheet.paste(im,(x,y));draw.text((x,y-18),f'Page {j+1}',fill='black')
        sheet.save(out/f'contact_{start//20+1:02}.png')
    report=dict(status='PASS',pages=len(d),chapters=heads,glossary_start=next(p for lv,h,p in toc if h=='用語集'),preserved_chapters_from_3a1be89=preserved,review_related_chapters=[2,3,4,22,25],toc_links=len(links),public_uris=uris,local_uris=0,link_conversion_regression='PASS',all_pages_rendered=True,visual_review='pending',source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),pdf_sha256=hashlib.sha256(pdf.read_bytes()).hexdigest())
(out/'checks.json').write_bytes((json.dumps(report,ensure_ascii=False,indent=2)+'\n').encode())
print(json.dumps({k:v for k,v in report.items() if k not in ('public_uris','chapters')},ensure_ascii=True))
