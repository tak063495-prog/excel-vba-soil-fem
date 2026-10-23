"""Build a manual markdown into PDF with TOC page numbers, internal TOC links,
PDF bookmarks (chapters and sections) and printed page numbers.
New file (2026-10-21); build_manual_pdf.py is kept unchanged.
usage: python3 build_manual_pdf_v2.py MANUAL.md OUT.pdf
needs: markdown, google-chrome, pdftotext, pymupdf"""
import markdown,re,subprocess,sys,os,tempfile,shutil
from pathlib import Path
import pymupdf
from html import escape, unescape
from urllib.parse import urlsplit, unquote, quote
src,out=sys.argv[1],sys.argv[2]
here=os.path.dirname(os.path.abspath(src))
md=open(src,encoding='utf8').read()
chrome=os.environ.get('CHROME_PATH') or shutil.which('google-chrome') or shutil.which('chrome')
if not chrome:
    chrome=next((p for p in [r'C:\Program Files\Google\Chrome\Application\chrome.exe',r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'] if os.path.isfile(p)),None)
if not chrome: raise RuntimeError('Chrome/Edge executable not found; set CHROME_PATH')
heads=re.findall(r'^## (.+)$',md,re.M)
css="""@page{size:A4;margin:14mm 14mm 18mm}body{font-family:'Noto Sans CJK JP','Noto Sans JP',sans-serif;font-size:10.5pt;line-height:1.7;color:#222}
h1{font-size:21pt;border-bottom:3px solid #2f5d50}h2{font-size:15pt;background:#2f5d50;color:#fff;padding:4px 8px;break-before:page}
h3{break-after:avoid-page;page-break-after:avoid;font-size:12.5pt;border-left:6px solid #c00;padding-left:8px;margin-top:18px}table{border-collapse:collapse;width:100%;font-size:9pt;margin:6px 0}th,td{border:1px solid #aaa;padding:3px 5px;vertical-align:top}th{background:#e6efe9}
p:has(>img){break-inside:avoid;margin:10px 0 0}.figure{break-inside:avoid}img{max-width:100%;max-height:150mm;border:1px solid #888;display:block;margin:0 auto}code{background:#f3f3f3;padding:0 3px}
blockquote{margin:8px 0;padding:6px 10px;background:#f1f6f3;border-left:6px solid #2f5d50;font-size:10pt;break-inside:avoid}blockquote p{margin:3px 0}.toc td{border:none;border-bottom:1px dotted #999;font-size:11pt}"""
tmp=os.path.join(here,'_manual_v2.html')
def public_document_links(body):
    """Publish document links; keep image sources local for rendering."""
    root = Path(__file__).resolve().parents[2]
    public_ref = os.environ.get('MANUAL_PUBLIC_REF', 'main')
    base = 'https://github.com/tak063495-prog/excel-vba-soil-fem/blob/' + quote(public_ref, safe='') + '/'
    def replace_link(match):
        href = unescape(match.group(3))
        url = urlsplit(href)
        if url.scheme or url.netloc:
            if url.scheme.lower() not in ('http', 'https', 'mailto'):
                raise ValueError('Non-public document link: ' + href)
            return match.group(0)
        if not url.path:
            return match.group(0)
        target = (Path(here) / unquote(url.path)).resolve()
        relative = target.relative_to(root)
        if not target.exists():
            raise FileNotFoundError(target)
        public = base + quote(relative.as_posix(), safe='/')
        if url.query: public += '?' + url.query
        if url.fragment: public += '#' + url.fragment
        return match.group(1) + match.group(2) + escape(public, quote=True) + match.group(2)
    return re.sub(r"(<a\b[^>]*\bhref\s*=\s*)([\"'])(.*?)\2", replace_link, body, flags=re.I | re.S)

def render(toc):
    body=markdown.markdown(md.replace('{{TOC}}',toc),extensions=['tables'])
    body=re.sub(r'(<p>\s*<img\b[^>]*>\s*</p>)\s*(<p>図[^<]*(?:<(?!/?p\b)[^>]*>[^<]*)*</p>)',r'<div class="figure">\1\2</div>',body)
    body=public_document_links(body)
    open(tmp,'w',encoding='utf8').write(f"<html><head><meta charset=utf-8><style>{css}</style></head><body>{body}</body></html>")
    subprocess.run([chrome,'--headless=new','--no-sandbox','--disable-gpu','--no-pdf-header-footer','--print-to-pdf='+os.path.abspath(out),Path(tmp).as_uri()],check=True,capture_output=True,timeout=300)
def toc_of(pages): return '<table class="toc"><tr><th>目次</th><th>ページ</th></tr>'+''.join(f'<tr><td>{x}</td><td style="text-align:right">{p}</td></tr>' for x,p in zip(heads,pages))+'</table>'
def find_pages():
    # Explicit UTF-8 prevents Japanese heading/bookmark corruption on Windows.
    if shutil.which('pdftotext'):
        txt=subprocess.run(['pdftotext','-enc','UTF-8','-layout',out,'-'],capture_output=True,text=True,encoding='utf-8',check=True).stdout.split('\f')
    else:
        with pymupdf.open(out) as d: txt=[p.get_text(sort=True) for p in d]
    compact=lambda s:re.sub(r'\s+','',s)
    # A continued TOC page can begin with a chapter title; choose its body occurrence.
    pages=[max((i+1 for i,t in enumerate(txt) if i>0 and compact(t).startswith(compact(h))),default=0) for h in heads]
    if not all(pages[i]>pages[i-1] for i in range(1,len(pages))) or not all(pages):
        raise RuntimeError('Chapter body page detection failed: '+str(pages))
    return pages,txt
render(toc_of(['00']*len(heads)))
for _ in range(2):
    pages,txt=find_pages(); render(toc_of(pages))
pages,txt=find_pages()
os.remove(tmp)
doc=pymupdf.open(out)
# section (###) pages: first page after its chapter page whose text has the heading line
toc=[];sub=re.findall(r'^(##|###) (.+)$',md,re.M);cur=1
for lv,h in sub:
    if lv=='##':
        p=pages[heads.index(h)];cur=p if isinstance(p,int) else cur;toc.append([1,h,cur])
    else:
        p=next((i+1 for i in range(cur-1,len(txt)) if any(l.strip()==h for l in txt[i].splitlines())),None)
        if p: toc.append([2,h,p]);cur=p
doc.set_toc(toc)
if os.environ.get('MANUAL_PDF_TITLE'):
    metadata = doc.metadata
    metadata.update(title=os.environ['MANUAL_PDF_TITLE'], author='2DSoilFEM', subject='操作・解析例・理論・用語集')
    doc.set_metadata(metadata)
# internal links on the TOC page(s)
for pno in range(min(4,len(doc))):
    pg=doc[pno]
    for h,p in zip(heads,pages):
        if not isinstance(p,int):continue
        for r in pg.search_for(h)[:1]:
            pg.insert_link({'kind':pymupdf.LINK_GOTO,'from':pymupdf.Rect(r.x0,r.y0,pg.rect.width-40,r.y1),'page':p-1})
# printed page numbers (same as the TOC numbers = physical page)
n=len(doc)
for i,pg in enumerate(doc):
    if i==0:continue
    w=pg.rect.width;y=pg.rect.height-22
    pg.insert_text((w/2-18,y),f'{i+1} / {n}',fontsize=8.5,fontname='helv',color=(.35,.35,.35))
tmp2=out+'.tmp.pdf';doc.save(tmp2,garbage=3,deflate=True);doc.close();os.replace(tmp2,out)
print('pages',n,'chapters',pages)
