"""Build docs/manuals/2DSoilFEM_Manual.md into PDF (two passes for TOC page numbers).
usage: python3 build_manual_pdf.py OUT.pdf   (needs: markdown, google-chrome, pdftotext)"""
import markdown,re,subprocess,sys,os
here=os.path.dirname(os.path.abspath(__file__)); out=sys.argv[1]
md=open(os.path.join(here,'2DSoilFEM_Manual.md'),encoding='utf8').read()
heads=re.findall(r'^## (.+)$',md,re.M)
css="""@page{size:A4;margin:14mm 14mm 16mm}body{font-family:'Noto Sans CJK JP','Noto Sans JP',sans-serif;font-size:10.5pt;line-height:1.7;color:#222}
h1{font-size:21pt;border-bottom:3px solid #2f5d50}h2{font-size:15pt;background:#2f5d50;color:#fff;padding:4px 8px;break-before:page}
h3{font-size:12.5pt;border-left:6px solid #c00;padding-left:8px;margin-top:18px}table{border-collapse:collapse;width:100%;font-size:9pt;margin:6px 0}th,td{border:1px solid #aaa;padding:3px 5px;vertical-align:top}th{background:#e6efe9}
p:has(>img){break-inside:avoid;margin:10px 0 0}img{max-width:100%;max-height:150mm;border:1px solid #888;display:block;margin:0 auto}code{background:#f3f3f3;padding:0 3px}
blockquote{margin:8px 0;padding:6px 10px;background:#f1f6f3;border-left:6px solid #2f5d50;font-size:10pt;break-inside:avoid}blockquote p{margin:3px 0}.toc td{border:none;border-bottom:1px dotted #999;font-size:11pt}"""
def render(toc):
    body=markdown.markdown(md.replace('{{TOC}}',toc),extensions=['tables'])
    h=os.path.join(here,'_manual.html'); open(h,'w',encoding='utf8').write(f"<html><head><meta charset=utf-8><style>{css}</style></head><body>{body}</body></html>")
    subprocess.run(['google-chrome','--headless=new','--no-sandbox','--disable-gpu','--no-pdf-header-footer','--print-to-pdf='+out,'file://'+h],check=True,capture_output=True,timeout=180)
def toc_of(pages): return '<table class="toc"><tr><th>目次</th><th>ページ</th></tr>'+''.join(f'<tr><td>{x}</td><td style="text-align:right">{p}</td></tr>' for x,p in zip(heads,pages))+'</table>'
render(toc_of(['00']*len(heads)))
for _ in range(2):
    txt=subprocess.run(['pdftotext','-layout',out,'-'],capture_output=True,text=True).stdout.split('\f')
    pages=[]
    for hd in heads:
        p=next((i+1 for i,t in enumerate(txt) if i>0 and t.lstrip().startswith(hd)),'?'); pages.append(p)
    render(toc_of(pages))
os.remove(os.path.join(here,'_manual.html')); print(pages)
