"""Check additive manual editing and render all pages for visual review."""
from pathlib import Path
import re, json, subprocess, hashlib, argparse
import pymupdf
from PIL import Image, ImageDraw

ROOT=Path(__file__).resolve().parents[1]
ap=argparse.ArgumentParser();ap.add_argument('--render',action='store_true');a=ap.parse_args()
source=ROOT/'docs/manuals/2DSoilFEM_Manual.md'
pdf=source.with_name('2DSoilFEM_Manual_20261022_WATER_02.pdf')
out=ROOT/'tests/tmp/joint_water/manual_review';out.mkdir(parents=True,exist_ok=True)
old=subprocess.check_output(['git','show','abc79d3:docs/manuals/2DSoilFEM_Manual.md'],cwd=ROOT).decode('utf-8')
new=source.read_text(encoding='utf-8')
def chapters(s):
    parts=re.split(r'(?m)^(## \d+\. .+)$',s)
    return {int(re.match(r'## (\d+)',parts[i]).group(1)):parts[i+1] for i in range(1,len(parts),2)}
before=chapters(old);after=chapters(new)
preserved=[]
for n in range(2,25):
    target=after[n]

    assert target.strip()==before[n].strip(),f'Unrelated chapter changed: {n}'
    preserved.append(n)
for image in re.findall(r'!\[[^\]]*\]\(([^)]+)\)',new):
    if not image.startswith(('https:','http:')):assert (source.parent/image).exists(),image
with pymupdf.open(pdf) as d:
    toc=d.get_toc();heads=[(h,p) for lv,h,p in toc if lv==1]
    assert len(heads)==26 and all(p>0 for h,p in heads)
    assert heads[-2][0].startswith('25. 水条件') and heads[-1][0]=='26. 付録'
    assert any(h=='用語集' for lv,h,p in toc)
    for h,p in heads:
        compact=lambda s:re.sub(r'\s+','',s)
        assert compact(h) in compact(d[p-1].get_text()),(h,p)
    links=[ln for pg in list(d)[:4] for ln in pg.get_links() if ln.get('kind')==pymupdf.LINK_GOTO]
    assert len(links)>=26
    new_start=heads[-2][1];appendix=heads[-1][1]
    glossary=next(p for lv,h,p in toc if h=='用語集')
    text='\n'.join(p.get_text() for p in list(d)[new_start-1:appendix-1])
    for term in ['HYDROSTATIC','STEADY','EFFECTIVE','EXPOSED','107','137.60','1.5625','固定水圧','CONTINUOUS','PRESCRIBED','JOINT','48 kN','接合結果','17桁']:
        assert term in text,term
    assert '水条件の用語' in '\n'.join(p.get_text() for p in list(d)[glossary-1:])
    if a.render:
        for i,p in enumerate(d):p.get_pixmap(matrix=pymupdf.Matrix(1,1),alpha=False).save(out/f'page_{i+1:03}.png')
        for start in range(0,len(d),20):
            sheet=Image.new('RGB',(1200,1480),'#e8e8e8');draw=ImageDraw.Draw(sheet)
            for j in range(start,min(start+20,len(d))):
                im=Image.open(out/f'page_{j+1:03}.png');im.thumbnail((228,330))
                x=(j-start)%5*240+6;y=(j-start)//5*370+24
                sheet.paste(im,(x,y));draw.text((x,y-18),f'Page {j+1}',fill='black')
            sheet.save(out/f'contact_{start//20+1:02}.png')
    report={'status':'PASS','pages':len(d),'new_chapter_start':new_start,'appendix_start':appendix,'glossary_start':glossary,'chapters_2_to_24_preserved_from_WATER_01':preserved,'toc_links':len(links),'chapters':heads,'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'pdf_sha256':hashlib.sha256(pdf.read_bytes()).hexdigest(),'all_pages_rendered':a.render,'visual_review':'pending'}
(out/'checks.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k:v for k,v in report.items() if k not in ['chapters']},ensure_ascii=True,indent=2))
