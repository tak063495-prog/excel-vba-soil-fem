"""Check manual structure/links and render every page for visual review."""
from pathlib import Path
import argparse,json,re
import pymupdf
from PIL import Image,ImageDraw

def main():
 ap=argparse.ArgumentParser();ap.add_argument('pdf',type=Path);ap.add_argument('out',type=Path);ap.add_argument('--source',type=Path,default=Path('docs/manuals/2DSoilFEM_Manual.md'));a=ap.parse_args()
 a.out.mkdir(parents=True,exist_ok=True);md=a.source.read_text(encoding='utf-8');d=pymupdf.open(a.pdf)
 chapters=re.findall(r'^## (\d+)\. (.+)$',md,re.M)
 assert [int(n) for n,t in chapters]==list(range(1,27))
 assert '水' in chapters[7][1] and '付録' in chapters[-1][1]
 images=re.findall(r'!\[[^\]]*\]\(([^)]+)\)',md)+re.findall(r'<img[^>]+src="([^"]+)"',md)
 for p in images:assert (a.source.parent/p).exists(),p
 toc=d.get_toc();ct=[r for r in toc if r[0]==1]
 assert len(ct)==26 and [int(re.match(r'\d+',r[1])[0]) for r in ct]==list(range(1,27))
 assert len({r[2] for r in ct})==26
 links=[link.get('uri','') for p in d for link in p.get_links() if link.get('uri')]
 assert not any(x.lower().startswith(('file:','c:')) for x in links)
 assert all(x.lower().startswith(('http:','https:','mailto:')) for x in links)
 assert any('water_display_examples.zip' in x for x in links)
 thumbs=[]
 for i,p in enumerate(d):
  pix=p.get_pixmap(matrix=pymupdf.Matrix(1.2,1.2),alpha=False);pix.save(a.out/f'page_{i+1:03}.png')
  im=Image.frombytes('RGB',[pix.width,pix.height],pix.samples);im.thumbnail((280,395))
  thumb=Image.new('RGB',(300,425),'#ddd');thumb.paste(im,((300-im.width)//2,20));ImageDraw.Draw(thumb).text((8,5),str(i+1),fill='black');thumbs.append(thumb)
 for start in range(0,len(thumbs),20):
  sheet=Image.new('RGB',(1200,2125),'white')
  for j,im in enumerate(thumbs[start:start+20]):sheet.paste(im,((j%4)*300,(j//4)*425))
  sheet.save(a.out/f'contact_{start+1:03}.png')
 report={'status':'PASS','pages':len(d),'chapters':ct,'bookmarks':len(toc),'images':len(images),'public_links':len(links),'rendered_pages':len(thumbs),'visual_review':'pending separate human/agent inspection'}
 (a.out/'checks.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(report,ensure_ascii=False))

if __name__=='__main__':main()
