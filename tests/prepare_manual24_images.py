from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
import json
import argparse
ap=argparse.ArgumentParser();ap.add_argument('--base',type=Path,default=Path('tests/tmp/manual24'));ap.add_argument('--out',type=Path,default=Path('docs/manuals/images/ui24'));args=ap.parse_args();base=args.base;out=args.out;out.mkdir(parents=True,exist_ok=True)
font=ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf',22)
plan=json.loads((base/'capture_plan.json').read_text(encoding='utf-8'));annotations={}
for q in plan:
 name=q['name'];image=Image.open(base/'raw'/f'{name}.jpg').convert('RGB');kind=name.split('_')[-1];frames=[];right=1370;bottom=900
 if name=='home':bottom=910
 elif kind=='materials':
  n={'ex04':3,'ex08':8,'ex09':3,'ex11':3,'ex10':6,'ex11b':3}.get(name.split('_')[0],1);bottom=min(1050,540+n*32);frames=[((35,501,842,501+n*29),'1'),((843,501,1147,501+n*29),'2')]
 elif kind=='coords':
  n={'basic':6,'ex02':6,'ex04':9,'ex05':4,'ex08':18,'ex08taller':18,'ex09':16,'ex11':8}.get(name.split('_')[0],8);bottom=min(1180,540+n*29);frames=[((35,501,337,501+n*29),'1')]
 elif kind=='regions':
  n={'ex04':3,'ex08':8,'ex09':3,'ex11':2}.get(name.split('_')[0],1);bottom=560+n*29;frames=[((35,501,235,501+n*29),'2'),((236,501,337,501+n*29),'3'),((338,501,439,501+n*29),'4')]
 elif kind=='stages':
  right=1580;bottom=1190 if name.startswith('ex08') else 1050 if name.startswith(('ex09','ex10')) else 900;frames=[((150,501,815,bottom-35),'1')]
  if name=='water_stages':frames.append(((815,501,1048,bottom-35),'2'))
 elif kind in ('analysis','start'):bottom=985;frames=[((337,501,540,703),'1'),((337,704,540,958),'2')]
 elif kind=='mesh':bottom=900;frames=[((337,501,540,551),'1'),((337,552,540,754),'2'),((337,755,540,805),'3'),((337,806,540,857),'4')]
 elif kind=='save':bottom=840;frames=[((337,501,540,805),'1')]
 elif kind=='speed':bottom=1240;frames=[((337,501,540,650),'1')]
 elif kind in ('view','xy','view50','select'):bottom=1060;frames=[((337,501,540,551),'1'),((337,552,540,603),'2')]
 elif kind=='mode':right=1560;bottom=660;frames=[((327,501,427,613),'1')]
 elif kind=='profiles':bottom=720;frames=[((35,501,439,614),'2')]
 elif kind=='boundary':bottom=1050;frames=[((35,501,641,1000),'1')]
 elif name=='water_materials':bottom=740;frames=[((35,501,540,613),'3')]
 elif kind in ('joints','joint'):bottom=740;frames=[((35,501,439,585),'1'),((540,501,741,585),'2')]
 elif kind=='loads':bottom=1100 if name.startswith(('ex12','ex13')) else 750;frames=[((35,501,540,bottom-40),'1'),((540,501,944,bottom-40),'2')]
 elif name.startswith('basic_') and kind in ('nodes','elements','results','diagnostic'):bottom=1140;frames=[]
 elif name in ('basic_preview','basic_plot','basic_plot50'):bottom=1310;right=1370;frames=[]
 if name in ('ex11_materials','ex11b_materials'):
  frames[1]=((843,501,1347,501+n*29),'2')
 if name=='ex03_nodes':frames=[((134,455,337,1120),'1'),((338,455,540,1120),'2')];bottom=1140
 if name in ('ex04_stages','ex05_stages','ex08_stages','ex10_stages','ex12_stages'):frames.append(((150,501,501,bottom-35),'2'))
 if name=='ex08taller_coords':frames=[((35,949,337,1006),'1'),((236,949,337,1006),'2')]
 if name=='ex13_loads':frames=[((35,455,1044,500),'1'),((35,501,134,758),'2'),((135,501,439,758),'3'),((440,501,540,758),'4')]
 if name=='ex12_loads':bottom=1092;frames=[((35,455,1044,500),'1'),((35,501,134,1090),'2'),((135,501,439,1090),'3'),((440,501,540,1090),'4')]
 if name=='basic_start':frames=[((39,329,274,363),'1'),((280,329,520,363),'2')]
 if name=='water_materials':bottom=740;frames=[((35,501,540,613),'3')]
 if name=='ex08_view':frames=[((337,957,540,1010),'1'),((35,901,335,956),'2'),((337,1011,540,1040),'3')];bottom=1060
 annotations[name]=[label for box,label in frames]
 draw=ImageDraw.Draw(image)
 for (x1,y1,x2,y2),label in frames:
  draw.rectangle((x1,y1,x2,y2),outline='#c40000',width=4)
  mx=x1+16;my=max(240,y1-17)
  if name=='ex08_view' and label=='3':mx=x2+22;my=y1+14
  if name in ('ex04_stages','ex05_stages','ex08_stages','ex10_stages','ex12_stages') and label=='1':mx=x2-20
  draw.ellipse((mx-15,my-15,mx+15,my+15),fill='#c40000',outline='white',width=2);draw.text((mx-7,my-13),label,fill='white',font=font)
 image.crop((0,218,right,bottom)).save(out/f'{name}.png',optimize=True)
# Shape views pair coordinate and region sections in the same figure.
for prefix in ('basic','ex02','ex04','ex05','ex08','ex09','ex11'):
 a=Image.open(out/f'{prefix}_coords.png');b=Image.open(out/f'{prefix}_regions.png');target=Image.new('RGB',(max(a.width,b.width),a.height+b.height+20),'white');target.paste(a,(0,0));target.paste(b,(0,a.height+20));target.save(out/f'{prefix}_shape.png',optimize=True)
# Fixed and broken-line water input steps each retain their real source model.
for name,items in [('water_input',['water_mode','water_profiles','water_materials']),('polyline_input',['polyline_mode','polyline_profiles'])]:
 ims=[Image.open(out/f'{n}.png') for n in items];target=Image.new('RGB',(max(x.width for x in ims),sum(x.height for x in ims)+20*(len(ims)-1)),'white');y=0
 for im in ims:target.paste(im,(0,y));y+=im.height+20
 target.save(out/f'{name}.png',optimize=True)
(base/'image_annotations.json').write_text(json.dumps(annotations,ensure_ascii=False,indent=2),encoding='utf-8')
print('Prepared',len(list(out.glob('*.png'))),'guided actual-screen figures')
