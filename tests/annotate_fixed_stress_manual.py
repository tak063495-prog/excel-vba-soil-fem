"""Annotate actual Excel captures; preserve data pixels and trace every crop."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import hashlib,json,zipfile
ROOT=Path(__file__).resolve().parents[1]
CAP=ROOT/'tests/tmp/manual_fs'
OUT=ROOT/'docs/manuals/images/fixed_stress'
# Coordinates refer to original screenshots, never rescaled before marking.
# (crop, [(box, circle center, exact cells/meaning), ...])
SPECS={
'coupled_material':('screens',(20,350,685,525),[([27,432,207,502],[40,421],'A7:B8 material/alpha'),([208,432,480,502],[222,421],'C7:E8 storage/kx/ky'),([571,432,663,502],[585,421],'G7:G8 porosity')]),
'fill_stages':('screens',(20,350,1060,685),[([235,432,340,549],[248,421],'C7:C9 BIRTH material'),([130,539,235,638],[144,421],'B10:B12 SRM/WAIT'),([720,432,825,678],[733,421],'G7:G13 coupling')]),
'fill_time':('screens',(20,350,950,565),[([27,432,208,550],[40,421],'A7:B10 stage/coupling'),([298,432,481,550],[312,421],'D7:E10 duration/initial dt'),([845,432,936,550],[858,421],'J7:J10 time unit')]),
'fill_detail':('screens',(20,350,950,540),[([27,432,207,525],[40,421],'A7:B9 stage/unit'),([208,432,390,525],[222,421],'C7:D9 HOLD/INSTANT'),([390,432,663,525],[405,421],'E7:G9 dt range/save')]),
'drain_sets':('screens3',(20,350,590,550),[([27,432,298,525],[40,421],'A7:C9 set/element/side'),([298,432,480,525],[312,421],'D7:E9 HEAD/profile')]),
'water_series':('screens3',(20,350,590,560),[([27,432,117,550],[40,421],'A7:A10 series ID'),([117,432,208,550],[131,421],'B7:B10 stage time'),([208,432,298,550],[222,421],'C7:C10 water profile')]),
'curve_detail':('screens3',(20,350,950,540),[([571,432,753,525],[585,421],'G7:H9 save interval/series ID')]),
'named_loads':('screens3',(20,350,1040,645),[([208,432,480,626],[222,421],'C7:E13 direction/type/value'),([936,432,1027,626],[949,421],'K7:K13 named load set')]),
'mesh_t6':('screens3',(20,350,1140,845),[([297,788,481,835],[311,777],'D15:E15 T6 option; mesh settings C30')]),
'time_history':('screens2',(20,350,1210,600),[([27,432,390,599],[40,421],'A7:D12 stage/time/dt'),([390,432,754,599],[405,421],'E7:H12 iterations/residuals'),([844,432,936,599],[858,421],'J7:J12 MAINLINE/SRM_BRANCH')]),
'water_balance':('screens3',(20,350,1040,805),[([27,432,298,800],[40,421],'A7:C20 stage/time'),([298,432,754,800],[312,421],'D7:H20 storage/skeleton/flow/closure'),([754,432,936,800],[767,421],'I7:J20 birth/death water')]),
'time_select':('screens2',(20,350,730,705),[([116,473,208,501],[130,512],'B8 saved state ID'),([27,525,571,700],[40,715],'A10:F16 stage/time/branch/Fs'),([298,456,480,489],[312,445],'this time display button')]),
}
def main():
 OUT.mkdir(parents=True,exist_ok=True)
 if not (CAP/'screens/coupled_material.jpg').exists():
  with zipfile.ZipFile(ROOT/'docs/manual-fs-evidence/20261024/actual_excel_captures.zip') as archive:
   for member in archive.infolist():
    assert (CAP/member.filename).resolve().is_relative_to(CAP.resolve())
    archive.extract(member,CAP)
 font=ImageFont.truetype('C:/Windows/Fonts/arialbd.ttf',16)
 manifest={'status':'PASS','method':'Actual Excel screenshots, crop and numbered red frames only. Full original captures and metadata are archived.','figures':{}}
 for name,(folder,box,marks) in SPECS.items():
  src=CAP/folder/(name+'.jpg'); meta=src.with_suffix('.json')
  if not src.exists() or not meta.exists():raise FileNotFoundError(src)
  record=json.loads(meta.read_text(encoding='utf-8-sig'));assert record['status']=='PASS'
  im=Image.open(src).convert('RGB');assert im.width>=box[2] and im.height>=box[3]
  # For selector, place the list marker within crop instead of below it.
  if name=='time_select': marks[1]=(marks[1][0],[585,544],marks[1][2])
  crop=im.crop(box);d=ImageDraw.Draw(crop)
  for i,(r,c,cell) in enumerate(marks,1):
   r=[r[0]-box[0],r[1]-box[1],r[2]-box[0],r[3]-box[1]];cx,cy=c[0]-box[0],c[1]-box[1]
   d.rectangle(r,outline='#d00000',width=3)
   d.ellipse((cx-11,cy-11,cx+11,cy+11),fill='#d00000',outline='white',width=1)
   d.text((cx,cy),str(i),font=font,fill='white',anchor='mm')
  out=OUT/(name+'_guided.png');crop.save(out,optimize=True)
  manifest['figures'][out.name]={'source':str(src.relative_to(ROOT)).replace('\\','/'),'source_sha256':hashlib.sha256(src.read_bytes()).hexdigest(),'capture_metadata':record,'crop_box':box,'marks':[{'number':i,'box':r,'circle':c,'cellrefs':cell} for i,(r,c,cell) in enumerate(marks,1)]}
 for stem in ('partial_birth_view','water_time_view','srm_branch_view'):
  src=ROOT/'docs/fixed-stress-evidence/20261024'/(stem+'.jpg');out=OUT/(stem+'.png')
  Image.open(src).convert('RGB').save(out,optimize=True)
  manifest['figures'][out.name]={'source':src.relative_to(ROOT).as_posix(),'source_sha256':hashlib.sha256(src.read_bytes()).hexdigest(),'conversion':'PNG encoding only; accepted actual result'}
 (OUT/'annotation_manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
 print(json.dumps({'status':'PASS','figures':len(manifest['figures'])}))
if __name__=='__main__':main()
