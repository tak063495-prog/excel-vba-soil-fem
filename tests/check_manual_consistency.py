"""Check the three reviewed documentation inconsistencies without running Excel."""
import argparse,hashlib,json,re
from pathlib import Path
from openpyxl import load_workbook

ROOT=Path(__file__).resolve().parents[1]
def load(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def main():
 ap=argparse.ArgumentParser();ap.add_argument('--report',type=Path,required=True);a=ap.parse_args()
 s=(ROOT/'docs/manuals/2DSoilFEM_Manual.md').read_text(encoding='utf-8')
 quick=next(x for x in s.splitlines() if x.startswith('|WATER|'))
 full=s.split('### 8.5 ')[1].split('### 8.6 ')[0]
 for text in (quick,full):
  assert text.index('N列')<text.index('C列')<text.index('B5')
  assert '工程ID' in text and ('同じID' in text or '同一ID' in text)
  assert '停止' in text
 detail=s.split('### 8.5 ')[1].split('### 8.6 ')[0]
 assert 'N列が空欄なら「ステージ」C列、さらにC列も空欄なら「水条件」B5' in detail
 assert 'N列とC列の両方を指定する場合は同じID' in detail
 assert '33候補比較を省き' in s and '保存メッシュを保持' in s
 assert '上表の72要素・261節点・Fsを新規生成の期待値にはしません' in s
 assert '330要素という値や旧Fsを再現の期待値にしません' in s
 assert 'C29=2を選んだ場合、その直接生成の内部で' in (ROOT/'docs/MESH_TERRAIN_COMPLEX.md').read_text(encoding='utf-8')
 screens=load(ROOT/'docs/manuals/water_screen_sources.json')
 expected={'図4-5-A':('cutoff','STEADY'),'図4-5-B':('one_way','STEADY'),'図4-5-C':('fixed13','HYDROSTATIC'),'図4-5-D':('cutoff','STEADY')}
 assert len(screens)==4
 assert '解析例14 折れ線' in s and '解析例18 水圧' in s
 assert s.index('## 22. 解析例13')<s.index('## 23. 解析例14')<s.index('## 27. 解析例18')
 for stem in ('input_mode','input','boundary','stage_plan','stages','polyline_input'):
  assert 'water_display/'+stem+'_guided.png' in s
 for row in screens:
  assert (row['model'],row['mode'])==expected[row['figure']]
  fixture=ROOT/row['fixture'];image=ROOT/'docs/manuals/images/water_display'/row['image']
  assert hashlib.sha256(fixture.read_bytes()).hexdigest()==row['fixture_sha256']
  assert hashlib.sha256(image.read_bytes()).hexdigest()==row['image_sha256']
  if row.get('original_image'):
   assert hashlib.sha256((image.parent/row['original_image']).read_bytes()).hexdigest()==row['original_image_sha256']
  assert load(fixture)['water']['mode']==row['mode']
  assert load(ROOT/row['ui_action'])['source_book_sha256']==row['calculated_book_sha256']
  caption=next(x for x in s.splitlines() if x.startswith(row['figure']+'　'))
  assert row['model'] in caption
 counts=[]
 for name,nn,ne in [('ex01_material_20261021.xlsm',261,72),('ex02_shape_mesh1m_20261022.xlsm',1065,330)]:
  w=load_workbook(ROOT/'docs/manuals/examples'/name,read_only=True,data_only=True,keep_links=False)
  actual=[sum(isinstance(r[0],(int,float)) for r in w[sheet].values) for sheet in ('節点データ','要素データ')];w.close()
  assert actual==[nn,ne],(name,actual)
  counts.append(dict(book=name,nodes=nn,elements=ne))
 report=dict(status='PASS',water_priority='N (matching stage ID), C, B5; N/C conflict rejected',screen_sources=screens,saved_mesh_counts=counts,scope='Documentation/static input inspection only; no Excel recalculation or XLSM/VBA/DLL modification.')
 a.report.parent.mkdir(parents=True,exist_ok=True);a.report.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');print('PASS: WATER priority, current/legacy mesh routes and four screen sources')
if __name__=='__main__':main()
