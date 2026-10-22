"""Read-only check of the manual stage example and remeshing instructions."""
from pathlib import Path
import io,json,zipfile,hashlib
from openpyxl import load_workbook
ROOT=Path(__file__).resolve().parents[1]
md=(ROOT/'docs/manuals/2DSoilFEM_Manual.md').read_text(encoding='utf-8')
fixture=json.loads((ROOT/'tests/fixtures/water_stage_one_way.json').read_text(encoding='utf-8-sig'))
archive=ROOT/'docs/manuals/examples/water_display_examples.zip'
with zipfile.ZipFile(archive) as z:
 name='one_way/water_stage_one_way.xlsm';raw=z.read(name)
 book=load_workbook(io.BytesIO(raw),read_only=True,data_only=True,keep_links=False)
 actual={int(r[0]):list(r[:5]) for r in book['節点データ'].values if isinstance(r[0],(int,float))};book.close()
expected={int(n[0]):n[:5] for n in fixture['nodes']}
assert actual==expected
bottom=[r for r in actual.values() if r[2]==0];interior=[r for r in bottom if 0<r[1]<2]
assert len(bottom)==5 and len(interior)==3
assert all(r[3:5]==[0,1] for r in interior)
assert all(r[3:5]==[1,1] for r in bottom if r[1] in (0,2))
assert all(r[3]==1 for r in actual.values() if r[1] in (0,2))
example=md.split('### 26.1 ')[1].split('### 26.2 ')[0]
assert '底面Y固定・左右X固定（底面の左右隅角はXY固定）' in example
assert '底面X/Y固定' not in example and '底面XY固定' not in example
assert 'C10=ROLLER、C11:C12=ROLLER、C13=NONE、C14=NONE' in example
step=md.split('#### 水理境界・接合・節点載荷があるときの再メッシュ')[1].split('**触らないセル**')[0]
for token in ('A216:F415','F列まで空欄','B3をNONEに変えるだけでは解除になりません','215行目','D列を0','A:J','新しい要素番号・辺番号','水圧プレビュー'):
 assert token in step,token
guards=(ROOT/'src/vba/MeshAdapt.bas').read_text(encoding='utf-8-sig');assert 'Application.CountA(ws.Range("A216:F"' in guards
assert 'If token = "FORCE" Then Err.Raise' in guards
assert '元の配布ブック、VBA、DLL、倍精度・材料則・収束条件は変更していません。' not in md
report=dict(status='PASS',method='Read-only saved workbook/fixture/source guard comparison; Excel not rerun.',archive_sha256=hashlib.sha256(archive.read_bytes()).hexdigest(),book=name,book_sha256=hashlib.sha256(raw).hexdigest(),nodes=len(actual),bottom_nodes=bottom,bottom_interior_X_free_nodes=[r[0] for r in interior],remeshing='Archive old conditions, clear full water boundary input cells, disable JOINT, remove numbered load definitions, regenerate, recreate boundary/load mapping, preview and solve again.')
out=ROOT/'docs/water-display-evidence/20261023/boundary_correction';out.mkdir(parents=True,exist_ok=True);(out/'checks.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');print('PASS: 21 saved node restraints match fixture; interior bottom X is free; remeshing guard instructions match code')
