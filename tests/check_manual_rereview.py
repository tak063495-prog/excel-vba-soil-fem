"""Check published N01-N05 evidence, repeated road results, and PDF contents."""
from pathlib import Path
import gzip, hashlib, json, zipfile
from pypdf import PdfReader

root = Path(__file__).resolve().parents[1]
evidence = root / 'docs/manuals/construction-evidence/20261021/manual-rereview'
read = lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
for name, count in (('generic_checks.json', 94), ('road_checks.json', 143)):
    report = read(evidence / name)
    assert report['status'] == 'PASS'
    assert len(report['checks']) == count
road = read(evidence / 'road_checks.json')
assert road['analyses'] == 2 and len(road['analysis_seconds']) == 2
tables = []
for name in ('baseline', 'run1', 'run2'):
    tables.append(json.loads(gzip.decompress((evidence / f'road_{name}_tables.json.gz').read_bytes())))
assert tables[0] == tables[1] == tables[2]
assert len(tables[0]) == 9
stages = json.loads(tables[0]['ステージ結果'])
rows = [stages[i:i+20] for i in range(20, len(stages), 20)]
assert len(rows) == 11 and [row[0] for row in rows] == list(range(1, 12))
assert all(row[3] == 'PASS' and row[4] == 1 for row in rows)
assert abs(rows[0][7] - 1440) < 1e-8 and abs(rows[-1][7] - 1842) < 1e-8
pdf = root / 'docs/manuals/2DSoilFEM_Detailed_Manual_20261020.pdf'
reader = PdfReader(pdf)
validation = read(root / 'docs/manuals/manual_validation_20261020.json')
assert len(reader.pages) == validation['pages']
texts = [p.extract_text() for p in reader.pages]
def page(title):return texts[next(p['page'] for p in validation['page_plan'] if p['title']==title)-1]
material=page('2材料の割当：境界番号と行順は別')
assert 'I列の点列' in material and 'H列の点列' not in material
assert '内部fn=Kn×gapは引張正' in texts[12]
assert 'G列の法線応力σn=−fn' in texts[12] and 'I列gapは開口正' in texts[12]
assert '旧UI_04' in texts[9] and '旧入力規則と凡例' in texts[9]
assert '値貼付け' in texts[9] and 'KEY名' in texts[9]
assert '現行UI_05で最初から解析' in page('増分図：実際の操作順')
assert all(f'{i+1} / {len(texts)}' in text for i, text in enumerate(texts))
glossary=next(p['page'] for p in validation['page_plan'] if p['chapter']=='付録C / 用語集')
assert all('用語集' in t for t in texts[glossary-1:])
assert sum(1 for p in reader.pages for a in p.get('/Annots', []) if a.get_object().get('/Subtype') == '/Link') >= 4
assert hashlib.sha256(pdf.read_bytes()).hexdigest() == validation['pdf_sha256']
for archive in (root / 'workbook/SoilFEM_UI_20261020.zip', root / 'docs/manuals/construction-evidence/20261021/manual-review/road_ui05_workbook.zip'):
    with zipfile.ZipFile(archive) as z:
        assert len([n for n in z.namelist() if n.endswith('.dll')]) == 1
        assert len([n for n in z.namelist() if n.endswith('.xlsm')]) == 1
print('PASS N01-N05 preserved; two archived 11-stage road runs match all nine tables;',len(texts),'current pages')
