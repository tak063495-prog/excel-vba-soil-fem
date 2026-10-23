"""Check manual example entry points without opening or modifying Excel."""
from pathlib import Path
import io
import json
import re
import xml.etree.ElementTree as ET
import zipfile

root = Path(__file__).resolve().parents[1]
md = (root/'docs/manuals/2DSoilFEM_Manual_20261024.md').read_text(encoding='utf-8')
driver = (root/'tests/run_groundwater_fixture.ps1').read_text(encoding='utf-8-sig')
rows = {'profiles': (12, 1, 4), 'materials': (12, 6, 5),
        'boundaries': (216, 1, 6), 'stagePlans': (12, 12, 9),
        'timeOptions': (12, 29, 10), 'drainSets': (216, 22, 7),
        'timeCurves': (420, 22, 3)}
for key, (row, col, width) in rows.items():
    assert re.search(rf'Rows \$water {row} {col} \$data.water.{key} {width}', driver), key
assert 'Rows $water 12 22 $data.water.coupledMaterials' in driver
assert 'if($data.loads[0].Count -ge 11){11}else{10}' in driver
assert 'Kは載荷セットID' in md and '水条件 AC12（AC:AL列）' in md
ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
cases = ['column_smoke', 'mandel_hybrid_refine', 'fill_30day', 'wet_birth_ramp',
         'watercurve', 'water_external_linear', 'saturated_slope_srm']
archive = zipfile.ZipFile(root/'docs/fixed-stress-evidence/20261024/accepted_records.zip')
members = set(archive.namelist())
report = []
for case in cases:
    fixture = json.loads((root/f'tests/fixtures/fixed_stress/{case}.json').read_text(encoding='utf-8'))
    ids = {row[10] for row in fixture.get('loads', []) if len(row) > 10}
    references = {row[8] for row in fixture['water'].get('timeOptions', []) if len(row) > 8 and row[8]}
    assert references <= ids, (case, references-ids)
    entry = {'fixture': case, 'load_set_ids': len(ids), 'time_set_references': len(references)}
    book = f'accepted/{case}/fresh_results.xlsm'
    if case != 'mandel_hybrid_refine':
        assert book in members
        xlsm = zipfile.ZipFile(io.BytesIO(archive.read(book)))
        workbook = ET.fromstring(xlsm.read('xl/workbook.xml'))
        rid = next(sheet.attrib['{'+ns['r']+'}id'] for sheet in workbook.find('m:sheets', ns) if sheet.attrib['name'] == '時間履歴')
        rels = ET.fromstring(xlsm.read('xl/_rels/workbook.xml.rels'))
        target = next(rel.attrib['Target'] for rel in rels if rel.attrib['Id'] == rid)
        target = target.lstrip('/') if target.startswith('/') else 'xl/'+target
        shared = [''.join(item.itertext()) for item in ET.fromstring(xlsm.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in xlsm.namelist() else []
        sheet = ET.fromstring(xlsm.read(target))
        paths = []
        for cell in sheet.findall('.//m:sheetData/m:row/m:c', ns):
            if re.fullmatch(r'K(?:[2-9]|\d{2,})', cell.attrib['r']):
                value = cell.find('m:v', ns)
                text = (shared[int(value.text)] if cell.attrib.get('t') == 's' else value.text) if value is not None else ''.join(cell.itertext())
                if not text:
                    continue
                path = f'accepted/{case}/'+text.replace('\\', '/')
                assert path+'/meta.txt' in members, (case, text)
                for file in ('node.csv', 'gauss.csv', 'water.csv'):
                    assert path+'/'+file in members, (case, path, file)
                paths.append(text)
        assert paths
        entry.update(book=book, saved_frames=len(paths), first_relative_path=paths[0])
    else:
        assert book not in members
        assert len(ids) == len(references) == 176
        entry['entry_point'] = 'JSON driver; archived Mandel ZIP contains CSV/logs'
    report.append(entry)
out = root/'tests/tmp/manual_practice/input_routes.json'
out.parent.mkdir(parents=True, exist_ok=True)
out.write_text(json.dumps({'status': 'PASS', 'cases': report}, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
print('PASS', len(report), 'fixtures; all archived frame paths and named load references resolved')
