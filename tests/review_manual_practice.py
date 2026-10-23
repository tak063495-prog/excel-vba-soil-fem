"""Render every manual page and verify chapter, figure, and evidence references."""
from pathlib import Path
import hashlib
import json
import re
import shutil
import subprocess

import pymupdf
from PIL import Image, ImageDraw

root = Path(__file__).resolve().parents[1]
source = root / 'docs/manuals/2DSoilFEM_Manual_20261024.md'
pdf = source.with_suffix('.pdf')
md = source.read_text(encoding='utf-8')
digest = hashlib.sha256(pdf.read_bytes()).hexdigest()
out = root / 'tests/tmp/manual_practice' / digest[:12]
out.mkdir(parents=True, exist_ok=True)
doc = pymupdf.open(pdf)
heads = re.findall(r'^## (.+)$', md, re.M)
toc = [row for row in doc.get_toc() if row[0] == 1]
assert [row[1] for row in toc] == heads
assert all(a[2] < b[2] for a, b in zip(toc, toc[1:]))
compact = lambda value: re.sub(r'\s+', '', value)
for _, heading, page in toc:
    assert compact(doc[page-1].get_text()).startswith(compact(heading)), (heading, page)
captions = re.findall(r'(?m)^図(\d+(?:-[\dA-Za-z]+)*)(?=[　 ])', md)
references = set(re.findall(r'図(\d+(?:-[\dA-Za-z]+)*)', md))
assert references <= set(captions)
assert len(captions) == len(set(captions)), 'Duplicate figure captions'
links = {item['uri'] for page in doc for item in page.get_links() if item.get('uri')}
assert all(uri.startswith(('https://', 'http://', 'mailto:')) for uri in links)
assert '20261024_FIXED_STRESS_07' in doc[0].get_text()
assert 'M000' in md and 'Kは載荷セットID' in md
assert all(f'accepted/{name}/fresh_results.xlsm' in md for name in
           ['column_smoke', 'fill_30day', 'wet_birth_ramp', 'saturated_slope_srm'])
evidence = {}
for name, path in {
    'road': 'docs/manuals/construction-evidence/20261021/manual-applied/exercise_summary.json',
    'tunnel': 'docs/manuals/construction-evidence/20261020/tunnel10/independent_checks.json',
    'bearing_flat': 'docs/manuals/construction-evidence/20261021/bearing04/bearing_flat_fine_s050000/bearing_checks.json',
    'bearing_slope': 'docs/manuals/construction-evidence/20261021/bearing04/bearing_slope_fine_s050000/bearing_checks.json',
}.items():
    file = root/path
    data = json.loads(file.read_text(encoding='utf-8'))
    assert data['status'] == 'PASS'
    evidence[name] = {'path': path, 'sha256': hashlib.sha256(file.read_bytes()).hexdigest()}
    if name == 'road':
        assert data['expected_added_weight_kN'] == data['observed_added_reaction_kN'] == 57
        assert data['baseline']['reaction_y_kN'] == 1842 and data['taller']['reaction_y_kN'] == 1899
    if name.startswith('bearing_'):
        assert f"{data['q_kPa']:.3f}" in md
        assert data['actual_settlement_m'] == .05
pages = out/'pages'
pages.mkdir(exist_ok=True)
poppler = shutil.which('pdftoppm')
if not poppler:
    poppler = str(Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies/native/poppler/Library/bin/pdftoppm.exe')
subprocess.run([poppler, '-scale-to', '1000', '-png', str(pdf), str(pages/'page')], check=True)
files = sorted(pages.glob('page-*.png'))
assert len(files) == len(doc)
boards = out/'boards'
boards.mkdir(exist_ok=True)
for start in range(0, len(files), 20):
    board = Image.new('RGB', (1250, 1480), '#d9dfdf')
    draw = ImageDraw.Draw(board)
    for j, file in enumerate(files[start:start+20]):
        im = Image.open(file).convert('RGB')
        im.thumbnail((238, 340))
        x, y = (j % 5)*250, (j//5)*370
        board.paste(im, (x+(250-im.width)//2, y+22))
        draw.text((x+8, y+4), str(start+j+1), fill='black')
    board.save(boards/f'{start+1:03}-{min(start+20,len(files)):03}.jpg', quality=88)
review = [1, 2, 3, 4, 44, 45, 46, 47, 74, 99, 100, 104, 121, 126, 149, 150, 151,
          160, 161, 167, 168, 191, 192, 200, 219, 224, 227, 231, 234, len(doc)]
for page in review:
    doc[page-1].get_pixmap(matrix=pymupdf.Matrix(1.8, 1.8)).save(out/f'detail-{page:03}.png')
report = dict(status='PASS', pdf_sha256=digest, pages=len(doc), chapters=toc,
              examples_contiguous_chapters=[13, 36], unique_figure_captions=len(captions),
              all_pages_rendered=len(files), boards=len(list(boards.glob('*.jpg'))),
              detail_pages=review, evidence=evidence, external_links=len(links),
              local_file_links=0, visual_review='pending')
(out/'checks.json').write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n', encoding='utf-8')
print(out)
print('PASS; rendered', len(files), 'pages; inspect boards and detail images before publication.')
