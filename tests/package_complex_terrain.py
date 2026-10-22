"""Publish only accepted complex-terrain evidence and the exact tested XLSM."""
from pathlib import Path
import json,hashlib,shutil,zipfile,re
ROOT=Path(__file__).resolve().parents[1];tmp=ROOT/'tests/tmp/complex_terrain/acceptance9';ev=ROOT/'docs/mesh-evidence/complex-terrain';ev.mkdir(parents=True,exist_ok=True)
def load(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def write(p,d):p.write_bytes((json.dumps(d,ensure_ascii=False,indent=2)+'\n').encode())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
proof=load(tmp/'verification.json');pack=load(tmp/'package.json');sources=load(tmp/'final_sources.json');mesh=load(tmp/'final_mesh/independent.json');memory=load(tmp/'delaunay_memory.json')
assert all(x['status']=='PASS' for x in (proof,pack,sources,mesh,memory))
assert (tmp/'queue.log').read_text(encoding='utf-8-sig').strip().endswith('ALL PASS')
book=tmp/'release/2DSoilFEM_ui_20261023.xlsm';dll=tmp/'release/soilfem_native_x64.dll'
assert sha(book).upper()==pack['package_sha256'] and sha(dll)==sha(ROOT/'workbook/ui_20261022/soilfem_native_x64.dll')
for name,dst in [('verification.json','verification.json'),('final_mesh/independent.json','independent.json'),('package.json','package.json'),('final_sources.json','sources.json'),('final_mesh/summary.json','generation.json'),('delaunay_memory.json','delaunay_memory.json')]:shutil.copyfile(tmp/name,ev/dst)
for case in ['step_wall_fill','box_tunnel','octagon_tunnel','two_holes','gravity_taper_blocks','lwall_taper_blocks','circle32_tunnel','circle64_tunnel']:shutil.copyfile(tmp/f'final_mesh/{case}.png',ev/f'{case}.png')
for case in ['wall','tunnel']:
    assert (tmp/f'ui_{case}_final/ready.txt').exists() and (tmp/f'ui_{case}_final/close.txt').exists()
    shutil.copyfile(tmp/f'ui_{case}_final/excel_{case}.jpg',ev/f'excel_{case}.jpg')
with zipfile.ZipFile(ev/'records.zip','w',zipfile.ZIP_DEFLATED) as z:
    folders=['final_mesh','simple_regression','guard_regression','wet_joint']
    names=['step_wall_fill','box_tunnel','octagon_tunnel','gravity_taper_blocks','lwall_taper_blocks','circle64_tunnel']
    folders += [prefix+name+'_solve' for name in names for prefix in ['final_','reference_']]
    for folder in folders:
        for p in sorted((tmp/folder).rglob('*')):
            if not p.is_file() or p.suffix.lower()=='.dll':continue
            if p.suffix.lower()=='.xlsm' and not (folder.endswith('_solve') and folder.startswith('final_') or folder=='wet_joint'):continue
            z.write(p,folder+'/'+p.relative_to(tmp/folder).as_posix())
    z.write(dll,'soilfem_native_x64.dll')
    reference=tmp/'delaunay_old/delaunay_memory_probe.xlsm'
    assert sha(reference).upper()==load(tmp/'delaunay_old/result.json')['book_sha256']
    z.write(reference,'delaunay_old/reference_memory_book.xlsm')
    for folder in ['delaunay_old','delaunay_new']:
        for p in sorted((tmp/folder).iterdir()):
            if p.is_file() and p.suffix.lower() in ('.csv','.json','.txt'):z.write(p,folder+'/'+p.name)
    z.writestr('README.txt','The six final calculated copies include shape inputs and registered material data. To rerun a book in a subfolder, copy the single DLL from this archive root beside that XLSM. Mesh-only workbook copies are omitted: regenerate them with complex_terrain.json and tests/run_complex_terrain.ps1; all generated coordinates, connectivity, rejection records and logs remain. Reference solver runs retain input hashes, complete outputs and logs; reproduce using the previous PR40 release. The exact pre-memory-optimization workbook is delaunay_old/reference_memory_book.xlsm. This is meshing/elastic self-weight verification, not a construction or full SRM study.\n')
    for name in ['build9.log','package.log','package.json','verification.json','final_sources.json','embedded_sources.json','queue.log','final_mesh.log','simple_regression.log','guard_regression.log','delaunay_memory.json','delaunay_old.log','delaunay_new.log']+[name+'_elastic.json' for name in names]:
        z.write(tmp/name,name)
    z.write(ROOT/'tests/fixtures/complex_terrain.json','complex_terrain.json')
    for case in ['wall','tunnel']:z.write(tmp/f'ui_{case}_final/excel_{case}.jpg',f'ui/excel_{case}.jpg')
dst=ROOT/'workbook/ui_20261023';shutil.copyfile(book,dst/book.name)
record=load(dst/'BUILD_RECORD.json')
historical=dict(record.get('prior_direct_01_validation',{}))
historical.update({key:record.pop(key) for key in ['terrain_geometry','rejections','elastic_dofs','elastic_comparison','water_joint','evidence'] if key in record})
record.update(mesh_build='20261022_TERRAIN_BLOCKS_02',book_sha256=sha(book),dll_sha256=sha(dll),dll_unchanged=True,modules=pack['modules'],compile='PASS',source_verification=sources['normalization'],canonical_inputs_unchanged=pack['canonical_inputs_unchanged'],prior_direct_01_validation=historical,complex_terrain=dict(geometry=mesh,elastic=proof['elastic'],evidence='docs/mesh-evidence/complex-terrain',solver_precision_law_tolerance_changes=False),delaunay_memory=memory)
write(dst/'BUILD_RECORD.json',record)
readme=(dst/'README.txt').read_text(encoding='utf-8-sig')
readme=readme.replace('地形格子の直接生成対応 20261022_TERRAIN_DIRECT_01','地形格子の複雑形状対応 20261022_TERRAIN_BLOCKS_02')
readme=readme.replace('操作・適用条件・実測・制約は同梱MESH_GUIDE.mdを参照してください。','段差・切土の水平/鉛直な複数材料、矩形空洞、凸多角形トンネルと覆工を追加。32/64角形の円形近似も検証。重力式擁壁や縦壁・底盤のテーパーは、土を含め凸四辺形ブロックへ分けた入力で対応（穴なし、完全共有辺、重なり・T字接続不可）。ブロックごとに材料番号を登録します。\n操作・適用条件・実測・制約はMESH_COMPLEX_GUIDE.md、初回直接生成の速度比較はMESH_GUIDE.mdを参照してください。')
if 'MESH_MEMORY_GUIDE.md' not in readme:readme+='\nデローニ分割の境界情報を疎な辺辞書へ変更し、入力点数の二乗配列を廃止。旧版との境界出力一致・Excelメモリー実測はMESH_MEMORY_GUIDE.mdを参照してください。\n'
(dst/'README.txt').write_bytes(readme.encode())
manual=(ROOT/'docs/MESH_TERRAIN_COMPLEX.md').read_text(encoding='utf-8')
guide=re.sub(r'!\[[^\]]*\]\([^)]*\)','（図・実Excel画面はオンライン説明を参照）',manual)
guide='オンライン説明：https://github.com/tak063495-prog/excel-vba-soil-fem/blob/perf/terrain-mesh-light/docs/MESH_TERRAIN_COMPLEX.md\n\n'+guide
# The offline guide links to repository resources, not absent local evidence files.
def external(m):
    target=m.group(2)
    if target.startswith(('http:','https:','#')):return m.group(0)
    resolved=(ROOT/'docs'/target).resolve().relative_to(ROOT).as_posix()
    return '['+m.group(1)+'](https://github.com/tak063495-prog/excel-vba-soil-fem/blob/perf/terrain-mesh-light/'+resolved+')'
guide=re.sub(r'\[([^\]]+)\]\(([^)]+)\)',external,guide)
oldguide=(ROOT/'docs/MESH_TERRAIN_PERF.md').read_text(encoding='utf-8-sig')
oldguide=re.sub(r'!\[[^\]]*\]\([^)]*\)','（図はオンライン説明を参照）',oldguide)
oldguide=re.sub(r'\[([^\]]+)\]\(([^)]+)\)',external,oldguide).encode()
zipfile_path=ROOT/'workbook/SoilFEM_UI_20261023.zip'
with zipfile.ZipFile(zipfile_path,'w',zipfile.ZIP_DEFLATED) as z:
    for name in [book.name,'soilfem_native_x64.dll','README.txt','BUILD_RECORD.json']:z.write(dst/name,name)
    z.write(ROOT/'docs/GROUNDWATER.md','WATER_GUIDE.md');z.writestr('MESH_GUIDE.md',oldguide);z.writestr('MESH_COMPLEX_GUIDE.md',guide.encode());z.writestr('MESH_MEMORY_GUIDE.md',re.sub(r'\[([^\]]+)\]\(([^)]+)\)',external,(ROOT/'docs/MESH_DELAUNAY_MEMORY.md').read_text(encoding='utf-8')).encode())
with zipfile.ZipFile(zipfile_path) as z:
    assert sum(name.lower().endswith('.dll') for name in z.namelist())==1
    assert hashlib.sha256(z.read(book.name)).hexdigest()==sha(book)
write(ev/'release.json',dict(status='PASS',book_sha256=sha(book),zip_sha256=sha(zipfile_path),records_sha256=sha(ev/'records.zip'),dll_sha256=sha(dll),dll_unchanged=True,files_in_zip=8,excel_build=pack['excel_build'],version='20261022_TERRAIN_BLOCKS_02'))
print(json.dumps(load(ev/'release.json'),indent=2))
