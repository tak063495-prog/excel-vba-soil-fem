"""Curate accepted terrain tests and package the exact tested release workbook."""
from pathlib import Path
import hashlib,json,shutil,zipfile
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
ROOT=Path(__file__).resolve().parents[1];tmp=ROOT/'tests/tmp/mesh_light';ev=ROOT/'docs/mesh-evidence/20261022';ev.mkdir(parents=True,exist_ok=True)
def load(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def write(p,data):p.write_bytes((json.dumps(data,ensure_ascii=False,indent=2)+'\n').encode())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
report=load(tmp/'release_comparison.json');guard=load(tmp/'guards_checks.json');elastic=load(tmp/'elastic_comparison.json');pack=load(tmp/'package.json')
assert all(x['status']=='PASS' for x in (report,guard,elastic,pack))
assert load(tmp/'ui/ui.json')['status']==load(tmp/'ui2/ui.json')['status']=='PASS'
assert load(tmp/'wet_joint/check.json')['status']==load(tmp/'wet_joint/compare22.json')['status']=='PASS'
assert len(guard['rejected_cases'])==8 and all(x['mesh_unchanged'] for x in guard['rejected_cases'])
book=tmp/'release/2DSoilFEM_ui_20261023.xlsm';dll=tmp/'release/soilfem_native_x64.dll'
assert sha(book).upper()==pack['package_sha256']
assert sha(dll)==sha(ROOT/'workbook/ui_20261022/soilfem_native_x64.dll')
for src,name in [('release_comparison.json','comparison.json'),('guards_checks.json','guards.json'),('elastic_comparison.json','elastic_comparison.json'),('python_experiment.json','python_experiment.json'),('package.json','package.json'),('terrain_direct_elastic.json','elastic_input.json')]:shutil.copyfile(tmp/src,ev/name)
for src,name in [('ui/terrain_mesh.jpg','terrain_mesh.jpg'),('ui/terrain_error.jpg','terrain_error.jpg'),('ui2/terrain_settings.jpg','terrain_settings.jpg'),('wet_joint/check.json','wet_joint_check.json'),('wet_joint/compare22.json','wet_joint_comparison.json')]:shutil.copyfile(tmp/src,ev/name)
write(ev/'ui_checks.json',dict(status='PASS',mesh=load(tmp/'ui/ui.json'),settings=load(tmp/'ui2/ui.json'),visual_review='Actual Excel screenshots reviewed: generated mesh, method2 description without button overlap, reason-specific error dialog.'))
rows=report['cases'][:3];x=range(3);fig,axes=plt.subplots(1,2,figsize=(10,3.6))
for ax,key0,key1,ylabel in [(axes[0],'reference_median_seconds','candidate_median_seconds','Seconds (median of 3)'),(axes[1],'reference_elements',None,'Q8 elements')]:
 old=[r[key0] for r in rows];new=[r[key1] if key1 else r['geometry']['elements'] for r in rows]
 ax.bar([i-.18 for i in x],old,.36,label='Previous');ax.bar([i+.18 for i in x],new,.36,label='Direct terrain')
 ax.set_xticks(list(x),['Slope h=2','Slope h=1','Flat h=1']);ax.set_ylabel(ylabel);ax.legend();ax.grid(axis='y',alpha=.2)
fig.suptitle('Excel mesh creation + preview; direct mode uses its own grid');fig.tight_layout();fig.savefig(ev/'timing.png',dpi=180);plt.close(fig)
with zipfile.ZipFile(ev/'records.zip','w',zipfile.ZIP_DEFLATED) as z:
 for folder in ['baseline3','baseline_extended2','release_benchmark','release_guards','elastic_new4','elastic_ref3','wet_joint']:
  for p in sorted((tmp/folder).rglob('*')):
   if p.is_file():z.write(p,folder+'/'+p.relative_to(tmp/folder).as_posix())
 for name in ['build_final.log','package.json','ui/ui.json','ui2/ui.json','terrain_direct_elastic.json','release_benchmark.log','release_guards.log']:
  z.write(tmp/name,name)
dst=ROOT/'workbook/ui_20261023';shutil.copyfile(book,dst/book.name)
record=dict(build='20261022_WATER_02',mesh_build='20261022_TERRAIN_DIRECT_01',verification_date='2026-10-22',base_main='6930bf04c29932fa37bfc125ee04004dd645cfd4',book_sha256=sha(book),dll_sha256=sha(dll),dll_unchanged=True,compile='PASS',source_verification=pack['source_verification'],modules=pack['modules'],canonical_inputs_unchanged=pack['canonical_inputs_unchanged'],defaults=pack['defaults'],excel_version=pack['excel_version'],excel_build=pack['excel_build'],terrain_geometry='Python independent Q8 area/25-point Jacobian/all coordinates/connectivity PASS',rejections=8,elastic_dofs=1074,elastic_comparison=elastic,water_joint='initial full vectors/metrics exact; independent and resume PASS',solver_law_tolerance_changes=False,evidence='docs/mesh-evidence/20261022')
write(dst/'BUILD_RECORD.json',record)
readme=(dst/'README.txt').read_text(encoding='utf-8-sig')
first,_,rest=readme.partition('\n\n')
new='地形格子の直接生成対応 20261022_TERRAIN_DIRECT_01 / 配布名20261023\n方式2ではデローニを通らず、適用できない場合は理由付きで停止します。基本の計算識別は20261022_WATER_02です。DLLは従来と同一ですが、ブックはメッシュ生成を更新しているため初回23版とは異なります。\n操作・適用条件・実測・制約は同梱MESH_GUIDE.mdを参照してください。\n\n'+rest
(dst/'README.txt').write_bytes(new.encode())
zipname=ROOT/'workbook/SoilFEM_UI_20261023.zip'
with zipfile.ZipFile(zipname,'w',zipfile.ZIP_DEFLATED) as z:
 for p in [dst/book.name,dst/dll.name,dst/'BUILD_RECORD.json',dst/'README.txt']:z.write(p,p.name)
 z.write(ROOT/'docs/GROUNDWATER.md','WATER_GUIDE.md')
 guide=(ROOT/'docs/MESH_TERRAIN_PERF.md').read_text(encoding='utf-8').split('![')[0]
 guide+='\n図・実測・全記録：https://github.com/tak063495-prog/excel-vba-soil-fem/blob/perf/terrain-mesh-light/docs/MESH_TERRAIN_PERF.md\n'
 z.writestr('MESH_GUIDE.md',guide)
with zipfile.ZipFile(zipname) as z:
 assert len(z.namelist())==6 and len([n for n in z.namelist() if n.endswith('.dll')])==1
 assert hashlib.sha256(z.read(book.name)).hexdigest()==sha(book)
write(ev/'release.json',dict(status='PASS',book_sha256=sha(book),dll_sha256=sha(dll),zip_sha256=sha(zipname),zip_entries=6,dll_count=1,records_sha256=sha(ev/'records.zip')))
print(json.dumps(dict(status='PASS',book_sha256=sha(book),zip_sha256=sha(zipname),evidence_files=len(list(ev.iterdir()))),indent=2))
