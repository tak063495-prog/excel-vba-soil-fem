"""Check that the published workbook, source, hashes and gravity evidence agree."""
from pathlib import Path
import hashlib,json,zipfile
from oletools.olevba import VBA_Parser
from check_embedded_sources import normalize

root=Path(__file__).resolve().parents[1]
archive=root/'delivery/2DSoilFEM_1m_gravity_fix_c6000_20261128.zip'
with zipfile.ZipFile(archive) as z:
    assert z.testzip() is None
    books=[n for n in z.namelist() if n.endswith('.xlsm')]
    assert len(books)==1
    base=books[0].rsplit('/',1)[0]
    book=z.read(books[0]);dll=z.read(base+'/soilfem_band_fast_x64.dll')
    record=json.loads(z.read(base+'/validation.json'))
    assert hashlib.sha256(book).hexdigest()==record['workbook_sha256']
    assert hashlib.sha256(dll).hexdigest()==record['dll_sha256']
    parser=VBA_Parser('verified.xlsm',data=book)
    modules={name:code for _,_,name,code in parser.extract_macros()};parser.close()
    source_files=list((root/'src/vba').glob('*.bas'))+[(root/'src/vba/ThisWorkbook.cls')]
    for p in source_files:
        assert p.name in modules,p.name
        assert normalize(p.read_text(encoding='utf-8-sig'))==normalize(modules[p.name]),p.name
    catalog=json.loads((root/'src/vba/modules.json').read_text(encoding='utf-8-sig'))
    assert {p.stem for p in source_files}=={m['name'] for m in catalog}
    assert not any('EXPECTED_GRAVITY_VERIFICATION_BOUNDARY' in c or 'EXPECTED_FIRST_INCREMENT_CONVERGED' in c for c in modules.values())
    gravity=record['gravity'];last=gravity['last_increment']
    assert gravity['passed'] and gravity['load_factor']==1 and gravity['srm_trials']==0
    assert gravity['manifest'][-1][2:5]==['1','GRAVITY','SUCCESS']
    assert gravity['accepted_increments']==8 and gravity['max_iterations_per_increment']<=50
    assert float(last['relative_residual'])<=1e-5
    assert float(last['correction_ratio'])<=1e-4 or float(last['relative_residual'])<=1e-8
    assert len(record['issue_regressions'])==5 and all(r['passed'] for r in record['issue_regressions'])
    assert record['correction_operator_regression']['result'][3]==0
    assert record['audit']['input_cells_identical']==710691
    assert record['final_global_convergence_unchanged'] and not record['production_instrumentation_included']
    if 'issue_repairs' in record:
        repairs=record['issue_repairs']
        assert repairs['73']['passed'] and repairs['75']['passed']
        assert repairs['73']['source_sha256']==repairs['75']['source_sha256']==record['workbook_sha256']
        for case in repairs['73']['cases']:
            assert case['passed'] and case['result'][4:6]==[4,2]
            assert case['metrics'][0]>0.1 and max(case['metrics'][1:6])==0
            assert case['explicit_conflict_control']['rejected']
        for name in ['signature_mismatch_before_write','validation_failure_after_write','normal_prepared_restore']:
            assert repairs['75'][name]['passed']
        assert all(r['status']=='PASS' for r in record['checkpoint_regressions'].values())
        assert record['solver_source_unchanged_from_previous_build']
    evidence=json.loads((root/'docs/verification/srm_gravity_1m_20261128.json').read_text(encoding='utf8'))
    assert evidence==record
    exports=[n for n in z.namelist() if n.startswith(base+'/修正ソース/') and n.endswith('.bas')]
    for export in exports:
        n=export.rsplit('/',1)[-1]
        exported=z.read(base+'/修正ソース/'+n)
        assert b'\r\r\n' not in exported,n
        assert normalize(exported.decode('utf-8-sig'))==normalize(modules[n]),n
print(f'PASS: {len(source_files)} VBA source modules, exact workbook/DLL hashes, and original gravity SUCCESS evidence')
