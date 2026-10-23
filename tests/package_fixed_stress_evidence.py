"""Package completed, independently checked Excel runs without transient workbooks."""
import argparse
import hashlib
import json
import shutil
import zipfile
from pathlib import Path


def sha(path):
    return hashlib.file_digest(path.open('rb'), 'sha256').hexdigest()


REPORTS = {
    'package': 'package30.json', 'matrix': 'matrix30/check.json',
    'monolithic': 'monolithic30.json', 'biot_anisotropic': 'monolithic_biot30.json',
    'srm_parent': 'srm_parent30.json', 'rollback': 'rollback30_check.json',
    'water_ramp': 'water_ramp30.json', 'water_profiles': 'water_profiles30c/water_profiles.json',
    'history': 'history30cold/comparison.json', 'water_history': 'water_history30cold/comparison.json',
    'guards': 'guards30/guards.json', 'ui': 'ui30/ui.json', 'mandel': 'mandel30_comparison.json',
}


def archive(target, groups, books=()):
    records = []
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for prefix, root in groups:
            for path in sorted(root.rglob('*')):
                if not path.is_file():
                    continue
                relative = path.relative_to(root)
                if path.suffix.lower() == '.xlsm':
                    if path.name != 'fresh_results.xlsm' or relative.parts[0] not in books:
                        continue
                elif path.suffix.lower() not in {'.json', '.csv', '.log', '.txt', '.svg'}:
                    continue
                name = (Path(prefix)/relative).as_posix()
                z.write(path, name)
                records.append({'path': name, 'bytes': path.stat().st_size, 'sha256': sha(path)})
        z.writestr('archive_manifest.json', json.dumps(records, indent=2, ensure_ascii=False))
    if target.stat().st_size >= 100_000_000:
        raise ValueError(f'archive exceeds GitHub ordinary-file limit: {target}')
    return {'path': target.name, 'bytes': target.stat().st_size, 'sha256': sha(target), 'entries': len(records)}


def main(work, out, release):
    if out.exists():
        raise ValueError('preserve previous evidence directory')
    for marker in ('final29_done.txt', 'after_post30_done.txt', 'final_visual30_done.txt'):
        if not (work/marker).exists():
            raise ValueError(f'queue incomplete: {marker}')
    reports = {}
    for name, relative in REPORTS.items():
        reports[name] = json.loads((work/relative).read_text(encoding='utf-8-sig'))
        if reports[name].get('status') != 'PASS':
            raise ValueError(f'unaccepted report {relative}')
    if reports['package']['package_sha256'].lower() != sha(release):
        raise ValueError('release workbook differs from tested package')
    cases = []
    for folder in sorted((work/'acceptance30').iterdir()):
        if not folder.is_dir():
            continue
        if (folder/'result.json').exists():
            result = json.loads((folder/'result.json').read_text(encoding='utf-8-sig'))
            resume = json.loads((folder/'resume.json').read_text(encoding='utf-8-sig'))
            if not all(x['status'].startswith('PASS|') for x in (result, resume)):
                raise ValueError(f'unaccepted case {folder.name}')
            cases.append({'case': folder.name, 'status': 'PASS', 'seconds': result['seconds'],
                          'resume_seconds': resume['seconds'], 'metrics': result['metrics'],
                          'fixture_sha256': result['fixture_sha256']})
        elif (folder/'expected_failure.json').exists():
            result = json.loads((folder/'expected_failure.json').read_text(encoding='utf-8-sig'))
            if result['status'].startswith('PASS|') or result['expected'] not in result['status']:
                raise ValueError(f'unexpected rejection {folder.name}')
            cases.append({'case': folder.name, 'status': 'EXPECTED_REJECTION', 'reason': result['status']})
        else:
            raise ValueError(f'no result in {folder}')
    if len(cases) != 18:
        raise ValueError(f'expected 17 suite inputs and Mandel, found {len(cases)}')
    out.mkdir(parents=True)
    for name, relative in REPORTS.items():
        shutil.copyfile(work/relative, out/(name+'.json'))
    for source in ('legacy29/comparison.json', 'legacy29/independent.json', 'legacy29/separated_independent.json'):
        shutil.copyfile(work/source, out/('legacy_'+Path(source).name))
    for source in ('mandel23_final.json', 'mandel25_final.json', 'mandel30_final.json'):
        shutil.copyfile(work/source, out/source)
    archives = [
        archive(out/'accepted_records.zip', [('accepted', work/'acceptance30')], books={
            'column_smoke', 'fill_30day', 'watercurve', 'wet_birth_ramp', 'column_nonlinear_srm',
            'saturated_slope_srm', 'water_external_linear'}),
        archive(out/'legacy_records.zip', [('legacy', work/'legacy29')]),
        archive(out/'mandel_refinement_records.zip', [
            ('coarse', work/'fixed_stress_mandel23'), ('fine', work/'mandel_refine25')]),
        archive(out/'continuation_records.zip', [
            ('birth_cold', work/'history30cold'), ('water_cold', work/'water_history30cold'),
            ('visible_plot_then_continue', work/'history30visible'),
            ('srm_view', work/'branch30screen'), ('ui', work/'ui30screen')]),
    ]
    for folder, name in (('ui30screen', 'time_coupling.jpg'),
                         ('history30visible', 'partial_birth_view.jpg'),
                         ('branch30screen', 'srm_branch_view.jpg')):
        shutil.copyfile(work/folder/name, out/name)
    shutil.copyfile(work/'mandel30_final.png', out/'mandel_pressure.png')
    for source in ('build30.log', 'package30.log', 'acceptance30.log', 'final29.log', 'cold_queue30.log'):
        shutil.copyfile(work/source, out/source)
    result = {'status': 'PASS', 'calculation_stamp': '20261024_FIXED_STRESS_05',
              'release_sha256': sha(release), 'cases': cases, 'archives': archives,
              'regression_scope': 'full legacy suite on build04; build05 changes only fixed-stress recovery iteration accounting, with the complete coupled suite rerun',
              'reports': {key: key+'.json' for key in REPORTS}}
    (out/'validation.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'status': 'PASS', 'cases': len(cases), 'archives': archives}, indent=2))


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    for name in ('work', 'out', 'release'):
        ap.add_argument('--'+name, type=Path, required=True)
    args = ap.parse_args()
    main(args.work, args.out, args.release)
