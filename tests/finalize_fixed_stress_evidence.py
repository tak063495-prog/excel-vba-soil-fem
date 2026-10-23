"""Append the compiled caption correction and its full regression to sealed numerical evidence."""
import argparse
import json
import shutil
from pathlib import Path
from package_fixed_stress_evidence import archive, sha


def read(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))


def finalize(work, out, release):
    summary = read(out/'validation.json')
    assert 'final_package' not in summary['reports'], 'preserve already finalized evidence'
    for item in summary['archives']:
        assert sha(out/item['path']) == item['sha256'], item['path']
    for name in ('final31_done.txt', 'final32_done.txt', 'final_visual31_done.txt', 'final_visual32_done.txt'):
        assert (work/name).read_text(encoding='utf-8-sig').strip() == 'PASS', name
    package = read(work/'package32.json')
    regression_package = read(work/'package31.json')
    regression = read(work/'comparison31.json')
    assert package['status'] == regression['status'] == 'PASS'
    assert sha(release) == package['package_sha256'].lower()
    for result in (work/'acceptance31').glob('*/result.json'):
        assert read(result)['book_sha256'].lower() == regression_package['package_sha256'].lower(), result
    smoke_checks = []
    for name in ('result.json', 'resume.json'):
        a, b = read(work/'acceptance31/column_smoke'/name), read(work/'final32_smoke'/name)
        assert b['status'].startswith('PASS|'), name
        for key in ('disp', 'stress', 'metrics'):
            assert a[key] == b[key], (name, key)
        smoke_checks.append({'record': name, 'fields': ['disp', 'stress', 'metrics'], 'max_difference': 0})
    assert read(work/'final32_smoke/result.json')['book_sha256'].lower() == sha(release)
    assert read(work/'ui32screen/ui.json')['status'] == 'PASS'
    summary['numerical_reference_sha256'] = summary['release_sha256']
    summary['release_sha256'] = sha(release)
    summary['regression_reference_sha256'] = regression_package['package_sha256'].lower()
    summary['release_scope'] = 'package31 corrects 3 water-view captions: all17cases rerun with exact numerical agreement to30, independent column/SRM/water-force checks repeated. Final package32 additionally removes reserved labels and gray styling from implemented time inputs; strict compile/all-source reopen, column fresh/resume exact agreement and actual input UI verified. Numerical algorithm unchanged from30.'
    summary['final_smoke_checks'] = smoke_checks
    summary['archives'].append(archive(out/'final_release_records.zip', [
        ('full17', work/'acceptance31'), ('final32/column_smoke', work/'final32_smoke'),
        ('final32/ui', work/'ui32screen')], books={
        'fresh_results.xlsm',
        'column_smoke', 'fill_30day', 'watercurve', 'wet_birth_ramp', 'column_nonlinear_srm',
        'saturated_slope_srm', 'water_external_linear'}))
    for source, destination, key in (
        ('package32.json', 'final_package.json', 'final_package'),
        ('package31.json', 'regression_package.json', 'regression_package'),
        ('comparison31.json', 'final_comparison.json', 'final_regression'),
        ('monolithic31.json', 'final_monolithic.json', 'final_monolithic'),
        ('monolithic_biot31.json', 'final_biot_anisotropic.json', 'final_biot_anisotropic'),
        ('srm_parent31.json', 'final_srm_parent.json', 'final_srm_parent'),
        ('water_ramp31.json', 'final_water_ramp.json', 'final_water_ramp'),
    ):
        assert read(work/source)['status'] == 'PASS', source
        shutil.copyfile(work/source, out/destination)
        summary['reports'][key] = destination
    for source in ('build31.log', 'acceptance31.log', 'build32.log'):
        shutil.copyfile(work/source, out/source)
    shutil.copyfile(work/'water31screen/water_time_view.jpg', out/'water_time_view.jpg')
    shutil.copyfile(work/'ui32screen/time_coupling.jpg', out/'time_coupling.jpg')
    correction = {'scope': 'water view captions and implemented-input labels/colors only',
                  'unified_diff': (work/'caption32.diff').read_text(encoding='utf-8-sig').splitlines()}
    (out/'display_correction.json').write_text(json.dumps(correction, indent=2, ensure_ascii=False), encoding='utf-8')
    summary['reports']['display_correction'] = 'display_correction.json'
    (out/'validation.json').write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding='utf-8')
    print(json.dumps({'status': 'PASS', 'release_sha256': summary['release_sha256']}, indent=2))


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    for name in ('work', 'out', 'release'):
        ap.add_argument('--'+name, type=Path, required=True)
    args = ap.parse_args()
    finalize(args.work, args.out, args.release)
