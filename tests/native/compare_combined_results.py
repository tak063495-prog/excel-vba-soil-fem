"""Exact regression for repackaging the existing two DLLs into one binary."""
import argparse
import hashlib
import json
import tempfile
from pathlib import Path
from compare_ui_results import compare, records


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--pair', action='append', nargs=2, type=Path, required=True)
    p.add_argument('--dll', type=Path, required=True)
    p.add_argument('--out', type=Path, required=True)
    args = p.parse_args()
    sha = hashlib.sha256(args.dll.read_bytes()).hexdigest().upper()
    rows = []
    with tempfile.TemporaryDirectory() as tmp:
        for before, after in args.pair:
            old, new = records(before), records(after)
            assert len(old) == len(new)
            for a, b in zip(old, new):
                assert a['NativeDllSha256'] == 'BE1BF49D61D4172CEB2BCD0A7C8850718DBBF4F01EB68E9F42BFD5AFB66482B7'
                assert a['MaterialDllSha256'] == 'F11C8F4BB5A5F27778DE9DBCFA88E9B6BB963806720E1694C26328C32ADD58CD'
                assert b['NativeDllSha256'] == b['MaterialDllSha256'] == sha
                status = b['Run'].split('|')[-1]
                assert status.startswith('NATIVE(soilfem_native_x64.dll;eval_calls=')
                assert ';fallback=0)' in status and ';eval_elements=0;' not in status
                # These are the only differences permitted by the packaging change.
                a['NativeDllSha256'] = a['MaterialDllSha256'] = sha
                a['Run'] = a['Run'].replace('NATIVE(soilfem_mat_x64.dll;', 'NATIVE(soilfem_native_x64.dll;')
            normalized = Path(tmp) / 'baseline.json'
            normalized.write_text(json.dumps(old), encoding='utf8')
            rows.extend(compare(normalized, after))
    result = dict(status='PASS', combined_dll_sha256=sha, comparisons=len(rows),
                  allowed_changes=['two expected DLL hashes to the verified combined hash',
                                   'material status library name'], runs=rows)
    args.out.write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf8')
    print('PASS', len(rows), 'exact combined-DLL comparisons')


if __name__ == '__main__':
    main()
