"""Compare every NATIVE_DLL mode of run_native_compare.ps1 output against OFF.
Displacements are compared bitwise (JSON keeps 17 significant digits);
counts = increments, factorizations, reuses, active plastic points."""
import json, sys
d = json.load(open(sys.argv[1], encoding='utf-8-sig'))
if isinstance(d, dict): d = [d]
by = {}
for r in d: by.setdefault((r['Case'], r['Variant']), {})[r['Mode']] = r
ok = True
for k, m in by.items():
    if 'OFF' not in m: continue
    a = [float(x) for x in m['OFF']['Disp']]; ra = m['OFF']['Run'].split('|')
    for mode, r in m.items():
        if mode == 'OFF': continue
        b = [float(x) for x in r['Disp']]; rb = r['Run'].split('|')
        same = sum(1 for x, y in zip(a, b) if x == y and (x != 0 or str(x) == str(y)))
        md = max(abs(x - y) for x, y in zip(a, b)); sc = max(abs(x) for x in a) or 1
        cnt = ra[4:8] == rb[4:8]; fos = m['OFF'].get('Fos') == r.get('Fos')
        bit = same == len(a) == len(b)
        ok &= bit and cnt and fos and ra[1] == rb[1]
        print(k, 'OFF vs', mode, f'bitwise {same}/{len(a)}', f'maxabs {md:.3e} rel {md/sc:.3e}',
              'counts', cnt, 'fos', fos, 'status', ra[1], rb[1], 't', round(float(ra[0]), 2), round(float(rb[0]), 2))
print('ALL_MATCH' if ok else 'MISMATCH')
