"""Disposable coarse QPC partitions of the material element loop (32-double ABI)."""
import argparse
import json
from pathlib import Path
from instrument_material import PROFILER, instrument_function

NAMES = ['sfm_eval', 'sfm_tangent', 'strain', 'mp_update', 'state_display',
         'stress_B', 'internal_force', 'commit', 'sfm_ilu0', 'sfm_regularized']

def instrument(s):
    s = s.replace('#include <windows.h>', '#include <windows.h>\n' + PROFILER, 1)
    start = s.index('__declspec(dllexport) int32_t __stdcall sfm_eval(')
    end = s.index('\n}', start) + 2
    body = s[start:end]
    changes = [
        ('        for (i = 0; i < 3; i++) { de[i]', '        mat_qp_enter(2);\n        for (i = 0; i < 3; i++) { de[i]'),
        ('        if (!mp_update', '        mat_qp_leave(2);\n        if (!mp_update'),
        ('        for (i = 0; i < 17; i++) W.st', '        mat_qp_enter(4);\n        for (i = 0; i < 17; i++) W.st'),
        ('        for (i = 0; i < 3; i++) {\n          for (j = 0; j < 3;', '        mat_qp_leave(4);\n        mat_qp_enter(5);\n        for (i = 0; i < 3; i++) {\n          for (j = 0; j < 3;'),
        ('      }\n    }\n    /* CalcForce', '        mat_qp_leave(5);\n      }\n    }\n    /* CalcForce'),
        ('      for (j = 0; j < 16; j++) {\n        int kj', '      mat_qp_enter(6);\n      for (j = 0; j < 16; j++) {\n        int kj'),
        ('      if (fp_bad())', '      mat_qp_leave(6);\n      if (fp_bad())'),
        ('      /* ---- commit', '      mat_qp_enter(7);\n      /* ---- commit'),
        ('    out[0] += dCalls;', '    mat_qp_leave(7);\n    out[0] += dCalls;'),
    ]
    for old, new in changes:
        assert body.count(old) == 1, old
        body = body.replace(old, new, 1)
    s = s[:start] + body + s[end:]
    for ident, name, kind in [(9, 'sfm_regularized', 'int32_t'), (8, 'sfm_ilu0', 'int32_t'),
                              (3, 'mp_update', 'int'), (1, 'sfm_tangent', 'int32_t'), (0, 'sfm_eval', 'int32_t')]:
        s = instrument_function(s, name, kind, ident)
    return s

if __name__ == '__main__':
    p = argparse.ArgumentParser()
    p.add_argument('--src', default='native/soilfem_mat.c')
    p.add_argument('--out', required=True)
    a = p.parse_args()
    out = Path(a.out)
    assert not out.exists()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(instrument(Path(a.src).read_text(encoding='utf-8-sig')), encoding='utf-8', newline='\n')
    out.with_suffix('.profile.json').write_text(json.dumps(NAMES, indent=2)+'\n')
