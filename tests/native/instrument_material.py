"""Create a disposable QPC-instrumented copy of soilfem_mat.c.

The generated source is for single-thread Excel profiling only.  It keeps the
MinGW/x87/no-CRT production source untouched and exports the same profile ABI
as the GMRES disposable profiler.
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


PROFILER = r'''
/* disposable single-thread profiler; do not ship this source or DLL */
static LARGE_INTEGER mat_qp_freq;
static LONGLONG mat_qp_in[10], mat_qp_ex[10], mat_qp_calls[10];
static LONGLONG mat_qp_start[16], mat_qp_child[16];
static int mat_qp_depth, mat_qp_id[16], mat_qp_bad;
static void mat_qp_enter(int id) {
  LARGE_INTEGER q;
  if (mat_qp_depth >= 16) { mat_qp_bad = 1; return; }
  QueryPerformanceCounter(&q); mat_qp_id[mat_qp_depth] = id;
  mat_qp_start[mat_qp_depth] = q.QuadPart; mat_qp_child[mat_qp_depth] = 0;
  mat_qp_depth++; mat_qp_calls[id]++;
}
static void mat_qp_leave(int id) {
  LARGE_INTEGER q; LONGLONG elapsed;
  QueryPerformanceCounter(&q);
  if (mat_qp_depth < 1 || mat_qp_id[mat_qp_depth - 1] != id) { mat_qp_bad = 1; return; }
  mat_qp_depth--; elapsed = q.QuadPart - mat_qp_start[mat_qp_depth];
  mat_qp_in[id] += elapsed; mat_qp_ex[id] += elapsed - mat_qp_child[mat_qp_depth];
  if (mat_qp_depth) mat_qp_child[mat_qp_depth - 1] += elapsed;
}
__declspec(dllexport) void __stdcall sfn_gmres_profile_reset(void) {
  int i; mat_qp_depth = 0; mat_qp_bad = 0; QueryPerformanceFrequency(&mat_qp_freq);
  for (i = 0; i < 10; i++) mat_qp_in[i] = mat_qp_ex[i] = mat_qp_calls[i] = 0;
}
__declspec(dllexport) int32_t __stdcall sfn_gmres_profile_read(double *out, int32_t n) {
  int i;
  if (!out || n < 32 || mat_qp_freq.QuadPart <= 0 || mat_qp_depth || mat_qp_bad) return -1;
  out[0] = (double)mat_qp_freq.QuadPart; out[1] = (double)mat_qp_bad;
  for (i = 0; i < 10; i++) {
    out[2 + 3 * i] = (double)mat_qp_calls[i];
    out[3 + 3 * i] = (double)mat_qp_in[i] / (double)mat_qp_freq.QuadPart;
    out[4 + 3 * i] = (double)mat_qp_ex[i] / (double)mat_qp_freq.QuadPart;
  }
  return 0;
}
'''

NODES = [
    ("sfm_eval", "int32_t"), ("sfm_tangent", "int32_t"),
    ("mp_update", "int"), ("spectral", "int"),
    ("eval_yield", "int"),
    ("c_pow", "double"), ("c_sqrt", "double"),
    ("c_atan", "double"), ("c_sin", "double"), ("c_cos", "double"),
]


def brace_end(source: str, start: int) -> int:
    depth = 1
    quote = False
    i = start
    while i < len(source):
        ch = source[i]
        if ch == '"' and (i == 0 or source[i - 1] != '\\'):
            quote = not quote
        elif not quote:
            if ch == '{': depth += 1
            elif ch == '}':
                depth -= 1
                if depth == 0: return i
        i += 1
    raise ValueError("unterminated function")


def instrument_function(source: str, name: str, kind: str, ident: int) -> str:
    m = re.search(r"\b" + re.escape(name) + r"\s*\([^;{}]*\)\s*\{", source)
    if not m: raise ValueError(f"function not found: {name}")
    start, end = m.end(), brace_end(source, m.end())
    body = source[start:end]
    # A comment such as "spectral return" is not a C return statement.
    body = re.sub(r'/\*.*?\*/|//[^\n]*', lambda m: '\n' * m.group().count('\n'), body, flags=re.S)
    if kind == "void":
        body = re.sub(r"\breturn\s*;", f"{{ mat_qp_leave({ident}); return; }}", body)
        body += f"\n mat_qp_leave({ident});\n"
    else:
        body = re.sub(
            r"\breturn\s+([^;]+);",
            lambda x: "{ " + kind + " mat_qp_result=(" + x.group(1)
            + f"); mat_qp_leave({ident}); return mat_qp_result; }}",
            body,
        )
        body += f"\n mat_qp_leave({ident});\n"
    return source[:start] + f"\n mat_qp_enter({ident});\n" + body + source[end:]


def instrument(source: str) -> tuple[str, list[str]]:
    # Keep the function-pointer names private, and route only their call sites
    # through timed wrappers.  Wrapper calls use the *_raw pointers.
    for name in ("c_pow", "c_sqrt", "c_atan", "c_sin", "c_cos"):
        source = re.sub(r"\b" + name + r"\b", name + "_raw", source)
    source = source.replace("#include <windows.h>", "#include <windows.h>\n" + PROFILER)
    wrappers = r'''
static double mat_qp_c_pow(double a, double b) { mat_qp_enter(5); double r = c_pow_raw(a, b); mat_qp_leave(5); return r; }
static double mat_qp_c_sqrt(double a) { mat_qp_enter(6); double r = c_sqrt_raw(a); mat_qp_leave(6); return r; }
static double mat_qp_c_atan(double a) { mat_qp_enter(7); double r = c_atan_raw(a); mat_qp_leave(7); return r; }
static double mat_qp_c_sin(double a) { mat_qp_enter(8); double r = c_sin_raw(a); mat_qp_leave(8); return r; }
static double mat_qp_c_cos(double a) { mat_qp_enter(9); double r = c_cos_raw(a); mat_qp_leave(9); return r; }
'''
    # Wrapper definitions are inserted after pointer declarations; call sites
    # are rewritten after that insertion so wrappers themselves stay raw.
    marker = "static int g_ready;"
    if marker not in source: raise ValueError("pointer declaration marker missing")
    source = source.replace(marker, marker + wrappers, 1)
    for name in ("c_pow", "c_sqrt", "c_atan", "c_sin", "c_cos"):
        source = re.sub(r"\b" + name + r"_raw\s*\(", "mat_qp_" + name + "(", source)
    # Restore wrapper bodies to raw pointer calls.
    source = source.replace("mat_qp_c_pow(a, b);", "c_pow_raw(a, b);")
    source = source.replace("mat_qp_c_sqrt(a);", "c_sqrt_raw(a);")
    source = source.replace("mat_qp_c_atan(a);", "c_atan_raw(a);")
    source = source.replace("mat_qp_c_sin(a);", "c_sin_raw(a);")
    source = source.replace("mat_qp_c_cos(a);", "c_cos_raw(a);")
    # The c_* entries are timed by the wrappers above; they are function
    # pointers in the production source, not definitions to brace-instrument.
    for ident, (name, kind) in reversed(list(enumerate(NODES[:5]))):
        source = instrument_function(source, name, kind, ident)
    return source, [name for name, _ in NODES]


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default="native/soilfem_mat.c")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    src, out = Path(args.src), Path(args.out)
    if out.exists() or src.resolve() == out.resolve(): raise ValueError("refusing to overwrite source")
    text, names = instrument(src.read_text(encoding="utf-8-sig"))
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(text, encoding="utf-8", newline="\n")
    out.with_suffix(".profile.json").write_text(json.dumps(names, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"out": str(out), "nodes": names}, indent=2))


if __name__ == "__main__":
    main()
