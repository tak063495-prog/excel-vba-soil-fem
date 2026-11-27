# Strict VBA band arithmetic

`soilfem_band_fast_x64.dll` is an optional 64-bit Excel backend for the existing
packed symmetric band factorization and solve. Put it beside the workbook and
import `FEMBandFast.bas` together with the updated `FEMNative.bas`. Source import
tools must use the updated `src/vba/modules.json`.

Build with Zig 0.14.1: `python build.py --zig path/to/zig.exe`. The supplied PE has
no imported runtime libraries and does not require AVX. Compiler reassociation,
fast math, reciprocal transformations and FMA contraction are disabled.

VBA evaluates compound expressions with extended intermediates on the tested
64-bit Excel. Every original VBA assignment is explicitly rounded back to Double;
using ordinary SSE2 `double` expressions alone produced different low bits. The
loader checks the ABI, exports, full loaded path and arithmetic probes before use.

The operation order, pivot threshold, pivot rejection index and RHS mutation
match the VBA routines. Missing/invalid DLLs, 32-bit Excel and an incompatible
arithmetic probe use the existing backend. A numeric pivot rejection is returned
to the existing solver, never hidden by fallback. `NATIVE_DLL=OFF` disables it.
It changes no material, load, step, Newton or SRM acceptance settings.

Verification on the supplied 1m model and 70 kernel cases is recorded in
`docs/20261127_issue_audit.md`. Exact equality is evidence for the tested inputs,
not a guarantee about untested platforms. The arithmetic probe disables this
backend where its basic arithmetic does not match Excel.
