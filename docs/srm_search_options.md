# Optional controls for large static SRM searches

`SRM_SCAN_STEP` controls the upward scan after the Fs=1 seed. The default
remains 0.1; the allowed numeric range is 0.001 to 1. A larger value such as
0.5 reduces the number of complete load-history trials needed to find an
upper bracket. The existing `SRM_TOL` still determines the final interval
width through bisection. The downward search below a failed Fs=1 seed and
fixed-Fs verification retain their existing behavior.

Out-of-range numeric scan values are rejected before any SRM trial. Workbook
setting reconstruction retains its existing behavior for nonnumeric cells:
it replaces them with the registered default, 0.1 for this setting.

`SRM_REUSE_STATIC_FS1=1` optionally reuses the already converged initial
GRAVITY and LOAD/APPLY history as the Fs=1 seed. Its default is 0. Reuse
requires unit strength, an accepted complete prefix, satisfied displacement
constraints and the unchanged engineering force tolerance. Construction,
water-change, origin-reset and coupled histories are excluded; they retain
normal replay. A deferred gravity failure is never accepted by this new path.
The existing gravity-only reuse policy is unchanged.

These controls change the search cost rather than the physical force or
material strength. `tests/check_srm_search_options.py` runs completed loaded
FE searches with both scan widths and both reuse choices, compares final
displacement/stress fields, rejects incomplete/changed/unconverged seed
states, and exercises invalid numeric and defaulted text settings.
