# 1m SRM speed and open issue audit, 2026-11-27

Base: main `6a9ad0c5bec98a3012d71a18e91f33f6cfab9c9e`.
Reference for the narrowly ported fixes: PR #68 head
`0ee780a33db8bcc3c5e46152b33de3c8b0a27569`.

The reported 88-minute interval was **startup through the first SRM Newton
update**, not a completed factor-of-safety search. The same stop point took
332.285 seconds instead of 5293.251 seconds, a 15.93x speedup on this computer.
The supplied model retains 24,568 nodes, 12,117 T6 elements and 1m spacing.
All 710,691 populated input/settings/cache cells and formulas were compared with
the earlier corrected 1m workbook and are identical. Geometry, strengths,
loads, stage plan and numerical acceptance settings were not changed.

## Changes

- An optional strict band factor/solve DLL replaces the dominant interpreted
  arithmetic, preserving operation order and VBA assignment rounding.
- The rejected line-search state is reconstructed at the restored displacement
  before residual and plastic-set classification. This SRM bug fix was already
  in the earlier corrected workbook used as the timing reference.
- #62: hydraulic rollback includes accepted boundary flow.
- #63: validate configured dt before applying saved adaptive dt. Legacy accepted
  subminimum dt can resume; new cutbacks stop at the configured minimum.
  This corrects the time-coupled bug; it is not used by the dry 1m benchmark.
- #64: comparison checkers reject non-finite JSON and CSV values before reduction.
- #65: transfer surviving pressure and reference pressure by the old local
  element map when a CONTINUOUS joint changes pressure representatives.
- #66: write hydraulic provenance and reject a coupled checkpoint missing
  `fixed_stress.csv`. Legacy inference uses only the executed prefix and skips
  disabled stage rows, including disabled rows before the saved stage.
- #67: save/restore Fs pass/fail/mid/width, bracket flag, interpretation and trial
  count; restore the SRM result row. A legacy dry SRM save explicitly reports
  `LEGACY_INTERVAL_UNAVAILABLE` rather than inventing an interval.

The Newton solver, residual tolerances, material model, SRM interval tolerance
and search strategy are unchanged. PR #68's additional solver, hourglass,
groundwater-weight and support features are not imported by this change.
The build stamp changes intentionally; do not bypass normal checkpoint
signature validation to load checkpoints from a different build.

## Open issues already addressed elsewhere

| Issues | Observed status and scope |
| --- | --- |
| #1 | Existing load-holding fixes and explicit restrictions. The dormant global mixed-u/p and CONSOL paths are rejected; JOINT remains a historyless cap with psi=0. This is not a new history-dependent joint model. |
| #26, #38 | Existing zero-local-increment guard and rerun plastic-result reset remain present. |
| #27 | Existing preflight blocks destructive formula relocation. |
| #43 | Existing complete groundwater snapshot restoration remains present. |
| #44, #45 | Existing internal-HEAD restriction and complete boundary profile checks remain present. |
| #46 | Existing result-history ordering is used for RESET_STRESS display. |
| #69 | PR #68 has the header fix. The supplied private phase-A workbook contains the diagnostic block and receives the same `7 + listRow` correction. Main's older UI module does not contain that block. |
| #70 | PR #68 excludes JOINT elements from the support-connection scan and accepts only initially active continuum nodes, avoiding JOINT's optional zero fields. The FaceSupport feature is absent from this main base and the supplied workbook. |
| #71 | PR #68 resolves `*` to the effective profile before dictionary access and validates the cache/profile. That compatibility routine is absent from this main base and supplied workbook. |

The #69-71 fixes in PR #68 were inspected at the reference SHA. #70-71 were not
independently executed here. Their presence in a draft PR does not mean that
they have been merged to main. No issue was closed and no existing PR merged.

## Validation performed

- Excel compilation and settings/native arithmetic-probe validation with this
  branch's actual changed source modules.
- 70 direct comparisons against the original private VBA band kernels: all
  factor entries, mutated RHS, solutions, success/failure and pivot indices
  equal bit for bit. Includes negative, zero and exact-tolerance pivots.
- Actual 1m matrix: 49,136 DOFs, bandwidth 531; all 26,140,352 factor entries,
  RHS and solution equal bit for bit. The reference factorization took 182s.
- Same 1m stage sequence through the first SRM Newton update: 29 factorizations,
  29 solves, two SRM trials; every one of nine rejected-attempt records exactly
  matches the earlier workbook after removing its build identifier.
- Native, missing-DLL, invalid-DLL and OFF small T6 gravity/SRM controls have
  identical displacement/force/RHS state bits. Missing DLL uses the old path.
- Real Excel small fixtures: dry fixed-Fs, minimum-step, joint pressure split,
  disabled-coupled-prefix and the fixed-stress column smoke case all pass.
- Checkpoint tests pass for #62, #63, #65, #66 and #67. They cover normal/legacy
  metadata, a missing hydraulic file, a disabled preceding coupling row,
  fixed-Fs state, every interval field, malformed/non-finite intervals and
  the displayed SRM coefficient/trial count.
- Explicit controls are distinguished from physical solves: #62 seeds flow
  only if the natural accepted flow is zero; #65 seeds accepted 100kPa before
  the real topology split; #63 forces one trial rejection and separately
  changes saved adaptive dt to exercise legacy compatibility on a constant,
  fully constrained field. The bracket-field test seeds metadata; it does not
  claim this elastic fixed-Fs fixture physically found a failure bracket.
- Four Python non-finite regression test methods pass, covering NaN and both
  infinities in input JSON/CSV and finite/mismatching controls.
- Private phase-A #69: actual Excel conversion with zero, one and three unused
  points retains the header, every unused ID, and four used coordinates.

The benchmark's deliberate stop exception is injected into a disposable test
copy. It is absent from the delivered workbook. No completed full-model SRM
result is asserted or included. Timing is computer-dependent.

Reproduce the focused tests with an Excel workbook built from this branch:

```powershell
python tests/generate_issue62_67_fixtures.py
python -m unittest discover -s tests -p test_nonfinite_regression.py -v
python tests/check_strict_band_excel.py --book path/to/book.xlsm --out evidence/new_kernel
# Optional, slow original-VBA comparison of the actual model matrix:
python tests/check_strict_band_excel.py --book path/to/1m.xlsm --out evidence/new_actual --actual
```

Use the existing `run_groundwater_fixture.ps1 -Resume` harness for the four
generated fixtures plus `fixed_stress/column_smoke.json`, each in its own run
directory named after the fixture. With those `fresh_results.xlsm` files:

```powershell
python tests/check_checkpoint_repairs_excel.py --runs-root evidence/targeted --out evidence/new_checkpoint_checks
```

For the supplied phase-A prepared-mesh workbook, use disposable test copies,
select mesh mode 1 for these small explicit-node fixtures, clear the entire
old node/element tables and unprotect input sheets before applying the fixture.
The production 1m workbook remains in its prepared-mesh mode 3.
