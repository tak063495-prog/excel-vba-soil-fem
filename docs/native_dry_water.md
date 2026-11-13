# Optional dry-element evaluation within a static groundwater model

`NATIVE_DRY_WATER=1` permits the existing material DLL to evaluate runs of
strictly dry elements in a HYDROSTATIC model. The default is 0. This option
requires a compatible material DLL and its normal AUTO/MAT selection.

Every trial evaluation rechecks all four integration points. Both current
and committed pore pressures must be exactly zero; partially wet elements,
joints, time-dependent water and coupled modes retain VBA evaluation. The
native run ends before the first ineligible element. Material parameters,
water ramp, density, force assembly order and convergence criteria are
unchanged. The wet-mode constitutive reevaluation at zero strain is retained:
the dry kernel's copy-only zero-increment shortcut is disabled in this mode.

No DLL ABI change is required. Missing or declining kernels retain the
existing element-by-element VBA fallback. `dry_water_elements` in the material
backend diagnostic counts elements actually evaluated through this path.

`tests/check_native_dry_water.py` compares Q8/T6 wetting/drying and a plastic
Q8 model with the original AUTO and VBA OFF paths, including checkpoint
resumption. `tests/check_native_dry_water_state.py` compares force, stress
state and material tangent at identical zero/nonzero trial increments and
checks that committed displacement is unchanged. Its timings measure material
evaluation only; they are not a whole-analysis speedup claim.
