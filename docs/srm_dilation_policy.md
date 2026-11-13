# Dilation policy before Davis conversion

For a strength-reduced material, SRM first computes
`phi_r = atan(tan(phi_input) / Fs)` and `c_r = c_input / Fs`.
The selected `SRM_PSI_POLICY` then controls dilation:

- `CAP`: use `min(psi_input, phi_r)`.
- `REDUCE`: reduce `tan(psi_input)` by Fs, then cap at `phi_r`.
- `KEEP`: retain `psi_input` without the cap.

These rules apply before any Davis equivalent conversion, for both flow
policies. Previously the cap was skipped for DAVIS, contradicting the setting
description and allowing dilation greater than the reduced friction angle.
For input phi=psi=30 degrees and Fs=1.2, CAP should produce phi=psi=25.6933809045
degrees; the old Davis path instead produced 25.6126285856 degrees and a
slightly smaller cohesion. This correction implements the selected policy;
it does not calibrate material strength to a reference safety factor.

Materials excluded from strength reduction do not acquire Fs-dependent
changes; their selected flow-policy conversion still applies. The existing
Davis conversion of nonassociated continuum materials and the separate JOINT
treatment remain in effect. KEEP retains its explicit uncapped
meaning.

`tests/check_srm_psi_policy.py` compares the production transformation against
independent trigonometric formulas for 96 combinations of flow policy,
dilation policy, input dilation, strength factor and reduction exclusion. Its
optional embedded-source mode reproduces the old CAP defect. Full analysis
and checkpoint regressions are separate from this material-parameter test.
