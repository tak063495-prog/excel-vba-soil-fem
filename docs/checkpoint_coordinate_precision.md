# Coordinate precision in water checkpoints

Checkpoint kinematics now use scientific notation with sixteen fractional
digits for coordinates and displacements. The previous general writer used
ten decimal places for coordinates and twelve for displacement. Reconstructing
geometry from those rounded coordinates could change shallow hydrostatic
pressure enough to fail the existing pressure-integrity check.

For legacy fixed-ten-coordinate files only, static HYDROSTATIC soil pressure
comparison includes a bound propagated from coordinate rounding:

```
delta = 5.1e-11 * sum(abs(N_i))
extra_pressure_bound = gamma_water * delta * (1 + maximum_abs_profile_slope)
```

Here `N_i` are the actual Q8 or T6 shape values at the checked integration
point. The original relative/absolute pressure check remains the base bound.
Scientific-coordinate files retain that original check. The saved pressures
are preserved on acceptance. This is a serialization-compatibility allowance,
not a change to mechanical equilibrium, constitutive or displacement convergence.

The allowance does not apply to STEADY or fixed-stress reconstruction. At a
partial-profile DRY endpoint it is zero if the coordinate uncertainty could
reach the pressure discontinuity; the continuous-gradient estimate must not
conceal a wet/dry jump. Input signatures, missing/duplicate rows, finite-value
checks and substantial pressure discrepancies still cause rejection.

`tests/check_checkpoint_coordinate_precision.py` uses a translated Q8/T6
model with shallow pressure. It checks fresh analysis and resume, scientific
and legacy frame loading, rejection of pressure corruption above the derived
bound, rejection of falsely precise rounded coordinates, and the zero
allowance at a DRY endpoint.
