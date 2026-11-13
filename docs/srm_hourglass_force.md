# Material-only SRM equilibrium for every flow policy

In SRM, the equilibrium force uses the constitutive material stresses. The
elastic Q8 hourglass penalty must not supply unlimited artificial bearing
strength after the material reaches its yield surface. The previous code
excluded that force for INCONSISTENT but retained it for DAVIS. Even when the
actual material had phi=psi and Davis made no parameter conversion, changing
only the policy label changed the internal force and its tangent.

Both SRM policies now use the same material-only equilibrium force. Native
and VBA evaluation and full-Newton tangents use the same rule. Hourglass
stiffness may still regularize an elastic fallback correction, whose acceptance
is checked against the unchanged material force. Ordinary non-SRM static
stabilization is unchanged. Cohesion, friction, dilation and the force-residual
convergence tolerances are not adjusted by this correction.

The model signature distinguishes the corrected DAVIS+SRM operator when the
configured hourglass factor is positive. Older checkpoints from that operator
are incompatible and must be recomputed; they must not be silently reused.
Histories with the already material-only INCONSISTENT SRM operator, ordinary
static histories and zero-hourglass histories retain their existing signature.

`tests/check_srm_hourglass_policy.py` compares the same associated material
state under both labels, with a nonaffine displacement that excites a nonzero
hourglass mode. It checks force, tangent, stress, unchanged committed state and
ordinary-static control, using native AUTO and VBA OFF. Its optional old-code
mode reproduces the label-dependent force and tangent. Full-model convergence
and compatible/incompatible checkpoint tests are separate checks.
