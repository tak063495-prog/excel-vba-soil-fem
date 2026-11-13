# Static partial water profiles and total body weight

`材料データ` column E remains the moist/total unit weight. Optional column N
specifies saturated **total** unit weight in the same units. Blank N preserves
the existing constant-weight behavior. An entered saturated value must be finite,
positive and at least the moist value.

For static `HYDROSTATIC`, Q8 and T6 body-weight assembly selects moist or saturated
weight at each physical integration point using that material's profile. Water
pressure remains a separate effective-stress/external-water contribution. Do not
enter submerged unit weight as saturated total weight: that would subtract water
effects twice. A waterline crossing an element is sampled by the existing
quadrature; refine the mesh when sensitivity of the integrated wet weight matters.

Variable saturated weight is currently restricted to static prescribed water
with one-way coupling. A nonblank N different from E with STEADY or FIXED_STRESS
is rejected. With water NONE, moist weight is used.

`GW_PROFILE_OUTSIDE` is available on the analysis settings page:

- `ERROR` (default): every queried point/loaded edge must be covered by its profile.
- `DRY`: static HYDROSTATIC only; pressure and external water load are zero outside
  the supplied X interval. There is no lateral head extrapolation. Face loads are
  clipped at profile endpoints, including waterline intersections inside an edge.

DRY does not choose a wall's exposure condition. Specify exposed structural faces
as DRY or EXPOSED according to the actual external water. Transient profiles and
coupled-stage plans do not support this partial-profile option.

The material and water signatures include the optional density/profile settings;
recompute the water preview after changing inputs. The density remains editable
in column N of the detailed material table.

# Recovery after a rejected Newton direction

Associated-flow line-search failure can now use the same column-scaled
Levenberg–Marquardt residual recovery already used by the nonassociated path.
The CSR tangent is assembled lazily when the ordinary band solver has not needed
it. The constitutive law, equilibrium residual and convergence tolerances are
unchanged. Recovery must decrease the actual nonlinear free-force residual.

The optional `sfn_regularized` export factors the normal equations in the selected
SSE2/AVX2 precision. STRICT and older DLLs retain the original x87 core. Packed
workspace layout avoids long-stride Cholesky reads. With physical water active,
only this assembled algebra may use the optional recovery kernel; material
updates continue through the existing physical-water implementation.

Validation includes:

- byte-identical x87 packed/original recovery and material-point comparisons;
- 48 independent dense-reference normal-equation cases plus rejection controls;
- an analytic partial-water unit square: weight 19, reaction X 14.7, Y 23.9;
- associated-flow recovery after one injected line-search rejection, with native
  AUTO and OFF agreeing in displacement to numerical precision;
- native recovery with physical-water material updates remaining in VBA.

Fixtures are in `tests/fixtures/static_density`. The native checks are
`tests/native/test_regularized_rowpack.py` and
`tests/native/test_recovery_precision.py`. Excel checks accept an isolated fixture
workbook via `--book`: `tests/check_associated_recovery.py` and
`tests/check_water_recovery_kernel.py`. These drivers only alter disposable test
copies and are not imported into production workbooks.
