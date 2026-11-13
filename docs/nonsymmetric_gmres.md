# Optional CSR-GMRES for nonassociated flow

`FLOW_POLICY=INCONSISTENT` evaluates the original Mohr-Coulomb cohesion,
friction and dilation angles and retains the complete nonsymmetric Newton
operator. `NONSYM_SOLVER=BAND` remains the default pivoted band solver.
Set `NONSYM_SOLVER=GMRES` in the analysis settings to select the existing
CSR right-preconditioned GMRES solver directly. This avoids allocating the
large nonsymmetric band factors. It does not apply the Davis equivalent
strength conversion.

The option uses ILU(0) when available and checks the true CSR residual against
`SOLVER_TOLERANCE`. `SOLVER_GMRES_RESTART`, `SOLVER_MAX_ITERATIONS` and
`SOLVER_MEMORY_LIMIT_MB` continue to apply. GMRES can fail on an ill-conditioned
plastic tangent; its failure alone does not establish an SRM limit state. The
existing regularized and elastic correction retries remain in effect and are
accepted only against the actual nonlinear material-force residual.

For GMRES models at or above `NONSYM_ELASTIC_START_DOF` (default 2000),
start with the elastic correction when its band fits the memory limit. Set
this threshold to 0 to start with full Newton. A failed elastic attempt still
receives the existing clean Newton retry and regularized recovery.

For the elastic correction retry, the correction matrix is symmetric. The
GMRES option uses the band Cholesky path for that retry when its estimated
memory fits the configured limit. The evaluated nonassociated stress and
internal force are unchanged. JOINT, mixed, and fixed-stress coupled operators
retain the nonsymmetric path. Insufficient band memory retains CSR-GMRES.

The optional native CSR residual, ILU(0), and GMRES kernels can now operate
with physical groundwater because they receive only assembled matrix/vector
data. The native material-update restriction under physical groundwater is
preserved: effective/total stress conversion and the material update remain
in VBA. Diagnostics report the actual solver mode and native kernel counters.

Validation uses actual Excel runs and checkpoint resumes: groundwater with
plastic Mohr-Coulomb points, mixed Q8/T6 gravity, and a moving-wall support
case. Native AUTO and OFF are compared with the pivoted BAND reference.
`tests/check_gmres_solver_option.py --book <compiled.xlsm> --out <new-folder>`
runs the water/plastic, mixed element and automatic elastic-start comparisons. Use the existing
`tests/run_groundwater_fixture.ps1` for the `face_support/gmres_water` fixture.
Test-only forced elastic operators are never part of the product source.
