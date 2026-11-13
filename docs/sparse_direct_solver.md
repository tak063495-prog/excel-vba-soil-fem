# Optional sparse direct solver

`NATIVE_SPARSE_DIRECT=1` uses `soilfem_sparse_x64.dll` beside the workbook for
an eligible symmetric band operator. The default is 0. `NATIVE_DLL=OFF`
retains the VBA/legacy path. Nonassociated Newton, JOINT and coupled operators keep
their existing solver selection. An existing symmetric elastic-correction retry
may use sparse LDLT while retaining the original nonassociated force law.
Missing/incompatible DLLs and rejected sparse
factorizations fall back to the existing band solver.

The optional DLL uses unmodified Eigen 3.4.0, AMD ordering, diagonal scaling and
sparse LDLT. It retains the original free-free coefficients, eliminates
prescribed displacement exactly, checks positive scaled pivots, and verifies
the actual unscaled linear residual with iterative refinement. A token owns
each factor; one workbook cannot accidentally reuse another workbook's factor.
Existing tangent/constraint generation and modified-Newton reuse rules remain
in effect. No cohesion, friction, dilation, constitutive force or nonlinear
convergence tolerance is changed.

The band assembly remains in VBA memory. This option replaces factorization;
it does not turn the whole application into a CSR-only program. The existing
band-memory guard remains. The DLL checks an estimate of its created factor
storage against the configured memory limit; that estimate excludes VBA's band
arrays and is not an operating-system peak-memory limit. Diagnostics show
`SPARSE_LDLT` and `SPARSE_DIRECT_FACTOR` events with scaled pivots, fill and time.

`NATIVE_WATER_TANGENT=1` separately enables the existing native element-matrix
integration with physical groundwater. It integrates `B^T D_alg B` using the
already evaluated `Spmat = D_alg B` after the
VBA material update. The physical-water native **material/force** restriction
remains. Its default is 0. The groundwater constitutive conversion and actual
self-weight vectors retain their existing path.

## Build and distribution

Use 64-bit MSVC and the verified, unmodified Eigen source directory:

```powershell
Invoke-WebRequest 'https://gitlab.com/libeigen/eigen/-/archive/3.4.0/eigen-3.4.0.tar.gz' -OutFile eigen-3.4.0.tar.gz
if ((Get-FileHash eigen-3.4.0.tar.gz -Algorithm SHA256).Hash -ne '8586084F71F9BDE545EE7FA6D00288B264A2B7AC3607B974E54D13E7162C1C72') { throw 'Eigen archive hash mismatch' }
tar -xf eigen-3.4.0.tar.gz
cmd /c native\build_sparse.cmd eigen-3.4.0
python tests/native/test_sparse_direct.py --dll native/bin/sparse/soilfem_sparse_x64.dll --report sparse-native-report.json
```

The build defines `EIGEN_MPL2_ONLY` and uses a static C++ runtime. Distribute the
DLL with `native/EIGEN_NOTICE.txt`, the licenses in `native/licenses/`, and access
to the corresponding Eigen source archive. The wrapper source is Apache-2.0.
The original `soilfem_native_x64.dll` is also needed for the existing optional
kernels; the new DLL does not replace it.

## Validation

The native checker covers independent dense solutions, prescribed values,
reused right-hand sides, in-place output, factor ownership, singular/indefinite
matrices, nonfinite inputs and the memory-estimate rejection. Real Excel checks
in `tests/check_sparse_direct_option.py` compare sparse, band, native OFF and
missing-DLL results and checkpoint resumes for moving supports, groundwater
plasticity, mixed Q8/T6 (dry and wet), opening/recontact, and nonassociated
exclusion. `tests/check_water_tangent_kernel.py` directly compares element
matrices, stresses, weights and self-weight vectors with the VBA path.

These tests do not establish completion or physical stability of an arbitrary
large model. Its gravity, loading and SRM stages must be checked separately.
