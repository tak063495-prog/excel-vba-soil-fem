/* soilfem_native_x64.dll : coarse-grained band kernels for VBA 2D Soil FEM.
 * One call factors the whole band, one call solves the whole RHS.
 * VBA owns every array (no allocation). Strict mode follows FEMSolver.bas
 * via x87 kernels; optional per-thread double modes use SSE2 or AVX2/FMA.
 * VBA 2-D arrays are column-major: a(i,j) is
 * at a[i + j*ld].
 * SPDX-License-Identifier: Apache-2.0 (same as repository)
 */
#include <stdint.h>
#include <math.h>
#include <intrin.h>
#include <emmintrin.h>
#include "sfn_simd.h"
#pragma fp_contract(off)
#define API __declspec(dllexport) int32_t __stdcall
enum { SF_OK = 0, SF_PIVOT = 1, SF_NONFINITE = 2, SF_BAD = 100 };
#define SF_ABI 0x00010001

static int fin(double v) { return isfinite(v) ? 1 : 0; }
/* x87 kernels (sfn_fp.asm): VBA evaluates Double expressions on the x87
 * stack and rounds only on store; these reproduce that bit for bit. */
extern void sfa_axpy_sub(int64_t n, double *y, int64_t ys, const double *x, int64_t xs, const double *s);
extern void sfa_dot_add(int64_t n, const double *a, int64_t as, const double *b, int64_t bs, double *s);
extern void sfa_dot_sub(int64_t n, const double *a, int64_t as, const double *b, int64_t bs, double *s);
extern void sfa_div(const double *a, const double *b, double *out);
extern void sfa_subdiv(const double *r, const double *s, const double *p, double *out);
extern void sfa_mulsub(const double *y, const double *s, const double *x, double *out);
#define D8 ((int64_t)sizeof(double))

/* Per-thread selection: old workbooks remain strict. No process FPU state
 * changes, and baseline code is compiled without /arch:AVX2. */
static __declspec(thread) int precision_mode;
API sfn_cpu_features(void) {
    int r[4], flags = 1;
    __cpuid(r, 0);
    if (r[0] < 7) return flags;
    __cpuidex(r, 1, 0);
    if ((r[2] & ((1 << 27) | (1 << 28) | (1 << 12))) != ((1 << 27) | (1 << 28) | (1 << 12))) return flags;
    if ((_xgetbv(0) & 6) != 6) return flags;
    __cpuidex(r, 7, 0);
    if (r[1] & (1 << 5)) flags |= 2;
    return flags;
}
API sfn_set_precision(int32_t mode) {
    if (mode < 0 || mode > 2) return SF_BAD;
    precision_mode = mode == 2 && !(sfn_cpu_features() & 2) ? 1 : mode;
    return SF_OK;
}
API sfn_get_precision(void) { return precision_mode; }

static void fp_axpy(int64_t n, double *y, int64_t ys, const double *x, int64_t xs, const double *s) {
    int64_t i = 0; double scale = *s;
    if (!precision_mode) { sfa_axpy_sub(n, y, ys, x, xs, s); return; }
    /* Strided scatter is memory-bound; scalar SSE2 avoids gather/scatter
     * overhead. AVX2/FMA is reserved for contiguous band row updates. */
    if (precision_mode == 2 && ys == D8 && xs == D8) { sfv_axpy(n, y, x, scale); return; }
    ys /= D8; xs /= D8;
    if (ys == 1 && xs == 1) {
        __m128d v = _mm_set1_pd(scale);
        for (; i + 2 <= n; i += 2)
            _mm_storeu_pd(y + i, _mm_sub_pd(_mm_loadu_pd(y + i), _mm_mul_pd(v, _mm_loadu_pd(x + i))));
    }
    for (; i < n; i++) y[i * ys] -= scale * x[i * xs];
}
static double fp_dot(int64_t n, const double *a, int64_t as, const double *b, int64_t bs) {
    int64_t i; double sum = 0.0;
    if (precision_mode == 2) return sfv_dot(n, a, as, b, bs);
    as /= D8; bs /= D8;
    for (i = 0; i < n; i++) sum += a[i * as] * b[i * bs];
    return sum;
}
static void fp_dot_add(int64_t n, const double *a, int64_t as, const double *b, int64_t bs, double *s) {
    if (!precision_mode) sfa_dot_add(n, a, as, b, bs, s);
    else *s += fp_dot(n, a, as, b, bs);
}
static void fp_dot_sub(int64_t n, const double *a, int64_t as, const double *b, int64_t bs, double *s) {
    if (!precision_mode) sfa_dot_sub(n, a, as, b, bs, s);
    else *s -= fp_dot(n, a, as, b, bs);
}
static void fp_div(const double *a, const double *b, double *out) {
    if (!precision_mode) sfa_div(a, b, out); else *out = *a / *b;
}
static void fp_subdiv(const double *r, const double *s, const double *p, double *out) {
    if (!precision_mode) sfa_subdiv(r, s, p, out); else *out = *r - *s / *p;
}
static void fp_mulsub(const double *y, const double *s, const double *x, double *out) {
    if (!precision_mode) sfa_mulsub(y, s, x, out); else *out = *y - *s * *x;
}
#define sfa_axpy_sub fp_axpy
#define sfa_dot_add fp_dot_add
#define sfa_dot_sub fp_dot_sub
#define sfa_div fp_div
#define sfa_subdiv fp_subdiv
#define sfa_mulsub fp_mulsub

API sfn_abi(void) { return SF_ABI; }

/* Symmetric band (upper storage, row i holds a(i,i..i+bw)).
 * band : VBA P6FactoredBand(0..ld-1, 0..bw), col-major, unchanged.
 * packed: output, n*(bw+1), row-major packed factor (= P6PackBand + P6BandFactor1D).
 * akj   : work, >= bw+1. */
API sfn_band_factor(int32_t n, int32_t bw, int32_t ld, const double *band,
                    double *packed, double *akj, double tol, int32_t *failK)
{
    int64_t stride, k, i, j, m1, rowK, rowIt;
    double pivot, aikv;
    if (!failK) return SF_BAD;
    *failK = -1;
    if (n < 1 || bw < 0 || ld < n || !band || !packed || !akj) return SF_BAD;
    stride = (int64_t)bw + 1;
    for (i = 0; i < n; i++)
        for (j = 0; j <= bw; j++)
            packed[i * stride + j] = band[i + j * (int64_t)ld];
    for (k = 0; k < n; k++) {
        int64_t m = (int64_t)bw + 1;
        if (k + bw + 1 > n) m = n - k;
        m1 = m - 1;
        rowK = k * stride;
        pivot = packed[rowK];
        if (fabs(pivot) <= tol) { *failK = (int32_t)k; return SF_PIVOT; }
        for (j = 1; j <= m1; j++) sfa_div(&packed[rowK + j], &pivot, &akj[j]);
        for (i = 1; i <= m1; i++) {
            rowIt = (k + i) * stride;
            aikv = packed[rowK + i];
            sfa_axpy_sub(m1 - i + 1, &packed[rowIt], D8, &akj[i], D8, &aikv);
        }
    }
    return SF_OK;
}

/* Same as P6BandSolve1D: rhs is modified (forward), x receives solution. */
API sfn_band_solve(int32_t n, int32_t bw, const double *packed, double *rhs,
                   double *x, double tol, int32_t *failK)
{
    int64_t stride, k, m, rowK;
    double s, piv;
    if (!failK) return SF_BAD;
    *failK = -1;
    if (n < 1 || bw < 0 || !packed || !rhs || !x) return SF_BAD;
    stride = (int64_t)bw + 1;
    for (k = 0; k <= (int64_t)n - 2; k++) {
        rowK = k * stride; piv = packed[rowK];
        if (fabs(piv) <= tol) { *failK = (int32_t)k; return SF_PIVOT; }
        sfa_div(&rhs[k], &piv, &rhs[k]);
        m = (int64_t)bw + 1; if (k + bw + 1 > n) m = n - k;
        { double rk = rhs[k]; sfa_axpy_sub(m - 1, &rhs[k + 1], D8, &packed[rowK + 1], D8, &rk); }
    }
    rowK = ((int64_t)n - 1) * stride; piv = packed[rowK];
    if (fabs(piv) <= tol) { *failK = n - 1; return SF_PIVOT; }
    sfa_div(&rhs[n - 1], &piv, &x[n - 1]);
    rhs[n - 1] = 0.0;
    for (k = (int64_t)n - 2; k >= 0; k--) {
        rowK = k * stride; s = 0.0;
        m = (int64_t)bw + 1; if (k + bw + 1 > n) m = n - k;
        sfa_dot_add(m - 1, &packed[rowK + 1], D8, &x[k + 1], D8, &s);
        sfa_subdiv(&rhs[k], &s, &packed[rowK], &x[k]);
    }
    return SF_OK;
}

/* Nonsymmetric band LU with partial pivoting (P6SolveNonsymmetricBand).
 * a     : VBA a(0..ld-1, 0..3b) col-major, column index = col - row + b.
 * lower : VBA lowerMult(0..ld-1, 0..b) col-major.
 * reason: 1 ZERO_PIVOT, 2 NONFINITE_MULTIPLIER. Values identical to VBA. */
#define A(r,c) a[(r) + ((int64_t)(c) - (r) + b) * (int64_t)ld]
API sfn_nsband_factor(int32_t n, int32_t b, int32_t ld, double *a, double *rhs,
                      int32_t *pivotRows, double *lower, int32_t *failK,
                      int32_t *reason, double *largestOut)
{
    int64_t k, i, j, lastRow, lastCol, pr, last;
    int64_t upper;
    double largest, tmp, pivot, mult;
    if (!failK || !reason || !largestOut) return SF_BAD;
    *failK = -1; *reason = 0; *largestOut = 0.0;
    if (n < 1 || b < 0 || ld < n || !a || !rhs || !pivotRows || !lower) return SF_BAD;
    last = (int64_t)n - 1; upper = 2 * (int64_t)b;
    for (k = 0; k <= last; k++) {
        lastRow = k + b; if (lastRow > last) lastRow = last;
        lastCol = k + upper; if (lastCol > last) lastCol = last;
        pr = k; largest = fabs(A(k, k));
        for (i = k + 1; i <= lastRow; i++) {
            tmp = fabs(A(i, k));
            if (tmp > largest) { largest = tmp; pr = i; }
        }
        *largestOut = largest;
        if (!fin(largest) || largest <= 1e-30) { *failK = (int32_t)k; *reason = 1; return SF_PIVOT; }
        pivotRows[k] = (int32_t)pr;
        if (pr != k) {
            for (j = k; j <= lastCol; j++) { tmp = A(k, j); A(k, j) = A(pr, j); A(pr, j) = tmp; }
            tmp = rhs[k]; rhs[k] = rhs[pr]; rhs[pr] = tmp;
        }
        pivot = A(k, k);
        for (i = k + 1; i <= lastRow; i++) {
            sfa_div(&A(i, k), &pivot, &mult);
            if (!fin(mult)) { *failK = (int32_t)k; *reason = 2; return SF_PIVOT; }
            A(i, k) = 0.0;
            lower[i + (i - k) * (int64_t)ld] = mult;
            if (mult != 0.0) {
                if (lastCol >= k + 1)
                    sfa_axpy_sub(lastCol - k, &A(i, k + 1), D8 * ld, &A(k, k + 1), D8 * ld, &mult);
                sfa_mulsub(&rhs[i], &mult, &rhs[k], &rhs[i]);
            }
        }
    }
    return SF_OK;
}

/* Forward replay of stored exchanges/multipliers (P6BandReplayForward and
 * the refinement forward pass). */
API sfn_nsband_forward(int32_t n, int32_t b, int32_t ld, const int32_t *pivotRows,
                       const double *lower, double *rhs)
{
    int64_t k, lastRow, last, pr; double tmp;
    if (n < 1 || b < 0 || ld < n || !pivotRows || !lower || !rhs) return SF_BAD;
    last = (int64_t)n - 1;
    for (k = 0; k <= last; k++) {
        pr = pivotRows[k];
        if (pr < 0 || pr > last) return SF_BAD;
        if (pr != k) { tmp = rhs[k]; rhs[k] = rhs[pr]; rhs[pr] = tmp; }
        lastRow = k + b; if (lastRow > last) lastRow = last;
        if (lastRow >= k + 1) { double rk = rhs[k];
            sfa_axpy_sub(lastRow - k, &rhs[k + 1], D8, &lower[(k + 1) + (int64_t)ld], D8 * (ld + 1), &rk); }
    }
    return SF_OK;
}

/* Back substitution with the U part. checkPivot=1 rejects |diag|<=1e-30
 * before dividing (refinement path); returns SF_NONFINITE with failK on a
 * nonfinite solution component (main path). */
API sfn_nsband_back(int32_t n, int32_t b, int32_t ld, const double *a, const double *rhs,
                    double *sol, int32_t checkPivot, int32_t *failK)
{
    int64_t i, lastCol, last, upper; double tmp;
    if (!failK) return SF_BAD;
    *failK = -1;
    if (n < 1 || b < 0 || ld < n || !a || !rhs || !sol) return SF_BAD;
    last = (int64_t)n - 1; upper = 2 * (int64_t)b;
    for (i = last; i >= 0; i--) {
        tmp = rhs[i]; lastCol = i + upper; if (lastCol > last) lastCol = last;
        if (lastCol >= i + 1) sfa_dot_sub(lastCol - i, &A(i, i + 1), D8 * ld, &sol[i + 1], D8, &tmp);
        if (checkPivot && fabs(A(i, i)) <= 1e-30) { *failK = (int32_t)i; return SF_PIVOT; }
        sfa_div(&tmp, &A(i, i), &sol[i]);
        if (!checkPivot && !fin(sol[i])) { *failK = (int32_t)i; return SF_NONFINITE; }
    }
    return SF_OK;
}

/* Optional recovery factorization uses the selected SSE2/AVX2 precision.
 * Material updates and STRICT arithmetic remain in the original x87 core.
 * Acceptance checks the actual normal residual before returning any direction. */
API sfn_regularized(const int64_t *R, double damping, double itTol) {
  int32_t n = (int32_t)R[0]; const int32_t *rp = (const int32_t *)(intptr_t)R[1], *ci = (const int32_t *)(intptr_t)R[2];
  const double *O = (const double *)(intptr_t)R[3]; const int32_t *nc = (const int32_t *)(intptr_t)R[4];
  const double *F = (const double *)(intptr_t)R[5]; double *D = (double *)(intptr_t)R[6], *a = (double *)(intptr_t)R[7];
  int32_t bw = (int32_t)R[8]; double *vw = (double *)(intptr_t)R[9], *o = (double *)(intptr_t)R[10];
  double *rhs = vw, *sol = vw + n, *cs = vw + 2 * n;
  int32_t row, col, p, q, c1, c2, i, j, fc, lr; double v, normB = 0.0, nres = 0.0, tmp;
#define A_(r, c) a[(int64_t)(r) * ((int64_t)bw + 1) + (c)]
  o[0] = 0; o[1] = 0;
  if (precision_mode == 0) return -1; /* preserve the frozen x87 core in STRICT mode */
  if (n < 1 || bw < 0) return -1;

  for (i = 0; i < n; i++) { rhs[i] = 0.0; sol[i] = 0.0; cs[i] = 0.0; }
  for (i = 0; i < n * (int64_t)(bw + 1); i++) a[i] = 0.0;
  for (row = 0; row < n; row++) if (nc[row] == 0) {
    if (!fin(F[row])) return 0;
    for (p = rp[row]; p < rp[row + 1]; p++) { col = ci[p]; if (nc[col] == 0) { v = O[p]; if (!fin(v)) return 0; cs[col] = cs[col] + v * v; } }
  }
  for (col = 0; col < n; col++) { cs[col] = sqrt(cs[col]); if (cs[col] < 1E-30) cs[col] = 1.0; }
  for (row = 0; row < n; row++) if (nc[row] == 0) {
    for (p = rp[row]; p < rp[row + 1]; p++) {
      c1 = ci[p];
      if (nc[c1] == 0 && O[p] != 0.0) {
        v = O[p] / cs[c1];
        rhs[c1] = rhs[c1] + v * F[row];
        for (q = p; q < rp[row + 1]; q++) {
          c2 = ci[q];
          if (nc[c2] == 0 && O[q] != 0.0) {
            if (c2 < c1 || c2 - c1 > bw) return 0;
            A_(c2, bw + c1 - c2) = A_(c2, bw + c1 - c2) + v * O[q] / cs[c2];
          }
        }
      }
    }
  }
  for (i = 0; i < n; i++) {
    A_(i, bw) = A_(i, bw) + damping;
    if (nc[i] != 0) A_(i, bw) = 1.0;
    normB = normB + rhs[i] * rhs[i];
  }
  if (!fin(normB) || normB <= 1E-60) return 0;
  for (i = 0; i < n; i++) {
    fc = i - bw; if (fc < 0) fc = 0;
    for (j = fc; j <= i; j++) {
      tmp = A_(i, bw + j - i);
      if (j > fc) fp_dot_sub(j - fc, &A_(i, bw + fc - i), D8, &A_(j, bw + fc - j), D8, &tmp);
      if (i == j) { if (tmp <= 0.0 || !fin(tmp)) return 0; A_(i, bw) = sqrt(tmp); }
      else A_(i, bw + j - i) = tmp / A_(j, bw);
    }
    tmp = rhs[i];
    if (i > fc) fp_dot_sub(i - fc, &A_(i, bw + fc - i), D8, &sol[fc], D8, &tmp);
    sol[i] = tmp / A_(i, bw);
  }
  o[1] = 1;
  for (i = n - 1; i >= 0; i--) {
    tmp = sol[i]; lr = i + bw; if (lr > n - 1) lr = n - 1;
    if (lr > i) fp_dot_sub(lr - i, &A_(i + 1, bw - 1), (int64_t)bw * D8, &sol[i + 1], D8, &tmp);
    sol[i] = tmp / A_(i, bw);
    if (!fin(sol[i])) return 0;
  }
  for (row = 0; row < n; row++) if (nc[row] == 0) {
    tmp = 0.0;
    for (p = rp[row]; p < rp[row + 1]; p++) { col = ci[p]; if (nc[col] == 0) tmp = tmp + O[p] * sol[col] / cs[col]; }
    for (p = rp[row]; p < rp[row + 1]; p++) { col = ci[p]; if (nc[col] == 0) rhs[col] = rhs[col] - O[p] * tmp / cs[col]; }
  }
  for (i = 0; i < n; i++) { tmp = rhs[i] - damping * sol[i]; nres = nres + tmp * tmp; }
  if (!fin(nres)) return 0;
  { double q2 = nres / normB; o[0] = sqrt(q2); }
  if (o[0] > itTol) return 0;
  for (i = 0; i < n; i++) { D[i] = sol[i] / cs[i]; if (nc[i] != 0) D[i] = 0.0; }

  return 1;
#undef A_
}
