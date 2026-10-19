/* Compiled separately with /arch:AVX2; entered only after CPUID + XGETBV.
 * No fast-math: NaN/Inf checks in the caller remain effective. */
#include <immintrin.h>
#include "sfn_simd.h"
static __m256d load4(const double *p, int64_t step) {
    if (step == 1) return _mm256_loadu_pd(p);
    return _mm256_set_pd(p[3 * step], p[2 * step], p[step], p[0]);
}
void sfv_axpy(int64_t n, double *y, const double *x, double s) {
    int64_t i = 0;
    const __m256d v = _mm256_set1_pd(s);
    for (; i + 8 <= n; i += 8) {
        _mm256_storeu_pd(y + i, _mm256_fnmadd_pd(v, _mm256_loadu_pd(x + i), _mm256_loadu_pd(y + i)));
        _mm256_storeu_pd(y + i + 4, _mm256_fnmadd_pd(v, _mm256_loadu_pd(x + i + 4), _mm256_loadu_pd(y + i + 4)));
    }
    for (; i + 4 <= n; i += 4)
        _mm256_storeu_pd(y + i, _mm256_fnmadd_pd(v, _mm256_loadu_pd(x + i), _mm256_loadu_pd(y + i)));
    for (; i < n; i++)
        _mm_store_sd(y + i, _mm_fnmadd_sd(_mm_set_sd(s), _mm_load_sd(x + i), _mm_load_sd(y + i)));
}
double sfv_dot(int64_t n, const double *a, int64_t as, const double *b, int64_t bs) {
    int64_t i = 0; double t[4], sum;
    __m256d acc = _mm256_setzero_pd();
    as /= 8; bs /= 8;
    for (; i + 4 <= n; i += 4)
        acc = _mm256_fmadd_pd(load4(a + i * as, as), load4(b + i * bs, bs), acc);
    _mm256_storeu_pd(t, acc); sum = (t[0] + t[1]) + (t[2] + t[3]);
    for (; i < n; i++) sum += a[i * as] * b[i * bs];
    return sum;
}
