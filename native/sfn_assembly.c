#include "sfn_assembly.h"
#include <limits.h>
#include <stddef.h>

/* Caller owns valid, disjoint arrays. This checks representable spans, not
 * whether arbitrary addresses are mapped. */
static int span(int64_t ptr, uint64_t bytes) {
    return ptr > 0 && bytes > 0 && bytes <= (uint64_t)INTPTR_MAX &&
        (uint64_t)ptr <= (uint64_t)INTPTR_MAX - (bytes - 1);
}

/* IEEE binary64 classification: no arithmetic, including for subnormals.
 * Both signed zeros are excluded, exactly as VBA's value <> 0#. */
static int nonzero(const double *p) {
    union { double d; uint64_t bits; } u;
    uint64_t a;
    u.d = *p;
    a = u.bits & UINT64_C(0x7fffffffffffffff);
    if (a >= UINT64_C(0x7ff0000000000000)) return -1;
    return a != 0;
}

__declspec(dllexport) int32_t __stdcall sfn_csr_scatter_list(const int64_t *P) {
    const char *base; const int32_t *ep;
    int32_t *start, *row, *col, *pos, *vi, *out;
    int32_t n, k, i, j, count = 0, w = 0;
    int64_t stride, offset, cap;
    uint64_t entries, extent;
    if (!P) return -1;
    if (P[2] < 1 || P[2] > INT_MAX / 256 || P[5] < 1 || P[5] > INT_MAX ||
        P[7] < -1 || P[7] > INT_MAX) return -1;
    n = (int32_t)P[2]; stride = P[1]; offset = P[3]; cap = P[7];
    if (stride < 2048 || offset < 0 || offset > stride - 2048 ||
        (uint64_t)stride > (uint64_t)INTPTR_MAX / (uint64_t)n) return -1;
    extent = (uint64_t)stride * (uint64_t)n;
    entries = (uint64_t)n * 256;
    if (!span(P[0], extent) || !span(P[4], entries * 4) ||
        !span(P[6], ((uint64_t)n + 1) * 4) || !span(P[12], 4)) return -1;
    if (cap >= 0) {
        uint64_t bytes = (uint64_t)(cap ? cap : 1) * 4;
        if (!span(P[8], bytes) || !span(P[9], bytes) ||
            !span(P[10], bytes) || !span(P[11], bytes)) return -1;
    }
    base = (const char *)(intptr_t)P[0]; ep = (const int32_t *)(intptr_t)P[4];
    start = (int32_t *)(intptr_t)P[6]; out = (int32_t *)(intptr_t)P[12];
    row = (int32_t *)(intptr_t)P[8]; col = (int32_t *)(intptr_t)P[9];
    pos = (int32_t *)(intptr_t)P[10]; vi = (int32_t *)(intptr_t)P[11];
    /* Validate everything before changing outputs, including count-only. */
    for (k = 0; k < n; ++k) {
        const double *K = (const double *)(base + stride * k + offset);
        for (i = 0; i < 16; ++i) for (j = 0; j < 16; ++j) {
            int nz = nonzero(&K[i + 16 * j]);
            int64_t index = k + (int64_t)n * (i + 16 * j);
            if (nz < 0) return -1;
            if (nz) {
                if (ep[index] < 0 || ep[index] >= P[5]) return -1;
                ++count;
            }
        }
    }
    if (cap >= 0 && count > cap) return -1;
    for (k = 0; k < n; ++k) {
        const double *K = (const double *)(base + stride * k + offset);
        start[k] = w;
        for (i = 0; i < 16; ++i) for (j = 0; j < 16; ++j) {
            if (nonzero(&K[i + 16 * j])) {
                if (cap >= 0) {
                    row[w] = i; col[w] = j;
                    pos[w] = ep[k + (int64_t)n * (i + 16 * j)];
                    vi[w] = k * 256 + i * 16 + j;
                }
                ++w;
            }
        }
    }
    start[n] = count; *out = count;
    return 0;
}
