#include "sfn_state.h"
#include <string.h>
#include <limits.h>

/* Sizes follow the checked inline VBA layouts; the caller owns all buffers. */
static const size_t FSZ[9] = {544,128,32,32,8,96,8,1536,288};

static int is_active(const int16_t *active, int32_t lo, int32_t hi, int32_t k) {
    return !active || k < lo || k > hi || active[(int64_t)k - lo] != 0;
}

SFN_API sfn_restore_committed(int32_t n, int64_t stride,
    const int64_t *dst, const int64_t *src, const int16_t *active,
    int32_t lo, int32_t hi, int16_t *plasticActive, int32_t *count) {
    int64_t delta = 0; int32_t k, im;
    if (n < 1 || !dst || !src || !plasticActive || !count) return 100;
    if (stride < 0 || (n > 1 && stride < 1536)) return 100;
    if (n > 1 && stride > INT64_MAX / (n - 1)) return 100;
    if (active && hi < lo) return 100;
    for (k=0;k<9;k++) {
        if (!dst[k] || !src[k]) return 100;
    }
    for (k=0;k<n;k++) {
        for (im=0;im<4;im++) {
            int oldv = plasticActive[4*(int64_t)k+im] != 0;
            int y = ((const int16_t *)(uintptr_t)src[4])[4*(int64_t)k+im] != 0;
            int ok = is_active(active, lo, hi, k);
            delta += (ok && y) - oldv;
        }
    }
    if ((int64_t)*count + delta > INT32_MAX || (int64_t)*count + delta < INT32_MIN) return 100;
    for (k=0;k<n;k++) {
        memcpy((char *)(uintptr_t)dst[0] + (int64_t)k*stride, (const char *)(uintptr_t)src[0] + (int64_t)k*stride, FSZ[0]);
        for (im=1; im<9; im++)
            memcpy((char *)(uintptr_t)dst[im] + (int64_t)k*stride, (const char *)(uintptr_t)src[im] + (size_t)k*FSZ[im], FSZ[im]);
        for (im=0;im<4;im++) {
            int y=((const int16_t *)(uintptr_t)src[4])[4*(int64_t)k+im]!=0;
            plasticActive[4*(int64_t)k+im]=(int16_t)((is_active(active, lo, hi, k) && y)?-1:0);
        }
    }
    *count = (int32_t)((int64_t)*count + delta); return 0;
}
