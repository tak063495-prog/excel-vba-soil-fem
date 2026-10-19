#ifndef SFN_STATE_H
#define SFN_STATE_H
#include <stdint.h>
#ifdef _WIN32
#define SFN_API __declspec(dllexport) int32_t __stdcall
#else
#define SFN_API int32_t
#endif
SFN_API sfn_restore_committed(int32_t n, int64_t stride,
    const int64_t *dst, const int64_t *src, const int16_t *active,
    int32_t lo, int32_t hi, int16_t *plasticActive, int32_t *count);
#endif
