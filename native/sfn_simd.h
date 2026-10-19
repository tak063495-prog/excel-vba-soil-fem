/* Internal double-precision kernels. Strides are bytes. */
#include <stdint.h>
void sfv_axpy(int64_t n, double *y, const double *x, double s);
double sfv_dot(int64_t n, const double *a, int64_t as, const double *b, int64_t bs);
