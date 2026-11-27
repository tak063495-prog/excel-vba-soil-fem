/* SoilFEM SRM band accelerator, ABI 0x10002. Windows x64 only.
 * Direct translation of FEMSolver.P6PackBand/P6BandFactor1D/P6BandSolve1D.
 * VBA x87 intermediates, rounded to double at each original assignment.
 * No reassociation, reciprocal approximation or FMA.
 * No allocation, threads, global floating-point state or external data access.
 * VBA owns all buffers. Input/output/work buffers must be disjoint.
 * Return 0=success, 1=the original pivot rejection, -2=invalid dimensions.
 */
#include <stdint.h>
#include <stddef.h>
#define fabs __builtin_fabs
#define API __declspec(dllexport)
/* No CRT initialization: leave Excel's floating-point environment untouched. */
int _DllMainCRTStartup(void *module, unsigned long reason, void *reserved)
{ (void)module; (void)reason; (void)reserved; return 1; }

API int32_t sfbs_abi(void) { return 0x10002; }

API double sfbs_probe(double a, double b, double c, int32_t op)
{
    if (op==0) return (double)((long double)a-(long double)b*(long double)c);
    if (op==1) return (double)((long double)a+(long double)b*(long double)c);
    if (op==2) return (double)((long double)a-(long double)b/(long double)c);
    return (double)((long double)b/(long double)c);
}

API int32_t sfbs_factor(int32_t n, int32_t bw, int32_t ld,
    const double * restrict band, double * restrict packed,
    double * restrict akj, double tol, int32_t * restrict fail_k)
{
    if (!band || !packed || !akj || !fail_k || n<1 || bw<0 || ld<n ||
        (uint64_t)n*((uint64_t)bw+1)>INT32_MAX || !(tol>=0)) return -2;
    *fail_k=-1;
    size_t stride=(size_t)bw+1;
    for (int32_t i=0;i<n;i++)
        for (int32_t j=0;j<=bw;j++)
            packed[(size_t)i*stride+j]=band[i+(size_t)ld*j];
    for (int32_t k=0;k<n;k++) {
        int32_t m1=bw;
        if (k+bw+1>n) m1=n-k-1;
        size_t row_k=(size_t)k*stride;
        double pivot=packed[row_k];
        if (fabs(pivot)<=tol) { *fail_k=k; return 1; }
        for (int32_t j=1;j<=m1;j++) akj[j]=(double)((long double)packed[row_k+j]/(long double)pivot);
        for (int32_t i=1;i<=m1;i++) {
            size_t row_it=(size_t)(k+i)*stride;
            /* The pivot row is untouched by this k update, so this is the
             * exact value stored in the original VBA aik(i) buffer. */
            double aikv=packed[row_k+i];
            int32_t lim=m1-i;
            for (int32_t t=0;t<=lim;t++)
                packed[row_it+t]=(double)((long double)packed[row_it+t]-(long double)aikv*(long double)akj[t+i]);
        }
    }
    return 0;
}

API int32_t sfbs_solve(int32_t n, int32_t bw,
    const double * restrict packed, double * restrict rhs,
    double * restrict x, double tol, int32_t * restrict fail_k)
{
    if (!packed || !rhs || !x || !fail_k || n<1 || bw<0 ||
        (uint64_t)n*((uint64_t)bw+1)>INT32_MAX || !(tol>=0)) return -2;
    *fail_k=-1;
    size_t stride=(size_t)bw+1;
    for (int32_t k=0;k<n-1;k++) {
        size_t row_k=(size_t)k*stride;
        double piv=packed[row_k];
        if (fabs(piv)<=tol) { *fail_k=k; return 1; }
        rhs[k]=(double)((long double)rhs[k]/(long double)piv);
        int32_t m1=bw;
        if (k+bw+1>n) m1=n-k-1;
        for (int32_t i=1;i<=m1;i++)
            rhs[k+i]=(double)((long double)rhs[k+i]-(long double)packed[row_k+i]*(long double)rhs[k]);
    }
    size_t row_k=(size_t)(n-1)*stride;
    double piv=packed[row_k];
    if (fabs(piv)<=tol) { *fail_k=n-1; return 1; }
    x[n-1]=(double)((long double)rhs[n-1]/(long double)piv);
    rhs[n-1]=0.0;
    for (int32_t k=n-2;k>=0;k--) {
        row_k=(size_t)k*stride;
        double s=0.0;
        int32_t m1=bw;
        if (k+bw+1>n) m1=n-k-1;
        /* Preserve the VBA sum order exactly. */
        for (int32_t j=1;j<=m1;j++) s=(double)((long double)s+(long double)packed[row_k+j]*(long double)x[k+j]);
        x[k]=(double)((long double)rhs[k]-(long double)s/(long double)packed[row_k]);
    }
    return 0;
}
