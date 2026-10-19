#include <windows.h>
#include <stdint.h>
#include <math.h>
#include <intrin.h>

static __forceinline double dot_tuned(int32_t s, int32_t e, const int32_t *ci, const double *v, const double *x) {
  __m128d a=_mm_setzero_pd(), b=_mm_setzero_pd(); double z[2], sum; int32_t p=s;
  for (; e-p>=4; p+=4) {
    __m128d x0=_mm_set_pd(x[ci[p+1]],x[ci[p]]), x1=_mm_set_pd(x[ci[p+3]],x[ci[p+2]]);
    a=_mm_add_pd(a,_mm_mul_pd(_mm_loadu_pd(v+p),x0)); b=_mm_add_pd(b,_mm_mul_pd(_mm_loadu_pd(v+p+2),x1));
  }
  _mm_storeu_pd(z,_mm_add_pd(a,b)); sum=z[0]+z[1];
  for (; p<e; ++p) sum += v[p]*x[ci[p]];
  return sum;
}
static int run_variant(int32_t n,const int32_t *rp,const int32_t *ci,const int32_t *dp,const double *ilu,const double *inv,const double *in,double *out,int variant) {
  int32_t r,p;
  for (r=0;r<n;r++) {
    double value=in[r];
    if (variant==0) for (p=rp[r];p<dp[r];p++) value-=ilu[p]*out[ci[p]];
    else value-=dot_tuned(rp[r],dp[r],ci,ilu,out);
    out[r]=value;
  }
  for (r=n-1;r>=0;r--) {
    double value=out[r];
    if (variant==0) for (p=dp[r]+1;p<rp[r+1];p++) value-=ilu[p]*out[ci[p]];
    else value-=dot_tuned(dp[r]+1,rp[r+1],ci,ilu,out);
    if (fabs(ilu[dp[r]])<=1e-12) return 0;
    out[r]=(variant==2)?value*inv[r]:value/ilu[dp[r]];
  }
  return 1;
}
/* One export: variants 0 scalar, 1 tuned indexed-dot, 2 tuned+inverse diagonal.
 * out[0..3] are seconds; result buffers receive one representative result. */
__declspec(dllexport) int __cdecl ilu_probe_run(int32_t n,const int32_t *rp,const int32_t *ci,const int32_t *dp,const double *ilu,const double *in,
  int32_t reps,double *out0,double *out1,double *out2,double *seconds) {
  int v,k; LARGE_INTEGER f,a,b; double *tmp=(double*)_alloca((size_t)n*sizeof(double)); double *inv=(double*)_alloca((size_t)n*sizeof(double));
  if (!rp||!ci||!dp||!ilu||!in||!out0||!out1||!out2||!seconds||n<1||reps<1) return -1;
  QueryPerformanceFrequency(&f); QueryPerformanceCounter(&a); for(k=0;k<n;k++) inv[k]=1.0/ilu[dp[k]]; QueryPerformanceCounter(&b); seconds[3]=(double)(b.QuadPart-a.QuadPart)/(double)f.QuadPart;
  for(v=0;v<3;v++) { QueryPerformanceCounter(&a); for(k=0;k<reps;k++) if(!run_variant(n,rp,ci,dp,ilu,inv,in,tmp,v)) return 0; QueryPerformanceCounter(&b); seconds[v]=(double)(b.QuadPart-a.QuadPart)/(double)f.QuadPart;
    if(v==0) { for(k=0;k<n;k++) out0[k]=tmp[k]; } else if(v==1) { for(k=0;k<n;k++) out1[k]=tmp[k]; } else { for(k=0;k<n;k++) out2[k]=tmp[k]; }
  }
  seconds[4]=(double)reps; return 1;
}
