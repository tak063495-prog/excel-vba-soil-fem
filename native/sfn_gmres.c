/* Right-preconditioned ILU(0) GMRES, from soilfem_mat.c (Apache-2.0).
 * Double SSE2/AVX2 arithmetic only; same two-pass MGS, true-residual checks,
 * tolerances, iteration limits, and failure/event contract as sfm_gmres.
 * G[13]=CSR diagonal positions, G[14]=workspace doubles, G[15]=nnz.
 * Strict mode declines so the original x87 DLL remains the reference.
 * SPDX-License-Identifier: Apache-2.0 */
#include <stdint.h>
#include <limits.h>
#include <math.h>
#include <stdlib.h>
#include <emmintrin.h>
#include "sfn_simd.h"
#pragma fp_contract(off)
extern int32_t __stdcall sfn_get_precision(void);
#define xd double
#define ABS(v) fabs(v)
static double vsqr(double v) { return sqrt(v); }
static double c_pow(double v,double exponent) { (void)exponent; return v*v; }
static int vfin(double v) { return isfinite(v) && fabs(v)<=1E+300; }
static double dot(int32_t n,const double *a,const double *b,int mode) {
    int32_t i=0; double tmp[2],sum;
    __m128d acc=_mm_setzero_pd();
    if(mode==2) return sfv_dot(n,a,8,b,8);
    for(;i+2<=n;i+=2) acc=_mm_add_pd(acc,_mm_mul_pd(_mm_loadu_pd(a+i),_mm_loadu_pd(b+i)));
    _mm_storeu_pd(tmp,acc); sum=tmp[0]+tmp[1];
    for(;i<n;i++) sum+=a[i]*b[i];
    return sum;
}
static void axpy(int32_t n,double *y,const double *x,double scale,int mode) {
    int32_t i=0; __m128d v=_mm_set1_pd(scale);
    if(mode==2) { sfv_axpy(n,y,x,scale); return; }
    for(;i+2<=n;i+=2) _mm_storeu_pd(y+i,_mm_sub_pd(_mm_loadu_pd(y+i),_mm_mul_pd(v,_mm_loadu_pd(x+i))));
    for(;i<n;i++) y[i]-=scale*x[i];
}
static void csr_matvec(int32_t n, const int32_t *rp, const int32_t *ci, const double *v, const double *x, double *y) {
  int32_t r, p;
  for (r = 0; r < n; r++) {
    y[r] = 0.0;
    for (p = rp[r]; p < rp[r + 1]; p++) y[r] = (xd)y[r] + (xd)v[p] * x[ci[p]];
  }
}
/* P6CSRMatVecResidualNorm: r = b - A x, returns ||r|| */
static double csr_resnorm(int32_t n, const int32_t *rp, const int32_t *ci, const double *v, const double *x, const double *b, double *r) {
  int32_t i; double rv, ss = 0.0;
  csr_matvec(n, rp, ci, v, x, r);
  for (i = 0; i < n; i++) { rv = (xd)b[i] - r[i]; r[i] = rv; ss = (xd)ss + (xd)rv * rv; }
  if (ss <= 0.0) return 0.0;
  return vsqr(ss);
}
static int ilu_apply(int32_t n, const int32_t *rp, const int32_t *ci, const int32_t *dp, const double *ilu, const double *in, double *out) {
    int32_t r,p; double value,diag;
    for(r=0;r<n;r++) {
        value=in[r];
        for(p=rp[r];p<dp[r];p++) value-=ilu[p]*out[ci[p]];
        out[r]=value;
    }
    for(r=n-1;r>=0;r--) {
        value=out[r]; diag=ilu[dp[r]];
        for(p=dp[r]+1;p<rp[r+1];p++) value-=ilu[p]*out[ci[p]];
        if(fabs(diag)<=0.000000000001) return 0;
        out[r]=value/diag;
    }
    return 1;
}

/* Binary64 SIMD reductions; caller retains the original scalar true-residual
 * calculation, convergence tolerance, two-pass MGS and iteration limits. */
static __forceinline double indexed_dot(int32_t start,int32_t end,const int32_t *ci,
    const double *v,const double *x) {
    int32_t p=start; double tmp[2], sum;
    __m128d a=_mm_setzero_pd(), b=_mm_setzero_pd();
    for(;end-p>=4;p+=4) {
        __m128d x0=_mm_set_pd(x[ci[p+1]],x[ci[p]]);
        __m128d x1=_mm_set_pd(x[ci[p+3]],x[ci[p+2]]);
        a=_mm_add_pd(a,_mm_mul_pd(_mm_loadu_pd(v+p),x0));
        b=_mm_add_pd(b,_mm_mul_pd(_mm_loadu_pd(v+p+2),x1));
    }
    _mm_storeu_pd(tmp,_mm_add_pd(a,b));sum=tmp[0]+tmp[1];
    for(;p<end;p++) sum+=v[p]*x[ci[p]];
    return sum;
}
static void csr_matvec_tuned(int32_t n,const int32_t *rp,const int32_t *ci,
    const double *v,const double *x,double *y) {
    int32_t r;
    for(r=0;r<n;r++) y[r]=indexed_dot(rp[r],rp[r+1],ci,v,x);
}
static int ilu_apply_tuned(int32_t n,const int32_t *rp,const int32_t *ci,const int32_t *dp,
    const double *ilu,const double *in,double *out) {
    int32_t r;
    for(r=0;r<n;r++) out[r]=in[r]-indexed_dot(rp[r],dp[r],ci,ilu,out);
    for(r=n-1;r>=0;r--) {
        double value=out[r]-indexed_dot(dp[r]+1,rp[r+1],ci,ilu,out);
        double diag=ilu[dp[r]];
        if(fabs(diag)<=0.000000000001) return 0;
        out[r]=value/diag;
    }
    return 1;
}

/* The factor is constant throughout one GMRES call. Prepare its diagonal
 * reciprocals once; keep the same pivot rejection and scalar true residual. */
static int ilu_apply_prepared(int32_t n,const int32_t *rp,const int32_t *ci,const int32_t *dp,
    const double *ilu,const double *inverse,const double *in,double *out) {
    int32_t r;
    for(r=0;r<n;r++) out[r]=in[r]-indexed_dot(rp[r],dp[r],ci,ilu,out);
    for(r=n-1;r>=0;r--) {
        double value=out[r]-indexed_dot(dp[r]+1,rp[r+1],ci,ilu,out);
        if(fabs(ilu[dp[r]])<=0.000000000001) return 0;
        out[r]=value*inverse[r];
    }
    return 1;
}

enum { EV_RESTART = 1, EV_NONFIN_TRUE = 2, EV_PREC = 3, EV_NONFIN_KRYLOV = 4, EV_ZERO_GIVENS = 5, EV_ZERO_UPPER = 6 };
static void ev_push(double *ev, int64_t cap, double *o, int t, double a, double b) {
  int64_t k = (int64_t)o[2];
  if (k >= cap) { o[3] = 1.0; return; }
  ev[3 * k] = t; ev[3 * k + 1] = a; ev[3 * k + 2] = b; o[2] = (double)(k + 1);
}
static int32_t gmres_impl(const int64_t *G, double dtol, int tuned, double *inverse) {
  if (!G) return -1;
  int mode = sfn_get_precision();
  int32_t n = (int32_t)G[0]; const int32_t *rp = (const int32_t *)(intptr_t)G[1], *ci = (const int32_t *)(intptr_t)G[2];
  const double *A = (const double *)(intptr_t)G[3], *ilu = (const double *)(intptr_t)G[4], *F = (const double *)(intptr_t)G[5];
  double *X = (double *)(intptr_t)G[6], *wk = (double *)(intptr_t)G[7]; int32_t m = (int32_t)G[8], maxIt = (int32_t)G[9];
  const int32_t *dp = (const int32_t *)(intptr_t)G[13];
  double *ev = (double *)(intptr_t)G[10]; int64_t cap = G[11]; double *o = (double *)(intptr_t)G[12];
  double *V, *Z, *H, *cs, *sn, *g, *y, *b, *x, *r, *w;
  int32_t i, j, k, pass, used, total = 0; int64_t ld = n, hl = (int64_t)m + 1;
  double normB = 0.0, beta, nextNorm, inner, rot, tmp, residual, tol;
  /* All shape/index checks precede writes. VBA supplies owned buffers. */
  if (n<1 || m<1 || m>n || maxIt<1 || mode<1 || mode>2 ||
      G[0]!=n || G[8]!=m || G[9]!=maxIt || !rp || !ci || !dp || !A || !ilu ||
      !F || !X || !wk || !ev || !o || !isfinite(dtol) || dtol<=0 ||
      cap<(int64_t)maxIt+2 || cap>INT32_MAX/3 || G[15]<1 || G[15]>INT32_MAX) return -1;
  if (G[14]<1 || G[14]>200000001 || (uint64_t)G[14] < (uint64_t)n*(2*(uint64_t)m+5)+((uint64_t)m+1)*m+4*(uint64_t)m+1) return -1;
  if (rp[0]!=0 || rp[n]!=G[15]) return -1;
  for (i=0;i<n;i++) {
    if(rp[i]<0 || rp[i+1]<rp[i] || rp[i+1]>G[15] || dp[i]<rp[i] || dp[i]>=rp[i+1] || ci[dp[i]]!=i) return -1;
    for(j=rp[i];j<rp[i+1];j++)
      if(ci[j]<0 || ci[j]>=n || (j>rp[i] && ci[j]<=ci[j-1])) return -1;
  }
  if(inverse) for(i=0;i<n;i++) {
    double diagonal=ilu[dp[i]];
    inverse[i]=(fabs(diagonal)<=0.000000000001) ? 0.0 : 1.0/diagonal;
  }
  o[0] = 1E+100; o[1] = 0; o[2] = 0; o[3] = 0;
  V = wk; Z = V + ld * (m + 1); H = Z + ld * m; cs = H + hl * m; sn = cs + m; g = sn + m; y = g + m + 1;
  b = y + m; x = b + n; r = x + n; w = r + n;
  for (i = 0; i < n; i++) {
    if (!vfin(F[i])) return 0;
    b[i] = F[i]; x[i] = 0.0;
    normB = (xd)normB + (xd)F[i] * F[i];
  }
  normB = vsqr(normB);
  if (!vfin(normB)) return 0;
  if (normB <= 1E-30) { for (i = 0; i < n; i++) X[i] = 0.0; o[0] = 0.0; return 1; }
  tol = (xd)dtol * normB;
  while (total < maxIt) {
    residual = csr_resnorm(n, rp, ci, A, x, b, r);
    if (!vfin(residual)) { ev_push(ev, cap, o, EV_NONFIN_TRUE, total, 0); return 0; }
    o[0] = (xd)residual / normB;
    ev_push(ev, cap, o, EV_RESTART, total, o[0]);
    if (residual <= tol) break;
    beta = residual;
    for (i = 0; i < n; i++) V[i] = (xd)r[i] / beta;
    for (i = 0; i <= m; i++) g[i] = 0.0;
    g[0] = beta; used = 0;
    for (j = 0; j < m; j++) {
      double *zj = Z + ld * j;
      int applied=inverse ? ilu_apply_prepared(n,rp,ci,dp,ilu,inverse,V+ld*j,zj) :
        (tuned ? ilu_apply_tuned(n,rp,ci,dp,ilu,V+ld*j,zj) : ilu_apply(n,rp,ci,dp,ilu,V+ld*j,zj));
      if (!applied) { ev_push(ev, cap, o, EV_PREC, total, 0); return 0; }
      if (tuned) csr_matvec_tuned(n, rp, ci, A, zj, w);
      else csr_matvec(n, rp, ci, A, zj, w);
      for (k = 0; k <= m; k++) H[k + hl * j] = 0.0;
      for (pass = 1; pass <= 2; pass++) {
        for (k = 0; k <= j; k++) {
          const double *vk = V + ld * k;
          inner = 0.0;
          inner = dot(n, w, vk, mode);
          H[k + hl * j] = (xd)H[k + hl * j] + inner;
          axpy(n, w, vk, inner, mode);
        }
      }
      { double ss = dot(n, w, w, mode); nextNorm = (ss <= 0.0) ? 0.0 : vsqr(ss); }
      H[j + 1 + hl * j] = nextNorm;
      if (!vfin(nextNorm)) { ev_push(ev, cap, o, EV_NONFIN_KRYLOV, total, 0); return 0; }
      if (nextNorm > 1E-30) { double *vn = V + ld * (j + 1); for (i = 0; i < n; i++) vn[i] = (xd)w[i] / nextNorm; }
      for (k = 0; k <= j - 1; k++) {
        tmp = (xd)cs[k] * H[k + hl * j] + (xd)sn[k] * H[k + 1 + hl * j];
        H[k + 1 + hl * j] = (xd)(-sn[k]) * H[k + hl * j] + (xd)cs[k] * H[k + 1 + hl * j];
        H[k + hl * j] = tmp;
      }
      { double p1 = c_pow(H[j + hl * j], 2.0), p2 = c_pow(H[j + 1 + hl * j], 2.0); double s2 = (xd)p1 + p2; rot = vsqr(s2); }
      if (rot <= 1E-30) { ev_push(ev, cap, o, EV_ZERO_GIVENS, total, 0); break; }
      cs[j] = (xd)H[j + hl * j] / rot; sn[j] = (xd)H[j + 1 + hl * j] / rot;
      H[j + hl * j] = rot; H[j + 1 + hl * j] = 0.0;
      tmp = (xd)cs[j] * g[j];
      g[j + 1] = (xd)(-sn[j]) * g[j]; g[j] = tmp;
      used = j + 1; total = total + 1; o[1] = total;
      if (ABS(g[j + 1]) <= tol || nextNorm <= 1E-30 || total >= maxIt) break;
    }
    if (used < 1) return 0;
    for (i = used - 1; i >= 0; i--) {
      tmp = g[i];
      for (k = i + 1; k <= used - 1; k++) tmp = (xd)tmp - (xd)H[i + hl * k] * y[k];
      if (ABS(H[i + hl * i]) <= 1E-30) { ev_push(ev, cap, o, EV_ZERO_UPPER, total, 0); return 0; }
      y[i] = (xd)tmp / H[i + hl * i];
    }
    for (k = 0; k <= used - 1; k++) axpy(n, x, Z + ld*k, -y[k], mode);
  }
  residual = csr_resnorm(n, rp, ci, A, x, b, r);
  if (!vfin(residual)) return 0;
  o[0] = (xd)residual / normB;
  if (residual > tol) return 0;
  for (i = 0; i < n; i++) if (!vfin(x[i])) return 0;
  for (i = 0; i < n; i++) X[i] = x[i];
  return 1;
}


__declspec(dllexport) int32_t __stdcall sfn_gmres(const int64_t *G, double tol) { return gmres_impl(G,tol,0,NULL); }
__declspec(dllexport) int32_t __stdcall sfn_gmres_tuned(const int64_t *G, double tol) { return gmres_impl(G,tol,1,NULL); }
__declspec(dllexport) int32_t __stdcall sfn_gmres_prepared(const int64_t *G, double tol) {
    double *inverse;int32_t rc;
    if(!G || G[0]<1 || G[0]>INT32_MAX || G[14]<1 || G[14]>200000001 || G[0]>G[14]/7) return -1;
    inverse=(double *)malloc((size_t)G[0]*sizeof(double));
    if(!inverse) return gmres_impl(G,tol,1,NULL);
    rc=gmres_impl(G,tol,1,inverse);
    free(inverse);return rc;
}
