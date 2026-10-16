/* soilfem_mat_x64.dll - coarse-grained material-point update, internal force
 * and element tangent for 2D Soil FEM (20261017_NATIVE_03).
 *
 * Bit-identity with VBA x64:
 *  - Arithmetic runs on the x87 FPU (-mfpmath=387) under the caller's control
 *    word, like VBA: intermediates stay in registers and are rounded when
 *    assigned to a double (-fexcess-precision=standard), passed by value or
 *    returned.
 *  - Sin/Cos/Atn/Sqr/^ call the MSVCR100.dll functions VBE7.DLL itself imports
 *    (calibrated: 0 mismatches in 20000 random arguments each).
 *  - VBA runtime errors (overflow, division by zero, invalid) are detected
 *    through the x87 and MXCSR sticky flags; such an element is declined
 *    untouched and VBA recomputes it, so VBA raises the same error.
 *  - Elements are processed in order and stop at the first element the DLL
 *    does not handle, so the internal-force sums keep VBA's order.
 * No CRT: the DLL never touches the process FPU control word.
 */
#include <stdint.h>
#include <windows.h>

typedef double (__cdecl *fn1)(double);
typedef double (__cdecl *fn2)(double, double);
static fn1 c_sin, c_cos, c_atan, c_sqrt;
static fn2 c_pow;
static int g_ready;

void *memset(void *d, int c, size_t n) { unsigned char *p = d; while (n--) *p++ = (unsigned char)c; return d; }
void *memcpy(void *d, const void *s, size_t n) { unsigned char *p = d; const unsigned char *q = s; while (n--) *p++ = *q++; return d; }

BOOL WINAPI DllMain(HINSTANCE h, DWORD r, LPVOID x) { (void)h; (void)r; (void)x; return TRUE; }

__declspec(dllexport) int32_t __stdcall sfm_init(void) {
  HMODULE m = GetModuleHandleA("MSVCR100.dll");
  if (!m) return -1; /* VBA's runtime not loaded: do not guess another CRT */
  c_sin = (fn1)(void (*)(void))GetProcAddress(m, "sin"); c_cos = (fn1)(void (*)(void))GetProcAddress(m, "cos");
  c_atan = (fn1)(void (*)(void))GetProcAddress(m, "atan"); c_sqrt = (fn1)(void (*)(void))GetProcAddress(m, "sqrt");
  c_pow = (fn2)(void (*)(void))GetProcAddress(m, "pow");
  if (!c_sin || !c_cos || !c_atan || !c_sqrt || !c_pow) return -2;
  g_ready = 1;
  return 3; /* API version */
}

/* ---------- FPU flag helpers ---------- */
static inline void fp_clear(void) {
  unsigned int mx; __asm__ volatile("fnclex");
  __asm__ volatile("stmxcsr %0" : "=m"(mx)); mx &= ~0x3Fu; __asm__ volatile("ldmxcsr %0" :: "m"(mx));
}
static inline int fp_bad(void) { /* IE | ZE | OE on x87 or SSE */
  unsigned short sw; unsigned int mx;
  __asm__ volatile("fnstsw %0" : "=m"(sw)); __asm__ volatile("stmxcsr %0" : "=m"(mx));
  return ((sw & 0x0D) != 0) || ((mx & 0x0D) != 0);
}

typedef long double xd; /* x87 extended expression type */
#define ABS(v) __builtin_fabsl((xd)(v))
static inline int is_finite(double v) { if (v != v) return 0; if (ABS(v) > 1E+300) return 0; return 1; }

static const double PI_VBA = 3.14159265358979;
static const double HALFPI_VBA = 1.5707963267949;
#define P2_TOL 0.0000000001

static double p2atan2(double y, double x) {
  double r;
  if (x > 0.0) r = c_atan(y / x);
  else if (x < 0.0) { if (y >= 0.0) r = (xd)c_atan(y / x) + PI_VBA; else r = (xd)c_atan(y / x) - PI_VBA; }
  else if (y > 0.0) r = HALFPI_VBA;
  else if (y < 0.0) r = -HALFPI_VBA;
  else r = 0.0;
  return r;
}

static void sort3(double v1, double v2, double v3, double *mx, double *md, double *mn) {
  double v[3]; int i, j; double t;
  v[0] = v1; v[1] = v2; v[2] = v3;
  for (i = 0; i < 2; i++) for (j = i + 1; j < 3; j++) if (v[i] < v[j]) { t = v[i]; v[i] = v[j]; v[j] = t; }
  *mx = v[0]; *md = v[1]; *mn = v[2];
}

static int eval_yield(double sx, double sy, double sz, double txy, double sinPhi, double coh, double cosPhi,
                      double *yv, double *smax, double *smid, double *smin, double *ang) {
  double mean, dev, radius, p1, p2;
  mean = 0.5 * ((xd)sx + sy);
  dev = 0.5 * ((xd)sx - sy);
  radius = c_sqrt((xd)dev * dev + (xd)txy * txy);
  p1 = (xd)mean + radius; p2 = (xd)mean - radius;
  *ang = 0.5 * (xd)p2atan2(2.0 * (xd)txy, (xd)sx - sy);
  sort3(p1, p2, sz, smax, smid, smin);
  *yv = (xd)*smax - *smin - ((xd)*smax + *smin) * sinPhi - 2.0 * (xd)coh * cosPhi;
  if (!is_finite(*yv) || !is_finite(*smax) || !is_finite(*smid) || !is_finite(*smin) || !is_finite(*ang)) return 0;
  return 1;
}

static double vonmises(double a, double b, double c, double t) {
  return c_sqrt(0.5 * ((xd)c_pow((xd)a - b, 2.0) + c_pow((xd)b - c, 2.0) + c_pow((xd)c - a, 2.0)) + 3.0 * (xd)c_pow(t, 2.0));
}

typedef struct {
  double stress[4]; double plastic[4]; double tangent[3][3];
  double mult, yieldf, pr[3], angle; int yielded, elastic, tanReady;
} mp_out;

typedef struct { double d00, d01, d22, sinPhi, cosPhi, sinPsi, young, poisson, coh; } mp_mat;

/* P2FillSpectralTangent */
static int fill_tangent(double nF[3][2], double qF[3][2], double aM[2][2], int count, const int rank[3],
                        const double pT[3], const double pC[3], double tA, const mp_mat *m, mp_out *o) {
  double iH[2][2] = {{0, 0}, {0, 0}}, aP[3][3] = {{0}}, tD[3], rD[3], reD[3], phD[3] = {0, 0, 0}, tg[3][3];
  int i, j, f, g, col; double det, c, sn, dsx, dsy, dt, dsz, angD, gap;
  gap = (xd)pT[0] - pT[1];
  if (gap <= 0.000000000001 * (ABS(pT[0]) + ABS(pT[1]) + 1.0)) return 0;
  if (count == 1) iH[0][0] = 1.0 / (xd)aM[0][0];
  else {
    det = (xd)aM[0][0] * aM[1][1] - (xd)aM[0][1] * aM[1][0];
    iH[0][0] = (xd)aM[1][1] / det; iH[0][1] = -(xd)aM[0][1] / det;
    iH[1][0] = -(xd)aM[1][0] / det; iH[1][1] = (xd)aM[0][0] / det;
  }
  for (i = 0; i < 3; i++) for (j = 0; j < 3; j++) {
    if (i == j) aP[i][j] = 1.0;
    for (f = 0; f < count; f++) for (g = 0; g < count; g++)
      aP[i][j] = (xd)aP[i][j] - (xd)qF[i][f] * iH[f][g] * nF[j][g];
  }
  c = c_cos(tA); sn = c_sin(tA);
  for (col = 0; col < 3; col++) {
    dsx = 0.0; dsy = 0.0; dt = 0.0; dsz = 0.0;
    if (col == 0) { dsx = m->d00; dsy = m->d01; dsz = m->d01; }
    if (col == 1) { dsx = m->d01; dsy = m->d00; dsz = m->d01; }
    if (col == 2) dt = m->d22;
    tD[0] = (xd)c * c * dsx + (xd)sn * sn * dsy + 2.0 * (xd)c * sn * dt;
    tD[1] = (xd)sn * sn * dsx + (xd)c * c * dsy - 2.0 * (xd)c * sn * dt;
    tD[2] = dsz;
    for (i = 0; i < 3; i++) rD[i] = tD[rank[i]];
    for (i = 0; i < 3; i++) {
      reD[i] = 0.0;
      for (j = 0; j < 3; j++) reD[i] = (xd)reD[i] + (xd)aP[i][j] * rD[j];
      phD[rank[i]] = reD[i];
    }
    angD = ((xd)c * sn * ((xd)dsy - dsx) + ((xd)c * c - (xd)sn * sn) * dt) / gap;
    tg[0][col] = (xd)c * c * phD[0] + (xd)sn * sn * phD[1] + 2.0 * (xd)c * sn * ((xd)pC[1] - pC[0]) * angD;
    tg[1][col] = (xd)sn * sn * phD[0] + (xd)c * c * phD[1] + 2.0 * (xd)c * sn * ((xd)pC[0] - pC[1]) * angD;
    tg[2][col] = (xd)c * sn * ((xd)phD[0] - phD[1]) + ((xd)c * c - (xd)sn * sn) * ((xd)pC[0] - pC[1]) * angD;
    for (i = 0; i < 3; i++) if (!is_finite(tg[i][col])) return 0;
  }
  for (i = 0; i < 3; i++) for (j = 0; j < 3; j++) o->tangent[i][j] = tg[i][j];
  o->tanReady = 1;
  return 1;
}

/* P2TrySpectralReturn restricted to its regular branch.
 * return 1 = success with algorithmic tangent, 0 = decline element */
static int spectral(double tSx, double tSy, double tSz, double tTxy, const mp_mat *m, double tA, mp_out *o) {
  double pT[3], pTr[3], cor[3], pC[3], nF[3][2], mF[3][2], qF[3][2], aM[2][2], fT[2], mu[2];
  int rank[3], i, j, f, g, cand, count, sw;
  double radius, center, det, sTol, lTol, sScale, cA, sA, fY, vmax, vmid, vmin, fAng;
  double ex, ey, ez, gam, sX, sY, sZ, sXY;
  center = 0.5 * ((xd)tSx + tSy);
  radius = c_sqrt((xd)c_pow(0.5 * ((xd)tSx - tSy), 2.0) + c_pow(tTxy, 2.0));
  pT[0] = (xd)center + radius; pT[1] = (xd)center - radius; pT[2] = tSz;
  for (i = 0; i < 3; i++) rank[i] = i;
  for (i = 0; i < 2; i++) for (j = i + 1; j < 3; j++)
    if (pT[rank[j]] > pT[rank[i]]) { sw = rank[i]; rank[i] = rank[j]; rank[j] = sw; }
  for (i = 0; i < 3; i++) pTr[i] = pT[rank[i]];
  sScale = ABS(pTr[0]) + ABS(pTr[2]) + 2.0 * (xd)m->coh * m->cosPhi + 1.0;
  sTol = (xd)P2_TOL * sScale * 10.0;
  lTol = 0.000000000001 + (xd)P2_TOL * sScale / m->young;
  cA = c_cos(tA); sA = c_sin(tA);
  for (cand = 1; cand <= 3; cand++) {
    memset(nF, 0, sizeof nF); memset(mF, 0, sizeof mF); memset(qF, 0, sizeof qF);
    memset(aM, 0, sizeof aM); fT[0] = fT[1] = 0.0; mu[0] = mu[1] = 0.0;
    count = (cand > 1) ? 2 : 1;
    nF[0][0] = 1.0 - (xd)m->sinPhi; nF[2][0] = -1.0 - (xd)m->sinPhi;
    mF[0][0] = 1.0 - (xd)m->sinPsi; mF[2][0] = -1.0 - (xd)m->sinPsi;
    if (cand == 2) {
      nF[1][1] = 1.0 - (xd)m->sinPhi; nF[2][1] = -1.0 - (xd)m->sinPhi;
      mF[1][1] = 1.0 - (xd)m->sinPsi; mF[2][1] = -1.0 - (xd)m->sinPsi;
    } else if (cand == 3) {
      nF[0][1] = 1.0 - (xd)m->sinPhi; nF[1][1] = -1.0 - (xd)m->sinPhi;
      mF[0][1] = 1.0 - (xd)m->sinPsi; mF[1][1] = -1.0 - (xd)m->sinPsi;
    }
    for (f = 0; f < count; f++) {
      fT[f] = -2.0 * (xd)m->coh * m->cosPhi;
      for (i = 0; i < 3; i++) {
        fT[f] = (xd)fT[f] + (xd)nF[i][f] * pTr[i];
        for (j = 0; j < 3; j++) {
          if (i == j) qF[i][f] = (xd)qF[i][f] + (xd)m->d00 * mF[j][f];
          else qF[i][f] = (xd)qF[i][f] + (xd)m->d01 * mF[j][f];
        }
      }
    }
    for (f = 0; f < count; f++) for (g = 0; g < count; g++)
      for (i = 0; i < 3; i++) aM[f][g] = (xd)aM[f][g] + (xd)nF[i][f] * qF[i][g];
    if (count == 1) {
      if (ABS(aM[0][0]) <= 1E-30) continue;
      mu[0] = (xd)fT[0] / aM[0][0];
    } else {
      det = (xd)aM[0][0] * aM[1][1] - (xd)aM[0][1] * aM[1][0];
      if (ABS(det) <= 0.00000000000001 * (ABS((xd)aM[0][0] * aM[1][1]) + ABS((xd)aM[0][1] * aM[1][0]) + 1.0)) continue;
      mu[0] = ((xd)fT[0] * aM[1][1] - (xd)fT[1] * aM[0][1]) / det;
      mu[1] = ((xd)fT[1] * aM[0][0] - (xd)fT[0] * aM[1][0]) / det;
    }
    { int bad = 0;
      for (f = 0; f < count; f++) {
        if (!is_finite(mu[f]) || mu[f] < -(xd)lTol) { bad = 1; break; }
        if (mu[f] < 0.0) mu[f] = 0.0;
      }
      if (bad) continue; }
    { int bad = 0;
      for (i = 0; i < 3; i++) {
        cor[i] = pTr[i];
        for (f = 0; f < count; f++) cor[i] = (xd)cor[i] - (xd)mu[f] * qF[i][f];
        if (!is_finite(cor[i])) { bad = 1; break; }
      }
      if (bad) continue; }
    if (cor[0] < (xd)cor[1] - sTol || cor[1] < (xd)cor[2] - sTol) continue;
    { int bad = 0;
      for (f = 0; f < count; f++) {
        fY = -2.0 * (xd)m->coh * m->cosPhi;
        for (i = 0; i < 3; i++) fY = (xd)fY + (xd)nF[i][f] * cor[i];
        if (ABS(fY) > sTol) { bad = 1; break; }
      }
      if (bad) continue; }
    for (i = 0; i < 3; i++) pC[rank[i]] = cor[i];
    sX = (xd)pC[0] * c_pow(cA, 2.0) + (xd)pC[1] * c_pow(sA, 2.0);
    sY = (xd)pC[0] * c_pow(sA, 2.0) + (xd)pC[1] * c_pow(cA, 2.0);
    sXY = ((xd)pC[0] - pC[1]) * cA * sA;
    sZ = pC[2];
    if (!eval_yield(sX, sY, sZ, sXY, m->sinPhi, m->coh, m->cosPhi, &fY, &vmax, &vmid, &vmin, &fAng)) continue;
    if (ABS(fY) > sTol) continue;
    if (m->sinPhi > 0.000000000001 && ABS((xd)vmax - vmin) <= sTol) return 0; /* vertex: VBA */
    { /* P2ComplianceFromStressDelta(young, poisson, d22, ...) */
      double dSx = (xd)tSx - sX, dSy = (xd)tSy - sY, dSz = (xd)tSz - sZ, dT = (xd)tTxy - sXY;
      ex = ((xd)dSx - (xd)m->poisson * ((xd)dSy + dSz)) / m->young;
      ey = ((xd)dSy - (xd)m->poisson * ((xd)dSx + dSz)) / m->young;
      ez = ((xd)dSz - (xd)m->poisson * ((xd)dSx + dSy)) / m->young;
      if (ABS(m->d22) <= 1E-30) gam = 0.0; else gam = (xd)dT / m->d22;
    }
    if (ABS(m->sinPsi) <= 0.000000000001) {
      if (ABS((xd)ex + ey + ez) > 0.0000000001 * (ABS(ex) + ABS(ey) + ABS(ez) + ABS(gam) + 1.0)) continue;
    }
    o->stress[0] = sX; o->stress[1] = sY; o->stress[2] = sXY; o->stress[3] = sZ;
    o->plastic[0] = ex; o->plastic[1] = ey; o->plastic[2] = gam; o->plastic[3] = ez;
    o->mult = (xd)mu[0] + mu[1];
    o->pr[0] = vmax; o->pr[1] = vmid; o->pr[2] = vmin;
    o->angle = (xd)fAng * 180.0 / PI_VBA;
    o->yieldf = fY; o->yielded = 1; o->elastic = 0;
    fill_tangent(nF, qF, aM, count, rank, pT, pC, tA, m, o);
    return o->tanReady ? 1 : 0; /* without the algorithmic tangent VBA uses FD: decline */
  }
  return 0; /* no regular candidate: vertex return or failure, VBA */
}

/* P2MaterialPointUpdate with PreferredSubsteps = 1 (cached constants).
 * return 1 handled (elastic or regular spectral), 0 decline */
static int mp_update(const double prev[4], const double de[3], const mp_mat *m, mp_out *o, int *elasticFast) {
  double tS[4], yv, smax, smid, smin, ang, ys; int i;
  for (i = 0; i < 4; i++) if (!is_finite(prev[i])) return 0;
  for (i = 0; i < 3; i++) if (!is_finite(de[i])) return 0;
  memset(o, 0, sizeof *o);
  tS[0] = (xd)prev[0] + (xd)m->d00 * de[0] + (xd)m->d01 * de[1];
  tS[1] = (xd)prev[1] + (xd)m->d01 * de[0] + (xd)m->d00 * de[1];
  tS[2] = (xd)prev[2] + (xd)m->d22 * de[2];
  tS[3] = (xd)prev[3] + (xd)m->d01 * ((xd)de[0] + de[1]);
  if (!eval_yield(tS[0], tS[1], tS[3], tS[2], m->sinPhi, m->coh, m->cosPhi, &yv, &smax, &smid, &smin, &ang)) return 0;
  ys = ABS(smax) + ABS(smin) + 2.0 * (xd)m->coh * m->cosPhi;
  if (ys < 1.0) ys = 1.0;
  if (yv <= (xd)P2_TOL * ys) {
    for (i = 0; i < 4; i++) o->stress[i] = tS[i];
    o->tangent[0][0] = m->d00; o->tangent[0][1] = m->d01; o->tangent[0][2] = 0.0;
    o->tangent[1][0] = m->d01; o->tangent[1][1] = m->d00; o->tangent[1][2] = 0.0;
    o->tangent[2][0] = 0.0; o->tangent[2][1] = 0.0; o->tangent[2][2] = m->d22;
    o->pr[0] = smax; o->pr[1] = smid; o->pr[2] = smin;
    o->angle = (xd)ang * 180.0 / PI_VBA;
    o->yieldf = yv; o->yielded = 0; o->elastic = 1;
    *elasticFast = 1;
    return 1;
  }
  /* Core (stepCount = 1, de/1 = de): same trial and yield, then spectral return */
  {
    double cS[4], cyv, csmax, csmid, csmin, cang, cys; mp_out so;
    memset(&so, 0, sizeof so);
    cS[0] = (xd)prev[0] + (xd)m->d00 * de[0] + (xd)m->d01 * de[1];
    cS[1] = (xd)prev[1] + (xd)m->d01 * de[0] + (xd)m->d00 * de[1];
    cS[2] = (xd)prev[2] + (xd)m->d22 * de[2];
    cS[3] = (xd)prev[3] + (xd)m->d01 * ((xd)de[0] + de[1]);
    if (!eval_yield(cS[0], cS[1], cS[3], cS[2], m->sinPhi, m->coh, m->cosPhi, &cyv, &csmax, &csmid, &csmin, &cang)) return 0;
    cys = ABS(csmax) + ABS(csmin) + 2.0 * (xd)m->coh * m->cosPhi;
    if (cys < 1.0) cys = 1.0;
    if (cyv <= (xd)P2_TOL * cys) return 0; /* cannot happen (same expressions); stay safe */
    if (!spectral(cS[0], cS[1], cS[3], cS[2], m, cang, &so)) return 0;
    /* wrapper accumulation from zero: 0# + x (normalises -0) */
    for (i = 0; i < 4; i++) o->stress[i] = so.stress[i];
    for (i = 0; i < 4; i++) o->plastic[i] = 0.0 + (xd)so.plastic[i];
    o->mult = 0.0 + (xd)so.mult;
    o->yielded = so.yielded; o->elastic = !so.yielded;
    for (i = 0; i < 3; i++) { int j; for (j = 0; j < 3; j++) o->tangent[i][j] = so.tangent[i][j]; }
    o->pr[0] = so.pr[0]; o->pr[1] = so.pr[1]; o->pr[2] = so.pr[2];
    o->angle = so.angle; o->yieldf = so.yieldf; o->tanReady = 1;
    if (!so.yielded) return 0;
    *elasticFast = 0;
    return 1;
  }
}

/* ---------- element memory access (VBA Element_Data, column-major) ---------- */
enum { O_STRIDE, O_U, O_MATNO, O_ELNODE, O_BMAT, O_SMAT, O_SPMAT, O_STMAT, O_MSTMAT, O_PLSTRAIN, O_PLMULT,
       O_YIELDF, O_YIELDED, O_PRSTRESS, O_LASTSUB, O_STREAK, O_DPMAT, O_KMAT, O_DJ, O_ELMWEIGHT, O_HGCOUNT,
       O_HGSCALE, O_HGMODE, O_ISJOINT, O_SPVALID, O_TDIRTY, O_COUNT };
#define EP(base, off, T) ((T *)((base) + (off)))
#define BIX(i, j, im) ((i) + 3 * ((j) + 16 * (im)))
#define SIX(i, im) ((i) + 17 * (im))

/* params (int64 array, filled by VBA):
 * 0 elemBase 1 offsetsPtr 2 nElem 3 tdisp 4 incBase 5 inforce 6 activePtr(0=all) 7 activeLo 8 activeHi
 * 9 matTable 10 nMat 11 activePP 12 flags(bit0 zeroIncAllowed, bit1 trialInit, bit2 skipHgForce, bit3 elasticCorr, bit4 inco)
 * 13 outPtr 14 gaussN 15 maxDof (TDisp/incBase/iNForce upper bound)
 * out (int64): 0 materialCalls 1 bucket0 2 materialSuccess 3 elasticFast 4 stateChange 5 pointUpdate
 *              6 activeDelta 7 elementsDone 8 lastActive(-1) 9 declineReason 10 tangentElements */
#define MAT_STRIDE 16
enum { M_READY, M_D00, M_D01, M_D22, M_SINPHI, M_COSPHI, M_SINPSI, M_YOUNG, M_POISSON, M_COH, M_THICK, M_WEIGHT };

static int elem_active(const int64_t *P, int k) {
  const int16_t *a = (const int16_t *)(intptr_t)P[6];
  if (!a) return 1;
  if (k < P[7] || k > P[8]) return 1;
  return a[k - P[7]] != 0;
}

typedef struct { double st[17][4]; double sp[3][16][4]; double dp[3][3][4]; mp_out o[4]; int ef[4]; } elem_work;

/* returns the first element index NOT processed (nElem when finished) */
__declspec(dllexport) int32_t __stdcall sfm_eval(const int64_t *P, int32_t kStart) {
  char *base = (char *)(intptr_t)P[0]; const int64_t *off = (const int64_t *)(intptr_t)P[1];
  int nElem = (int)P[2]; const double *tdisp = (const double *)(intptr_t)P[3];
  const double *inc = (const double *)(intptr_t)P[4]; double *inf = (double *)(intptr_t)P[5];
  const double *mt = (const double *)(intptr_t)P[9]; int nMat = (int)P[10];
  int16_t *app = (int16_t *)(intptr_t)P[11]; int flags = (int)P[12]; int64_t *out = (int64_t *)(intptr_t)P[13];
  int zeroOK = flags & 1, trialInit = (flags >> 1) & 1, skipHg = (flags >> 2) & 1;
  int k, j, im, i, kk; static elem_work W;
  if (!g_ready) { out[9] = 1; return kStart; }
  for (k = kStart; k < nElem; k++) {
    char *e = base + off[O_STRIDE] * (int64_t)k;
    double *u = EP(e, off[O_U], double); const int32_t *eln = EP(e, off[O_ELNODE], int32_t);
    const double *B = EP(e, off[O_BMAT], double); double *St = EP(e, off[O_STMAT], double);
    const double *mSt = EP(e, off[O_MSTMAT], double); const double *dj = EP(e, off[O_DJ], double);
    int mat, zero; mp_mat m; double thick; int64_t dCalls = 0, dEF = 0, dSC = 0, dPU = 0, dAD = 0;
    double uloc[16];
    if (!elem_active(P, k)) continue;
    if (*EP(e, off[O_ISJOINT], int16_t)) { out[9] = 2; return k; }
    mat = *EP(e, off[O_MATNO], int32_t);
    if (mat < 0 || mat >= nMat || mt[mat * MAT_STRIDE + M_READY] == 0.0) { out[9] = 3; return k; }
    m.d00 = mt[mat * MAT_STRIDE + M_D00]; m.d01 = mt[mat * MAT_STRIDE + M_D01]; m.d22 = mt[mat * MAT_STRIDE + M_D22];
    m.sinPhi = mt[mat * MAT_STRIDE + M_SINPHI]; m.cosPhi = mt[mat * MAT_STRIDE + M_COSPHI]; m.sinPsi = mt[mat * MAT_STRIDE + M_SINPSI];
    m.young = mt[mat * MAT_STRIDE + M_YOUNG]; m.poisson = mt[mat * MAT_STRIDE + M_POISSON]; m.coh = mt[mat * MAT_STRIDE + M_COH];
    thick = mt[mat * MAT_STRIDE + M_THICK];
    if (trialInit) { const int32_t *ls = EP(e, off[O_LASTSUB], int32_t);
      for (im = 0; im < 4; im++) if (ls[im] > 1) { out[9] = 4; return k; } }
    for (j = 0; j < 16; j++) if (eln[j] > (int64_t)P[15]) { out[9] = 7; return k; }
    fp_clear();
    /* u(j) = TDisp - incrementBase; skipped dofs keep their stored value */
    zero = 1;
    for (j = 0; j < 16; j++) {
      uloc[j] = u[j];
      if (eln[j] < 0) continue;
      uloc[j] = (xd)tdisp[eln[j]] - inc[eln[j]];
      if (ABS(uloc[j]) > 1E-30) zero = 0;
    }
    if (zero && zeroOK) {
      for (im = 0; im < 4; im++) for (i = 0; i < 17; i++) W.st[i][im] = mSt[SIX(i, im)];
    } else {
      zero = 0;
      for (im = 0; im < 4; im++) {
        double prev[4], de[3]; mp_out *o = &W.o[im]; int ef = 0;
        prev[0] = mSt[SIX(0, im)]; prev[1] = mSt[SIX(1, im)]; prev[2] = mSt[SIX(2, im)]; prev[3] = mSt[SIX(16, im)];
        for (i = 0; i < 3; i++) { de[i] = 0.0; for (j = 0; j < 16; j++) de[i] = (xd)de[i] + (xd)B[BIX(i, j, im)] * uloc[j]; }
        if (!mp_update(prev, de, &m, o, &ef)) { out[9] = 5; fp_clear(); return k; }
        W.ef[im] = ef; dCalls++; if (ef) dEF++;
        for (i = 0; i < 17; i++) W.st[i][im] = St[SIX(i, im)];
        W.st[0][im] = o->stress[0]; W.st[1][im] = o->stress[1]; W.st[2][im] = o->stress[2];
        W.st[3][im] = o->pr[0]; W.st[4][im] = o->pr[2];
        W.st[5][im] = 0.5 * ((xd)o->pr[0] - o->pr[2]);
        W.st[6][im] = o->angle;
        W.st[7][im] = vonmises(o->stress[0], o->stress[1], o->stress[3], o->stress[2]);
        if (o->pr[2] > 0.0) W.st[8][im] = o->pr[0];
        else if (o->pr[0] > 0.0) W.st[8][im] = (xd)o->pr[0] - o->pr[2];
        else W.st[8][im] = -(xd)o->pr[2];
        W.st[9][im] = o->yieldf;
        W.st[16][im] = o->stress[3];
        for (i = 0; i < 3; i++) {
          for (j = 0; j < 3; j++) W.dp[i][j][im] = o->tangent[i][j];
          for (j = 0; j < 16; j++) { double s = 0.0; for (kk = 0; kk < 3; kk++) s = (xd)s + (xd)o->tangent[i][kk] * B[BIX(kk, j, im)]; W.sp[i][j][im] = s; }
        }
      }
    }
    /* CalcForce into a local copy of the touched rows, committed only if clean */
    { double fl[16]; int hgN = *EP(e, off[O_HGCOUNT], int32_t); double hgS = *EP(e, off[O_HGSCALE], double);
      const double *HM = EP(e, off[O_HGMODE], double); int md;
      double force[16]; int idx[16]; int nf = 0;
      (void)fl;
      for (j = 0; j < 16; j++) {
        int kj = eln[j]; double s;
        if (kj < 0) continue;
        s = 0.0;
        for (im = 0; im < 4; im++) for (i = 0; i < 3; i++)
          s = (xd)s + (xd)B[BIX(i, j, im)] * W.st[i][im] * dj[im] * thick;
        idx[nf] = kj; force[nf] = s; nf++;
      }
      if (fp_bad()) { out[9] = 6; fp_clear(); return k; }
      /* ---- commit (from here on the element is VBA-equivalent) ---- */
      for (j = 0; j < 16; j++) u[j] = uloc[j];
      for (im = 0; im < 4; im++) for (i = 0; i < 17; i++) St[SIX(i, im)] = W.st[i][im];
      for (j = 0; j < nf; j++) inf[idx[j]] = (xd)inf[idx[j]] + force[j];
      if (!skipHg && hgN > 0 && hgS > 0.0) {
        for (md = 0; md < hgN; md++) {
          double pv = 0.0;
          for (j = 0; j < 16; j++) if (eln[j] >= 0) pv = (xd)pv + (xd)HM[md + 4 * j] * tdisp[eln[j]];
          pv = (xd)hgS * pv;
          for (j = 0; j < 16; j++) if (eln[j] >= 0) inf[eln[j]] = (xd)inf[eln[j]] + (xd)pv * HM[md + 4 * j];
        }
      }
    }
    if (!zero) {
      double *Sp = EP(e, off[O_SPMAT], double), *Dp = EP(e, off[O_DPMAT], double);
      int16_t *spv = EP(e, off[O_SPVALID], int16_t);
      if (trialInit) {
        double *pls = EP(e, off[O_PLSTRAIN], double), *plm = EP(e, off[O_PLMULT], double);
        double *yf = EP(e, off[O_YIELDF], double), *prs = EP(e, off[O_PRSTRESS], double);
        int16_t *yl = EP(e, off[O_YIELDED], int16_t), *td = EP(e, off[O_TDIRTY], int16_t);
        int32_t *ls = EP(e, off[O_LASTSUB], int32_t), *sk = EP(e, off[O_STREAK], int32_t);
        for (im = 0; im < 4; im++) {
          mp_out *o = &W.o[im]; int y = o->yielded ? 1 : 0, wasY = yl[im] != 0;
          int16_t *ap = app + (im + 4 * (int64_t)k);
          for (i = 0; i < 4; i++) pls[i + 4 * im] = o->plastic[i];
          if (y != wasY) dSC++;
          if (y || wasY) *td = -1;
          if (!o->elastic) dPU++; /* PlasticOccurred = not elastic path */
          if (y && !*ap) dAD++;
          if (!y && *ap) dAD--;
          *ap = y ? -1 : 0;
          plm[im] = o->mult; yf[im] = o->yieldf; yl[im] = y ? -1 : 0;
          for (i = 0; i < 3; i++) prs[i + 3 * im] = o->pr[i];
          if (o->elastic) { /* P3DecaySubstepHint with lastHint <= 1 */
            ls[im] = 1; sk[im] = 0;
          } else { /* P3RememberSubstepHint(1) with lastHint <= 1 */
            if (ls[im] < 1) { /* lastHint treated as 1, field unchanged unless reset below */ }
            sk[im] = sk[im] + 1;
            if (sk[im] >= 2) sk[im] = 0;
          }
        }
      }
      for (im = 0; im < 4; im++) {
        for (i = 0; i < 3; i++) for (j = 0; j < 16; j++) Sp[BIX(i, j, im)] = W.sp[i][j][im];
        spv[im] = -1;
      }
      for (im = 0; im < 4; im++) for (i = 0; i < 3; i++) for (j = 0; j < 3; j++) Dp[i + 3 * (j + 3 * im)] = W.dp[i][j][im];
    }
    out[0] += dCalls; out[1] += dCalls; out[2] += dCalls; out[3] += dEF;
    out[4] += dSC; out[5] += dPU; out[6] += dAD; out[7] += 1; out[8] = k;
    fp_clear();
  }
  out[9] = 0;
  return nElem;
}

/* Element tangents (SpmatValid repair + SetElmStiffness) for every active,
 * non-joint element with a valid material that is dirty or INCONSISTENT.
 * Writes kmat, ElmWeight, repaired Spmat; flags 'done' per element (int16 array).
 * returns 0 ok, else 1 (VBA rebuilds everything itself). */
__declspec(dllexport) int32_t __stdcall sfm_tangent(const int64_t *P, int16_t *done) {
  char *base = (char *)(intptr_t)P[0]; const int64_t *off = (const int64_t *)(intptr_t)P[1];
  int nElem = (int)P[2]; const double *mt = (const double *)(intptr_t)P[9]; int nMat = (int)P[10];
  int flags = (int)P[12]; int64_t *out = (int64_t *)(intptr_t)P[13]; const double *gN = (const double *)(intptr_t)P[14];
  int elasticCorr = (flags >> 3) & 1, inco = (flags >> 4) & 1;
  int k, im, i, j, kk; int64_t cnt = 0;
  if (!g_ready) return 1;
  fp_clear();
  for (k = 0; k < nElem; k++) {
    char *e = base + off[O_STRIDE] * (int64_t)k; int mat; double thick, weight, *K, *ew;
    const double *B, *S, *dj; double *Sp; const double *Sm; const int16_t *spv;
    done[k] = 0;
    if (!elem_active(P, k)) continue;
    if (*EP(e, off[O_ISJOINT], int16_t)) continue;
    if (!*EP(e, off[O_TDIRTY], int16_t) && !inco) continue;
    mat = *EP(e, off[O_MATNO], int32_t);
    if (mat < 0 || mat >= nMat) continue;
    thick = mt[mat * MAT_STRIDE + M_THICK]; weight = mt[mat * MAT_STRIDE + M_WEIGHT];
    Sp = EP(e, off[O_SPMAT], double); Sm = EP(e, off[O_SMAT], double); spv = EP(e, off[O_SPVALID], int16_t);
    for (im = 0; im < 4; im++) if (!spv[im]) for (i = 0; i < 3; i++) for (j = 0; j < 16; j++) Sp[BIX(i, j, im)] = Sm[BIX(i, j, im)];
    B = EP(e, off[O_BMAT], double); S = elasticCorr ? Sm : Sp; dj = EP(e, off[O_DJ], double);
    K = EP(e, off[O_KMAT], double); ew = EP(e, off[O_ELMWEIGHT], double);
    for (i = 0; i < 256; i++) K[i] = 0.0;
    *ew = 0.0;
    for (im = 0; im < 4; im++) for (i = 0; i < 16; i++) {
      for (j = 0; j < 16; j++) {
        double s = 0.0;
        for (kk = 0; kk < 3; kk++) s = (xd)s + (xd)B[BIX(kk, i, im)] * S[BIX(kk, j, im)];
        K[i + 16 * j] = (xd)K[i + 16 * j] + (xd)s * dj[im] * thick;
      }
      if (i <= 7) *ew = (xd)*ew + (xd)weight * thick * gN[im + 4 * i] * dj[im];
    }
    done[k] = -1; cnt++;
  }
  out[10] = cnt;
  if (fp_bad()) { fp_clear(); return 1; }
  fp_clear();
  return 0;
}
