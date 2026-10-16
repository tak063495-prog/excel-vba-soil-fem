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
  return 4; /* API version */
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

/* =====================================================================
 * Phase C kernels (API 4): assembly, CSR/GMRES/ILU(0), regularized
 * correction, residual metrics, RCM. Same x87 discipline as above: every
 * VBA Double variable/array element is a C double (rounded on store),
 * every expression is evaluated in xd. Loop order follows FEMSolver.bas.
 * ===================================================================== */
static inline double vsqr(double v) { return c_sqrt(v); }
static inline int vfin(double v) { return is_finite(v); }

/* SetTotalMat band scatter: band(row, col) += Elem(el).kmat(li, lj) */
__declspec(dllexport) int32_t __stdcall sfm_band_scatter(int32_t n, const int32_t *row, const int32_t *col,
    const int32_t *el, const int32_t *li, const int32_t *lj, double *band, int64_t ld,
    const char *elemBase, int64_t stride, int64_t kmatOff) {
  int32_t s;
  if (!band || n < 0) return 1;
  for (s = 0; s < n; s++) {
    const double *K = (const double *)(elemBase + stride * (int64_t)el[s] + kmatOff);
    double *a = &band[row[s] + (int64_t)col[s] * ld];
    *a = (xd)*a + K[li[s] + 16 * lj[s]];
  }
  return 0;
}

/* P6AssembleCSRValuesDetailed */
__declspec(dllexport) int32_t __stdcall sfm_csr_scatter(const int64_t *P, int32_t nnz, const int32_t *scStart,
    const int32_t *scPos, const int32_t *scRow, const int32_t *scCol, double *vals, int64_t kmatOff) {
  const char *base = (const char *)(intptr_t)P[0]; const int64_t *off = (const int64_t *)(intptr_t)P[1];
  int nElem = (int)P[2]; int k, p, s;
  for (p = 0; p < nnz; p++) vals[p] = 0.0;
  for (k = 0; k < nElem; k++) {
    const double *K;
    if (!elem_active(P, k)) continue;
    K = (const double *)(base + off[O_STRIDE] * (int64_t)k + kmatOff);
    for (s = scStart[k]; s < scStart[k + 1]; s++) {
      p = scPos[s];
      if (p < 0 || p >= nnz) return 1;
      vals[p] = (xd)vals[p] + K[scRow[s] + 16 * scCol[s]];
    }
  }
  return 0;
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
__declspec(dllexport) int32_t __stdcall sfm_csr_resnorm(int32_t n, const int32_t *rp, const int32_t *ci, const double *v,
    const double *x, const double *b, double *r, double *norm) {
  if (!g_ready) return 1;
  *norm = csr_resnorm(n, rp, ci, v, x, b, r);
  return 0;
}

/* P6BuildCSRILU. out: [0]=failRow [1]=failCol ; minAbs. returns 1 ready, 0 failed */
__declspec(dllexport) int32_t __stdcall sfm_ilu0(int32_t n, int32_t nnz, const int32_t *rp, const int32_t *ci, const int32_t *dp,
    const double *vals, double *ilu, int32_t *fail, double *minAbs) {
  int32_t r, p, kp, u, c, kc; double dk, m;
  fail[0] = -1; fail[1] = -1; *minAbs = 1E+308;
  if (nnz <= 0) return 0;
  for (p = 0; p < nnz; p++) ilu[p] = vals[p];
  for (r = 0; r < n; r++) {
    for (p = rp[r]; p < rp[r + 1]; p++) {
      c = ci[p];
      if (c < r) {
        dk = ilu[dp[c]];
        if (!vfin(dk) || ABS(dk) <= 0.000000000001) { fail[0] = r; fail[1] = c; *minAbs = vfin(dk) ? (double)ABS(dk) : 0.0; return 0; }
        m = (xd)ilu[p] / dk;
        if (!vfin(m)) { fail[0] = r; fail[1] = c; *minAbs = ABS(dk); return 0; }
        ilu[p] = m;
        for (kp = rp[c]; kp < rp[c + 1]; kp++) {
          kc = ci[kp];
          if (kc > c) {
            int32_t q; u = -1;
            for (q = rp[r]; q < rp[r + 1]; q++) if (ci[q] == kc) { u = q; break; }
            if (u >= 0) ilu[u] = (xd)ilu[u] - (xd)m * ilu[kp];
          }
        }
      }
    }
    dk = ilu[dp[r]];
    if (!vfin(dk) || ABS(dk) <= 0.000000000001) { fail[0] = r; fail[1] = r; *minAbs = vfin(dk) ? (double)ABS(dk) : 0.0; return 0; }
    if (ABS(dk) < *minAbs) *minAbs = ABS(dk);
  }
  if (*minAbs == 1E+308) *minAbs = 0.0;
  return 1;
}

static int ilu_apply(int32_t n, const int32_t *rp, const int32_t *ci, const double *ilu, const double *in, double *out) {
  int32_t r, p, c; double rv, dv;
  for (r = 0; r < n; r++) {
    rv = in[r];
    for (p = rp[r]; p < rp[r + 1]; p++) { c = ci[p]; if (c < r) rv = (xd)rv - (xd)ilu[p] * out[c]; }
    out[r] = rv;
  }
  for (r = n - 1; r >= 0; r--) {
    rv = out[r]; dv = 0.0;
    for (p = rp[r]; p < rp[r + 1]; p++) { c = ci[p]; if (c == r) dv = ilu[p]; if (c > r) rv = (xd)rv - (xd)ilu[p] * out[c]; }
    if (ABS(dv) <= 0.000000000001) return 0;
    out[r] = (xd)rv / dv;
  }
  return 1;
}

/* Right-preconditioned restarted GMRES with ILU(0) = P6SolveCSRGMRES (ILU branch).
 * G (int64): 0 n 1 rowptr 2 col 3 vals 4 ilu 5 rhs(Force) 6 x out(Disp) 7 work 8 restart 9 maxIter
 *            10 events(double, 3 per event) 11 eventCap 12 out(double): [0]lastRes [1]iters [2]nEvents [3]overflow
 * dtol: P6IterativeTolerance. returns 1 success (x written), 0 failure. */
enum { EV_RESTART = 1, EV_NONFIN_TRUE = 2, EV_PREC = 3, EV_NONFIN_KRYLOV = 4, EV_ZERO_GIVENS = 5, EV_ZERO_UPPER = 6 };
static void ev_push(double *ev, int64_t cap, double *o, int t, double a, double b) {
  int64_t k = (int64_t)o[2];
  if (k >= cap) { o[3] = 1.0; return; }
  ev[3 * k] = t; ev[3 * k + 1] = a; ev[3 * k + 2] = b; o[2] = (double)(k + 1);
}
__declspec(dllexport) int32_t __stdcall sfm_gmres(const int64_t *G, double dtol) {
  int32_t n = (int32_t)G[0]; const int32_t *rp = (const int32_t *)(intptr_t)G[1], *ci = (const int32_t *)(intptr_t)G[2];
  const double *A = (const double *)(intptr_t)G[3], *ilu = (const double *)(intptr_t)G[4], *F = (const double *)(intptr_t)G[5];
  double *X = (double *)(intptr_t)G[6], *wk = (double *)(intptr_t)G[7]; int32_t m = (int32_t)G[8], maxIt = (int32_t)G[9];
  double *ev = (double *)(intptr_t)G[10]; int64_t cap = G[11]; double *o = (double *)(intptr_t)G[12];
  double *V, *Z, *H, *cs, *sn, *g, *y, *b, *x, *r, *w;
  int32_t i, j, k, pass, used, total = 0; int64_t ld = n, hl = (int64_t)m + 1;
  double normB = 0.0, beta, nextNorm, inner, rot, tmp, residual, tol;
  o[0] = 1E+100; o[1] = 0; o[2] = 0; o[3] = 0;
  if (!g_ready || n < 1 || m < 1) return -1;
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
      if (!ilu_apply(n, rp, ci, ilu, V + ld * j, zj)) { ev_push(ev, cap, o, EV_PREC, total, 0); return 0; }
      csr_matvec(n, rp, ci, A, zj, w);
      for (k = 0; k <= m; k++) H[k + hl * j] = 0.0;
      for (pass = 1; pass <= 2; pass++) {
        for (k = 0; k <= j; k++) {
          const double *vk = V + ld * k;
          inner = 0.0;
          for (i = 0; i < n; i++) inner = (xd)inner + (xd)w[i] * vk[i];
          H[k + hl * j] = (xd)H[k + hl * j] + inner;
          for (i = 0; i < n; i++) w[i] = (xd)w[i] - (xd)inner * vk[i];
        }
      }
      { double ss = 0.0, cv; for (i = 0; i < n; i++) { cv = w[i]; ss = (xd)ss + (xd)cv * cv; } nextNorm = (ss <= 0.0) ? 0.0 : vsqr(ss); }
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
    for (k = 0; k <= used - 1; k++) { const double *zk = Z + ld * k; for (i = 0; i < n; i++) x[i] = (xd)x[i] + (xd)zk[i] * y[k]; }
  }
  residual = csr_resnorm(n, rp, ci, A, x, b, r);
  if (!vfin(residual)) return 0;
  o[0] = (xd)residual / normB;
  if (residual > tol) return 0;
  for (i = 0; i < n; i++) if (!vfin(x[i])) return 0;
  for (i = 0; i < n; i++) X[i] = x[i];
  return 1;
}

/* P6SolveResidualRecovery numeric core (column-scaled LM normal equations, band Cholesky).
 * R (int64): 0 n 1 rowptr 2 col 3 origVals 4 nodeCond(int32) 5 force 6 disp out 7 bandWork (n*(bw+1)) 8 bw
 *            9 vecWork (4n) 10 out(double) [0]normal_relres [1]factorDone
 * returns 1 accepted (Disp written), 0 rejected */
__declspec(dllexport) int32_t __stdcall sfm_regularized(const int64_t *R, double damping, double itTol) {
  int32_t n = (int32_t)R[0]; const int32_t *rp = (const int32_t *)(intptr_t)R[1], *ci = (const int32_t *)(intptr_t)R[2];
  const double *O = (const double *)(intptr_t)R[3]; const int32_t *nc = (const int32_t *)(intptr_t)R[4];
  const double *F = (const double *)(intptr_t)R[5]; double *D = (double *)(intptr_t)R[6], *a = (double *)(intptr_t)R[7];
  int32_t bw = (int32_t)R[8]; double *vw = (double *)(intptr_t)R[9], *o = (double *)(intptr_t)R[10];
  double *rhs = vw, *sol = vw + n, *cs = vw + 2 * n; int64_t ld = n;
  int32_t row, col, p, q, c1, c2, i, j, k, fc, lr; double v, normB = 0.0, nres = 0.0, tmp;
#define A_(r, c) a[(r) + (int64_t)(c) * ld]
  o[0] = 0; o[1] = 0;
  if (!g_ready || n < 1 || bw < 0) return -1;
  fp_clear();
  for (i = 0; i < n; i++) { rhs[i] = 0.0; sol[i] = 0.0; cs[i] = 0.0; }
  for (i = 0; i < n * (int64_t)(bw + 1); i++) a[i] = 0.0;
  for (row = 0; row < n; row++) if (nc[row] == 0) {
    if (!vfin(F[row])) return 0;
    for (p = rp[row]; p < rp[row + 1]; p++) { col = ci[p]; if (nc[col] == 0) { v = O[p]; if (!vfin(v)) return 0; cs[col] = (xd)cs[col] + (xd)v * v; } }
  }
  for (col = 0; col < n; col++) { cs[col] = vsqr(cs[col]); if (cs[col] < 1E-30) cs[col] = 1.0; }
  for (row = 0; row < n; row++) if (nc[row] == 0) {
    for (p = rp[row]; p < rp[row + 1]; p++) {
      c1 = ci[p];
      if (nc[c1] == 0 && O[p] != 0.0) {
        v = (xd)O[p] / cs[c1];
        rhs[c1] = (xd)rhs[c1] + (xd)v * F[row];
        for (q = p; q < rp[row + 1]; q++) {
          c2 = ci[q];
          if (nc[c2] == 0 && O[q] != 0.0) {
            if (c2 < c1 || c2 - c1 > bw) return 0;
            A_(c2, bw + c1 - c2) = (xd)A_(c2, bw + c1 - c2) + (xd)v * O[q] / cs[c2];
          }
        }
      }
    }
  }
  for (i = 0; i < n; i++) {
    A_(i, bw) = (xd)A_(i, bw) + damping;
    if (nc[i] != 0) A_(i, bw) = 1.0;
    normB = (xd)normB + (xd)rhs[i] * rhs[i];
  }
  if (!vfin(normB) || normB <= 1E-60) return 0;
  for (i = 0; i < n; i++) {
    fc = i - bw; if (fc < 0) fc = 0;
    for (j = fc; j <= i; j++) {
      tmp = A_(i, bw + j - i);
      for (k = fc; k <= j - 1; k++) if (j - k <= bw) tmp = (xd)tmp - (xd)A_(i, bw + k - i) * A_(j, bw + k - j);
      if (i == j) { if (tmp <= 0.0 || !vfin(tmp)) return 0; A_(i, bw) = vsqr(tmp); }
      else A_(i, bw + j - i) = (xd)tmp / A_(j, bw);
    }
    tmp = rhs[i];
    for (j = fc; j <= i - 1; j++) tmp = (xd)tmp - (xd)A_(i, bw + j - i) * sol[j];
    sol[i] = (xd)tmp / A_(i, bw);
  }
  o[1] = 1;
  for (i = n - 1; i >= 0; i--) {
    tmp = sol[i]; lr = i + bw; if (lr > n - 1) lr = n - 1;
    for (j = i + 1; j <= lr; j++) tmp = (xd)tmp - (xd)A_(j, bw + i - j) * sol[j];
    sol[i] = (xd)tmp / A_(i, bw);
    if (!vfin(sol[i])) return 0;
  }
  for (row = 0; row < n; row++) if (nc[row] == 0) {
    tmp = 0.0;
    for (p = rp[row]; p < rp[row + 1]; p++) { col = ci[p]; if (nc[col] == 0) tmp = (xd)tmp + (xd)O[p] * sol[col] / cs[col]; }
    for (p = rp[row]; p < rp[row + 1]; p++) { col = ci[p]; if (nc[col] == 0) rhs[col] = (xd)rhs[col] - (xd)O[p] * tmp / cs[col]; }
  }
  for (i = 0; i < n; i++) { tmp = (xd)rhs[i] - (xd)damping * sol[i]; nres = (xd)nres + (xd)tmp * tmp; }
  if (!vfin(nres)) return 0;
  { double q2 = (xd)nres / normB; o[0] = vsqr(q2); }
  if (o[0] > itTol) return 0;
  for (i = 0; i < n; i++) { D[i] = (xd)sol[i] / cs[i]; if (nc[i] != 0) D[i] = 0.0; }
  fp_clear();
  return 1;
#undef A_
}

/* P3AccumulateResidualMetrics (non-mixed). M (int64): 0 n 1 internal 2 target 3 nodeCond 4 tdisp 5 reaction(0=lite)
 * 6 boundaryDisp 7 stageStartDisp(0=not ready) 8 out(double)
 * in: forceRef. out: 0 ResFull 1 ResFree 2 ForceFree 3 MaxAbsResFree 4 InternalFree 5 maxAbsDisp 6 dispForceSq
 *     7 reactX 8 reactY 9 energyValue */
__declspec(dllexport) int32_t __stdcall sfm_residual(const int64_t *M, double forceRef) {
  int32_t n = (int32_t)M[0], i; const double *in = (const double *)(intptr_t)M[1], *tg = (const double *)(intptr_t)M[2];
  const int32_t *nc = (const int32_t *)(intptr_t)M[3]; const double *td = (const double *)(intptr_t)M[4];
  double *re = (double *)(intptr_t)M[5]; const double *bd = (const double *)(intptr_t)M[6], *sd = (const double *)(intptr_t)M[7];
  double *o = (double *)(intptr_t)M[8]; int lite = (re == 0);
  double rf = 0, rs = 0, fs = 0, is = 0, mr = 0, md = 0, dfs = 0, rx = 0, ry = 0, en = 0, rv;
  if (!g_ready) return 1;
  for (i = 0; i < n; i++) {
    rv = (xd)in[i] - tg[i];
    if (!lite) re[i] = rv;
    if (ABS(rv) > rf) rf = ABS(rv);
    if (nc[i] == 0) {
      rs = (xd)rs + (xd)rv * rv; fs = (xd)fs + (xd)tg[i] * tg[i]; is = (xd)is + (xd)in[i] * in[i];
      if (ABS(rv) > mr) mr = ABS(rv);
      if (ABS(td[i]) > md) md = ABS(td[i]);
    } else {
      if (forceRef <= 0.000000000001) {
        int dd = (ABS(bd[i]) > 1E-30);
        if (!dd && sd) dd = (ABS(sd[i]) > 1E-30);
        if (dd) dfs = (xd)dfs + (xd)in[i] * in[i];
      }
      if (!lite) { if (i % 2 == 0) rx = (xd)rx + rv; else ry = (xd)ry + rv; }
    }
    if (!lite) en = (xd)en + (xd)td[i] * rv;
  }
  o[0] = rf; o[1] = vsqr(rs); o[2] = vsqr(fs); o[3] = mr; o[4] = vsqr(is); o[5] = md; o[6] = dfs; o[7] = rx; o[8] = ry; o[9] = en;
  return 0;
}

/* ---------- RCM (P6BuildRCMOrdering + P6ComputeRCMCandidateBandwidth) ---------- */
static void *halloc(size_t n) { return HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, n ? n : 1); }
static void hfree(void *p) { if (p) HeapFree(GetProcessHeap(), 0, p); }
/* nodes: nElem*8 (Elem.node), elnode: nElem*16 (Elem.ElNode). returns 0 ok, 1 range error, 2 alloc, 3 count */
__declspec(dllexport) int32_t __stdcall sfm_rcm(int32_t nNode, int32_t nElem, const int32_t *nodes, const int32_t *elnode,
    int32_t *oldToNew, int32_t *newToOld, int32_t *candBw) {
  int32_t *deg = 0, *cap = 0, **adj = 0, *queue = 0, *cm = 0, *nb = 0; char *vis = 0;
  int32_t e, a, b2, i, j, rc = 0, cmCount = 0, maxDeg = 0;
  if (nNode <= 0 || nElem <= 0) return 3;
  deg = halloc(sizeof(int32_t) * nNode); cap = halloc(sizeof(int32_t) * nNode); adj = halloc(sizeof(int32_t *) * nNode);
  queue = halloc(sizeof(int32_t) * nNode); cm = halloc(sizeof(int32_t) * nNode); vis = halloc(nNode);
  if (!deg || !cap || !adj || !queue || !cm || !vis) { rc = 2; goto done; }
#define ADDN(x, y) do { int32_t _t; for (_t = 0; _t < deg[x]; _t++) if (adj[x][_t] == (y)) break; \
    if (_t == deg[x]) { if (deg[x] == cap[x]) { int32_t nc2 = cap[x] ? cap[x] * 2 : 16; int32_t *np = halloc(sizeof(int32_t) * nc2); \
      if (!np) { rc = 2; goto done; } for (_t = 0; _t < deg[x]; _t++) np[_t] = adj[x][_t]; hfree(adj[x]); adj[x] = np; cap[x] = nc2; } \
      adj[x][deg[x]++] = (y); } } while (0)
  for (e = 0; e < nElem; e++) for (i = 0; i < 8; i++) {
    a = nodes[e * 8 + i]; if (a < 0) continue;
    for (j = i + 1; j < 8; j++) {
      b2 = nodes[e * 8 + j]; if (b2 < 0) continue;
      if (b2 >= nNode || a >= nNode) { rc = 1; goto done; }
      ADDN(a, b2); ADDN(b2, a);
    }
  }
  for (i = 0; i < nNode; i++) if (deg[i] > maxDeg) maxDeg = deg[i];
  nb = halloc(sizeof(int32_t) * (maxDeg + 1)); if (!nb) { rc = 2; goto done; }
  while (cmCount < nNode) {
    int32_t start = -1, sd = nNode + 1, head = 0, tail = 0;
    for (i = 0; i < nNode; i++) if (!vis[i] && deg[i] < sd) { start = i; sd = deg[i]; }
    if (start < 0) break;
    queue[0] = start; vis[start] = 1;
    while (head <= tail) {
      int32_t cur = queue[head++], cnt, o2, in2;
      cm[cmCount++] = cur;
      cnt = deg[cur];
      for (i = 0; i < cnt; i++) nb[i] = adj[cur][i];
      for (o2 = 1; o2 < cnt; o2++) { /* P6SortRCMNeighbors */
        int32_t key = nb[o2]; in2 = o2 - 1;
        while (in2 >= 0) { int32_t pv = nb[in2]; int mv = (deg[pv] > deg[key]) || (deg[pv] == deg[key] && pv > key);
          if (!mv) break; nb[in2 + 1] = pv; in2--; }
        nb[in2 + 1] = key;
      }
      for (i = 0; i < cnt; i++) { int32_t q = nb[i]; if (!vis[q]) { if (tail >= nNode - 1) { rc = 1; goto done; } queue[++tail] = q; vis[q] = 1; } }
    }
  }
  if (cmCount != nNode) { rc = 3; goto done; }
  for (i = 0; i < nNode; i++) { int32_t nd = cm[nNode - 1 - i]; newToOld[i] = nd; oldToNew[nd] = i; }
  { int32_t bw = 0, li, lj, na, nb2, da, db, df;
    for (e = 0; e < nElem; e++) for (li = 0; li < 16; li++) {
      if (elnode[e * 16 + li] < 0) continue; na = nodes[e * 8 + li / 2]; if (na < 0) continue;
      da = 2 * oldToNew[na] + li - (li / 2) * 2;
      for (lj = 0; lj < 16; lj++) {
        if (elnode[e * 16 + lj] < 0) continue; nb2 = nodes[e * 8 + lj / 2]; if (nb2 < 0) continue;
        db = 2 * oldToNew[nb2] + lj - (lj / 2) * 2; df = db - da; if (df < 0) df = -df; if (df > bw) bw = df;
      }
    }
    *candBw = bw; }
done:
  if (adj) for (i = 0; i < nNode; i++) hfree(adj[i]);
  hfree(adj); hfree(deg); hfree(cap); hfree(queue); hfree(cm); hfree(vis); hfree(nb);
  return rc;
#undef ADDN
}

/* Q8 hourglass stabilization, cached-mode path of P6AddQ8HourglassStabilization
 * (modes and HourglassKScaled already valid for the current shape generation
 * and scale), followed by P3KmatIsUnsymmetric. Only for elements whose tangent
 * this DLL just built (tdone). hgdone[k]=-1: handled (applied, or not required
 * when applyHg=0); unsym[k]=-1 when the element operator is unsymmetric.
 * H (int64 offsets): 0 HgShapeGen(i32) 1 Dmat(2,2) 2 HgKReady(i16) 3 HgKScaledReady(i16)
 *                    4 HgCachedScale 5 HgKScaled(16x16) 6 HgScale 7 HgModeCount(i32)
 * returns number of elements where hourglass stiffness was applied. */
__declspec(dllexport) int32_t __stdcall sfm_hg_apply(const int64_t *P, const int16_t *tdone, int16_t *hgdone, int16_t *unsym,
    const int64_t *H, int32_t shapeGen, double hgFactor, int32_t applyHg) {
  char *base = (char *)(intptr_t)P[0]; const int64_t *off = (const int64_t *)(intptr_t)P[1];
  int nElem = (int)P[2]; const double *mt = (const double *)(intptr_t)P[9]; int nMat = (int)P[10];
  int k, i, j, applied = 0;
  if (!g_ready) return -1;
  fp_clear();
  for (k = 0; k < nElem; k++) {
    char *e = base + off[O_STRIDE] * (int64_t)k; double *K; int mat;
    hgdone[k] = 0; unsym[k] = 0;
    if (!tdone[k]) continue;
    mat = *EP(e, off[O_MATNO], int32_t);
    if (mat < 0 || mat >= nMat) continue;
    K = EP(e, off[O_KMAT], double);
    if (applyHg) {
      double thick = mt[mat * MAT_STRIDE + M_THICK], shear, scale; const double *KS;
      if (!(hgFactor > 0.0) || !(thick > 0.0)) continue;
      if (*EP(e, H[7], int32_t) <= 0 || *EP(e, H[0], int32_t) != shapeGen) continue;
      shear = ABS(*EP(e, H[1], double));
      if (shear <= 1E-30) continue;
      scale = (xd)hgFactor * shear * thick;
      if (!*EP(e, H[2], int16_t) || !*EP(e, H[3], int16_t)) continue;
      if (ABS((xd)*EP(e, H[4], double) - scale) > 0.0) continue;
      *EP(e, H[6], double) = scale;
      KS = EP(e, H[5], double);
      for (j = 0; j < 16; j++) for (i = 0; i < 16; i++) { double v = KS[i + 16 * j]; if (ABS(v) > 0.0) K[i + 16 * j] = (xd)K[i + 16 * j] + v; }
      applied++;
    }
    { double sc = 0.0, gap = 0.0;
      for (i = 0; i < 16; i++) for (j = 0; j < 16; j++) {
        double a = K[i + 16 * j], b = K[j + 16 * i];
        if (ABS(a) > sc) sc = ABS(a);
        if (ABS((xd)a - b) > gap) gap = ABS((xd)a - b);
      }
      if (sc < 1.0) sc = 1.0;
      unsym[k] = ((xd)gap > (xd)0.0000001 * sc) ? -1 : 0; }
    hgdone[k] = -1;
  }
  fp_clear();
  return applied;
}
