"""Independent reference operators for fixed-stress Q8/T6 assembly.

No Excel/VBA symbols are imported here.  The returned blocks are suitable for
cross-checking the VBA module after a workbook build.
"""
import numpy as np


def q8_shape(x, y):
    n = np.zeros(8, dtype=np.result_type(x, y, float))
    n[0] = .25*(1-x)*(1-y)*(-1-x-y); n[1] = .25*(1+x)*(1-y)*(-1+x-y)
    n[2] = .25*(1+x)*(1+y)*(-1+x+y); n[3] = .25*(1-x)*(1+y)*(-1-x+y)
    n[4] = .5*(1-x*x)*(1-y); n[5] = .5*(1+x)*(1-y*y)
    n[6] = .5*(1-x*x)*(1+y); n[7] = .5*(1-x)*(1-y*y)
    return n


def t6_shape(x, y):
    r = (1+x)*(1-y)/4; s = (1+y)/2; a = 1-r-s
    return np.array([a*(2*a-1), r*(2*r-1), s*(2*s-1), 0,
                     4*a*r, 4*r*s, 0, 4*s*a], dtype=np.result_type(x, y, float))


def pressure_shape(kind, x, y):
    if kind == "q8":
        return np.array([.25*(1-x)*(1-y), .25*(1+x)*(1-y),
                         .25*(1+x)*(1+y), .25*(1-x)*(1+y)], dtype=np.result_type(x, y, float))
    r = (1+x)*(1-y)/4; s = (1+y)/2
    return np.array([1-r-s, r, s, 0], dtype=np.result_type(x, y, float))


def grad(fun, x, y):
    h = 1e-30
    return np.array([np.imag(fun(x+1j*h, y))/h,
                     np.imag(fun(x, y+1j*h))/h])


def quadrature(kind):
    z = [-np.sqrt(3/5), 0., np.sqrt(3/5)]; w = [5/9, 8/9, 5/9]
    for i, x in enumerate(z):
        for j, y in enumerate(z):
            if kind == "t6":
                yield (1+x)*(1-y)/4, (1+y)/2, w[i]*w[j]*(1-y)/8
            else:
                yield x, y, w[i]*w[j]


def reference_local(coords, kind, alpha, storage, kx, ky, gamma, E, nu, thickness=1.):
    coords = np.asarray(coords, float)
    M = np.zeros((4, 4)); H = np.zeros((4, 4)); L = np.zeros((4, 4)); Q = np.zeros((16, 4)); g = np.zeros(4)
    Kdr = E/(3*(1-2*nu))
    def triangle_shape(r, s):
        a = 1-r-s
        return np.array([a*(2*a-1),r*(2*r-1),s*(2*s-1),0,4*a*r,4*r*s,0,4*s*a],dtype=np.result_type(r,s,float))
    def triangle_pressure(r,s):
        return np.array([1-r-s,r,s,0],dtype=np.result_type(r,s,float))
    shape = q8_shape if kind == "q8" else triangle_shape
    pressure = (lambda a,b: pressure_shape("q8",a,b)) if kind == "q8" else triangle_pressure
    for x, y, wt in quadrature(kind):
        ns = shape(x, y)
        dxi = np.array([grad(lambda a, b, i=i: shape(a, b)[i], x, y)[0] for i in range(8)])
        deta = np.array([grad(lambda a, b, i=i: shape(a, b)[i], x, y)[1] for i in range(8)])
        J = np.array([[dxi @ coords[:, 0], deta @ coords[:, 0]],
                      [dxi @ coords[:, 1], deta @ coords[:, 1]]])
        det = np.linalg.det(J)
        assert det > 0
        invJT = np.linalg.inv(J).T
        grads = invJT @ np.vstack((dxi, deta))
        pn = pressure(x, y)
        pxi = np.array([grad(lambda a, b, i=i: pressure(a, b)[i], x, y)[0] for i in range(4)])
        peta = np.array([grad(lambda a, b, i=i: pressure(a, b)[i], x, y)[1] for i in range(4)])
        pg = invJT @ np.vstack((pxi, peta)); jac = det*wt*thickness
        M += storage*np.outer(pn, pn)*jac
        H += (kx*np.outer(pg[0], pg[0]) + ky*np.outer(pg[1], pg[1]))/gamma*jac
        L += alpha*alpha/Kdr*np.outer(pn, pn)*jac
        g += ky*pg[1]*jac
        for i in range(8):
            Q[2*i] += alpha*grads[0, i]*pn*jac
            Q[2*i+1] += alpha*grads[1, i]*pn*jac
    return M, H, L, Q, g


def test_q8_t6_partition_and_quadrature():
    for kind in ("q8", "t6"):
        for x, y, _ in quadrature(kind):
            assert np.isclose(pressure_shape(kind, x, y).sum(), 1.)
            assert np.isclose((q8_shape(x, y) if kind == "q8" else t6_shape(x, y)).sum(), 1.)


def test_skew_affine_blocks_are_symmetric_and_mass_integrates():
    q8 = np.array([[0., 0.], [2., .2], [2.2, 1.5], [-.1, 1.], [1., .1], [2.1, .85], [1.05, 1.25], [-.05, .5]])
    t6 = np.array([[0., 0.], [2., .2], [.2, 1.5], [0., 0.], [1., .1], [1.1, .85], [0., 0.], [.1, .75]])
    for coords, kind in ((q8, "q8"), (t6, "t6")):
        M, H, L, Q, g = reference_local(coords, kind, .8, 2., 1e-5, 2e-5, 9.81, 30000., .3)
        assert np.allclose(M, M.T); assert np.allclose(H, H.T); assert np.allclose(L, L.T)
        assert M.sum() > 0.0
        active = [0, 1, 2, 3] if kind == "q8" else [0, 1, 2]
        assert np.linalg.eigvalsh(M[np.ix_(active, active)]).min() > 0
        assert np.isfinite(Q).all() and np.isfinite(g).all()


def test_fixed_stress_signs_and_monolithic_limit():
    K = np.array([[4., 1.], [1., 3.]])
    Q = np.array([[.7], [.2]]); M = np.array([[2.]])
    fu = np.array([1., .5]); fp = np.array([.2]); dt = .4
    A = np.block([[K, -Q], [Q.T, M]])
    exact = np.linalg.solve(A, np.r_[fu, fp])
    u = np.zeros(2); p = np.zeros(1)
    for _ in range(2000):
        old = np.r_[u, p]
        stabilization = np.array([[.3]])
        p = np.linalg.solve(M+stabilization, fp-Q.T@u+stabilization@p)
        u = np.linalg.solve(K, fu + Q[:, 0]*p[0])
        if np.linalg.norm(np.r_[u, p]-old) < 1e-13: break
    assert np.allclose(np.r_[u, p], exact, atol=1e-10)
    assert np.linalg.norm(A@np.r_[u, p]-np.r_[fu, fp]) < 1e-10


def test_stabilized_flow_mass_residual():
    M = np.array([[2., .5], [.5, 1.]])
    p = np.array([1., 3.]); pc = np.array([1., 2.]); residual = M@(p-pc); residual[0] = 0.
    assert np.allclose(residual, [0., 1.])
