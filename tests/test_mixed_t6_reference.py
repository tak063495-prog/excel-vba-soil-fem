"""Independent reference checks for the six-node triangle used by mixed T6.

The checks are deliberately solver-independent: they exercise the same
quadratic shape/derivative convention and the four-point Duffy rule expected
by the native element routine.
"""

import numpy as np


def shape_and_derivatives(r, s):
    l1, l2, l3 = 1.0 - r - s, r, s
    n = np.array([
        l1 * (2 * l1 - 1), l2 * (2 * l2 - 1), l3 * (2 * l3 - 1),
        4 * l1 * l2, 4 * l2 * l3, 4 * l3 * l1,
    ])
    # derivatives with respect to (r,s)
    dr = np.array([1 - 4 * l1, 4 * l2 - 1, 0, 4 * (l1 - l2), 4 * l3, -4 * l3])
    ds = np.array([1 - 4 * l1, 0, 4 * l3 - 1, -4 * l2, 4 * l2, 4 * (l1 - l3)])
    return n, dr, ds


def duffy_points():
    # 2x2 Gauss-Legendre on (a,b), mapped to r=a, s=(1-a)b,
    # with Jacobian (1-a). Four strictly interior points.
    g = (1 - 1 / np.sqrt(3)) / 2, (1 + 1 / np.sqrt(3)) / 2
    for a in g:
        for b in g:
            yield a, (1 - a) * b, (1 - a) / 4


def triangle_jacobian(vertices):
    p0, p1, p2 = vertices
    return np.array([[p1[0] - p0[0], p2[0] - p0[0]],
                     [p1[1] - p0[1], p2[1] - p0[1]]])


def b_matrix(vertices, r, s):
    _, dr, ds = shape_and_derivatives(r, s)
    j = triangle_jacobian(vertices)
    invjt = np.linalg.inv(j).T
    grads = invjt @ np.vstack((dr, ds))
    b = np.zeros((3, 12))
    for i in range(6):
        b[0, 2 * i] = grads[0, i]
        b[1, 2 * i + 1] = grads[1, i]
        b[2, 2 * i] = grads[1, i]
        b[2, 2 * i + 1] = grads[0, i]
    return b


def test_quadratic_shape_partition_and_linear_reproduction():
    vertices = np.array([[0.0, 0.0], [2.0, 0.1], [0.2, 1.5]])
    mids = np.array([(vertices[0] + vertices[1]) / 2,
                     (vertices[1] + vertices[2]) / 2,
                     (vertices[2] + vertices[0]) / 2])
    points = np.vstack((vertices, mids))
    for r, s, _ in duffy_points():
        n, _, _ = shape_and_derivatives(r, s)
        assert np.isclose(n.sum(), 1.0)
        xy = n @ points
        assert np.allclose(xy, (1-r-s)*vertices[0] + r*vertices[1] + s*vertices[2])


def test_duffy_area_and_constant_body_force():
    vertices = np.array([[0.0, 0.0], [2.0, 0.1], [0.2, 1.5]])
    mids = np.array([(vertices[0] + vertices[1]) / 2,
                     (vertices[1] + vertices[2]) / 2,
                     (vertices[2] + vertices[0]) / 2])
    points = np.vstack((vertices, mids))
    detj = abs(np.linalg.det(triangle_jacobian(vertices)))
    mass = np.zeros(6)
    for r, s, w in duffy_points():
        n, _, _ = shape_and_derivatives(r, s)
        mass += n * w * detj
    assert np.isclose(mass.sum(), detj / 2)
    # Quadratic corner shape functions have zero signed integral; edge-midpoint
    # functions carry the positive consistent body-force weights.
    assert np.allclose(mass[:3], 0.0, atol=1e-12)
    assert np.all(mass[3:] > 0)


def test_affine_patch_rigid_modes_and_stiffness_rank():
    vertices = np.array([[0.0, 0.0], [2.0, 0.1], [0.2, 1.5]])
    d = np.vstack((vertices, (vertices[0] + vertices[1]) / 2,
                   (vertices[1] + vertices[2]) / 2, (vertices[2] + vertices[0]) / 2))
    ke = np.zeros((12, 12))
    for r, s, w in duffy_points():
        b = b_matrix(vertices, r, s)
        ke += b.T @ b * abs(np.linalg.det(triangle_jacobian(vertices))) * w
    assert np.linalg.matrix_rank(ke, tol=1e-10) == 9
    # u_x=x is represented exactly by quadratic interpolation at nodal values.
    dof = np.empty(12)
    dof[0::2] = d[:, 0]
    dof[1::2] = 0.0
    for r, s, _ in duffy_points():
        b = b_matrix(vertices, r, s)
        strain = b @ dof
        assert np.allclose(strain, [1.0, 0.0, 0.0])


def test_linear_head_has_constant_gradient_and_zero_interpolation_error():
    vertices = np.array([[0.0, 0.0], [2.0, 0.1], [0.2, 1.5]])
    mids = np.array([(vertices[0] + vertices[1]) / 2,
                     (vertices[1] + vertices[2]) / 2,
                     (vertices[2] + vertices[0]) / 2])
    points = np.vstack((vertices, mids))
    head = points[:, 0] + 2 * points[:, 1]
    expected = np.array([1.0, 2.0])
    j = triangle_jacobian(vertices)
    for r, s, _ in duffy_points():
        _, dr, ds = shape_and_derivatives(r, s)
        grad = np.linalg.inv(j).T @ np.array([dr @ head, ds @ head])
        assert np.allclose(grad, expected)


def test_duffy_stiffness_matches_standard_triangle_rule():
    vertices = np.array([[0., 0.], [2., .1], [.2, 1.5]])
    material = np.array([[3., 1., 0.], [1., 3., 0.], [0., 0., 1.]])
    standard = ((1/6, 1/6, 1/6), (2/3, 1/6, 1/6), (1/6, 2/3, 1/6))
    def integrate(rule):
        k = np.zeros((12,12))
        for r,s,w in rule:
            b = b_matrix(vertices,r,s)
            k += b.T @ material @ b * np.linalg.det(triangle_jacobian(vertices))*w
        return k
    kd = integrate(duffy_points())
    assert np.allclose(kd, integrate(standard), atol=1e-12, rtol=1e-12)
    nodes = np.vstack((vertices,(vertices[0]+vertices[1])/2,
                       (vertices[1]+vertices[2])/2,(vertices[2]+vertices[0])/2))
    for x,y in ((np.ones(6),np.zeros(6)),(np.zeros(6),np.ones(6)),(-nodes[:,1],nodes[:,0])):
        u=np.column_stack((x,y)).ravel()
        assert np.max(np.abs(kd@u)) < 1e-12
