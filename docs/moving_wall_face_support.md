# Provisional support attached to a moving wall

The optional `壁面支持` input is a massless, compression-only horizontal support
for soil beside a vertical wall. It returns its forces and moment to two nodes
on the existing wall. It adds no fixed support to the ground and no vertical
constraint to the soil. It is an idealized support for external/global stability,
not a design or a capacity check of a wall facing, anchor or connection.

Open `境界・水条件 → 壁面支持`. Input columns are:

| Column | Meaning |
|---|---|
| A | Unique positive support ID |
| B | Soil node ID |
| C | Upper reference node on the wall |
| D | Lower reference node on the wall |
| E | Assembled nodal spring stiffness, force/length (kN/m in the example units) |
| F | +1 for soil on the right of the wall, −1 for soil on the left |
| G | Comment |

All three nodes must have the same X coordinate. The reference nodes define the
wall's translation and its small rotation. The upper reference node must be
higher than the lower one. With `t=(y_soil-y_upper)/(y_upper-y_lower)`, the normal
gap is `sense*(u_soil-(1+t)*u_upper+t*u_lower)`. A negative gap carries compressive
load; an opening carries no load. Each closed support contributes `k*C*C^T` to
the tangent and `k*gap*C` to the internal force. These contributions cancel for
a common rigid translation or rotation and conserve resultant force and moment.

Use stiffness sensitivity and the reported gap to verify a sufficiently stiff
approximation. Avoid assigning infinite stiffness. A nodal spring assembled from
a distributed face penalty must account for tributary length and model thickness.
The inputs do not supply a real member stiffness or strength automatically.

`壁面支持結果` reports each support ID, normal gap, force on the soil and contact
state. The wall receives the corresponding opposite resultant and moment. The
support has no plastic history and remains outside SRM strength reduction.

The first implementation supports fixed-topology static GRAVITY, LOAD/APPLY,
UNLOAD and SRM stages. Reference nodes must belong to initially active elements.
Birth/death, displacement/stress resets, geometry updates and time coupling are
rejected when these supports are present. Support IDs refer to the current mesh;
save and clear the table before remeshing, then set the new node IDs. Editing the
table invalidates displayed results and the model preview.

Validation uses two stacked Q8 elements with a moving wall and an outward point
load. The BAND and full CSR assembly paths satisfy force balance and agree with
finite-difference tangents. Tests also verify rigid-motion invariance, zero force
on opening, action/reaction moment balance, checkpoint resumes, real front-sheet
write-through, dirty marking and protection against stale IDs after remeshing.
Fixtures are in `tests/fixtures/face_support`; audits are
`tests/check_face_support.py` and `tests/check_face_support_ui.py`. Audit drivers
operate on disposable copies and are not imported into production workbooks.
