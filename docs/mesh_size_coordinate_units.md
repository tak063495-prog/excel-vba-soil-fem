# Comparable size warnings for Q8 and T6

`MESH_QUALITY_MIN_JACOBIAN_RATIO` compares the local mapping size with the
square of the model span. A straight Q8 square of area A has determinant A/4.
A straight T6 of the same area uses the Duffy map and its smallest determinant
at the actual 2x2 Gauss points is A(1-1/sqrt(3))/4. Comparing these raw numbers
used to give T6 an additional size warning solely because of reference
coordinates.

The size-warning calculation now removes the fixed Duffy factor for T6. Thus
both affine elements have the same equivalent determinant A/4. The Q8 policy,
settings, physical-area check, geometry validity, independent shape targets,
and reported raw integration-coordinate determinants retain their meanings.
This changes the diagnostic count, not a material law or an analysis force.
Smaller triangles can still receive size warnings because their physical
area is smaller.

`tests/check_size_coordinate_units.py` exercises the real VBA with a unit-area
Q8 and a unit-area T6 sharing an edge. It checks thresholds on both sides,
uniform scaling and translation, and preservation of the raw determinant.
`--legacy` substitutes the historical raw-coordinate comparison to reproduce
the coordinate-dependent warning. Shape, invalid geometry, diagonal choice
and edge-flip controls also execute the existing VBA regression procedures.

Run with a canonical workbook and a new output directory:

```text
python tests/check_size_coordinate_units.py --book path/to/model.xlsm --out tests/tmp/size-units
```
