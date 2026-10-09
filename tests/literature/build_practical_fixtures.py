"""Additional dry/total-stress benchmarks; geometry assumptions remain explicit."""
import argparse
import json
from pathlib import Path

from build_fixtures import case, material, q8_mesh

SOURCE = "https://inside.mines.edu/~vgriffit/slope64/Griffiths%20and%20Lane%201999"
STAGES = [[1, "GRAVITY", None, None, 1, "自重"],
          [2, "SRM", None, None, 1, "強度低減"]]


def regular_axis(length, count):
    return [length * i / count for i in range(count + 1)]


def make_cases():
    # Retain the previously reviewed reference cases without changing their inputs.
    originals = json.loads(Path(__file__).with_name("fixtures.json").read_text(encoding="utf-8"))
    fixtures = [f for f in originals["fixtures"] if f["reference"]["source_id"] != "elastic_validation"]
    for f in fixtures:
        f["reference"]["comparison_role"] = (
            "primary_reconstructed" if f["reference"]["source_id"] == "griffiths_lane_1999_ex1"
            else "secondary_initial_stress_not_matched")

    for count, label in [(8, "coarse"), (12, "fine")]:
        # Fig6: 2H crest + 2H slope run + 2H toe bench; total depth D H=2H.
        # The 45-degree annotation belongs to the weak-layer outcrop, not the slope.
        us = regular_axis(60, count * 3)
        def top(x):
            return 20 if x <= 20 else 30 - .5 * x if x < 40 else 10
        nodes, elements, checks = q8_mesh(
            us, regular_axis(1, count), lambda x, v: (x, v * top(x)),
            lambda x, y: (int(abs(x) < 1e-9 or abs(x-60) < 1e-9 or abs(y) < 1e-9),
                          int(abs(y) < 1e-9)))
        assert abs(checks["area_sum"] - 900) < 1e-8
        reference = dict(
            source_id="griffiths_lane_1999_ex3_homogeneous", primary_url=SOURCE,
            source_pages="journal394-395, Fig6-8", H=10, D=2,
            fos_reported=1.47, fos_reference_range=[1.45, 1.50],
            reference_definition="Taylor1.47; source FE curve rounded to nearest0.05",
            comparison_role="primary_reconstructed", expected_mechanism="circular base failure",
            explicit_assumptions=[
                "cu2/cu1=1 removes the weak-layer material discontinuity",
                "Fig6 slope2H horizontal:1H vertical; foundation depthH, total depth2H",
                "H10m,gamma20,cu50 preserve cu/(gamma H)=0.25",
                "E100000kPa,nu0.3 are declared nominal reconstruction values",
                "both vertical sides ux=0,base ux=uy=0; side restraints are reconstruction assumptions",
                "stress-free gravity turn-on; phi=psi=0 total-stress Tresca, no pore-pressure model"])
        fixtures.append(case("griffiths_1999_phi0_"+label, nodes, elements, checks,
                             material(100000, .3, 20, 50, 0), STAGES, [], reference))

    # Fig5 only labels depth. The right extension is reconstructed from the drawing;
    # it must not be presented as the author's original nodal input.
    us = regular_axis(12, 6) + [12+20*i/10 for i in range(1,11)] + [32+16*i/8 for i in range(1,9)]
    def top_foundation(x):
        return 15 if x <= 12 else 21-.5*x if x < 32 else 5
    nodes, elements, checks = q8_mesh(
        us, regular_axis(1, 10), lambda x,v:(x,v*top_foundation(x)),
        lambda x,y:(int(abs(x)<1e-9 or abs(x-48)<1e-9 or abs(y)<1e-9),int(abs(y)<1e-9)))
    assert abs(checks["area_sum"]-460)<1e-8
    reference=dict(source_id="griffiths_lane_1999_ex2",primary_url=SOURCE,
                   source_pages="journal392-394,Fig2/5",H=10,D=1.5,fos_reported=1.4,
                   fos_reference_range=[1.35,1.40],comparison_role="secondary_extent_reconstructed",
                   expected_mechanism="toe failure with shallow foundation penetration",
                   explicit_assumptions=["same material in slope and foundation; foundation depthH/2,total height1.5H",
                                         "crest12m,slope run20m,right bench16m reconstructed from Fig5",
                                         "side rollers and fixed base; original mesh/side boundary not published",
                                         "H10,gamma20,c10,phi20,psi0,E100000,nu0.3; stress-free gravity turn-on"])
    fixtures.append(case("griffiths_1999_foundation",nodes,elements,checks,
                         material(100000,.3,20,10,20),STAGES,[],reference))
    return fixtures


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("--output",type=Path,default=Path(__file__).parent.parent/"tmp/practical_validation/fixtures.json")
    args=parser.parse_args()
    fixtures=make_cases()
    args.output.parent.mkdir(parents=True,exist_ok=True)
    args.output.write_text(json.dumps(dict(schema="practical_srm_fixtures_v1",fixtures=fixtures),
                                     ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    for f in fixtures:
        print(f["name"],len(f["nodes"]),len(f["elements"]),"area",round(f["checks"]["area_sum"],6))


if __name__=="__main__":
    main()
