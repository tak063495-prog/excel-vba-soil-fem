"""Independent analytical checks for the additional validation calculations."""
import argparse
import json
import math
from pathlib import Path

import numpy as np
from analyze_practical_results import displacement_geometry
from independent_bishop import evaluate, ground


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("--fixtures",type=Path,required=True)
    parser.add_argument("--bishop",type=Path,required=True)
    parser.add_argument("--output",type=Path,required=True)
    args=parser.parse_args()
    fixtures=json.loads(args.fixtures.read_text(encoding="utf-8-sig"))["fixtures"]
    fixture=next(f for f in fixtures if f["name"]=="griffiths_1999_coarse")
    flat=[]
    for n in fixture["nodes"]:
        flat.extend([n[0],.001*n[1]-.001*n[2],.002*n[1]-.002*n[2]])
    result=displacement_geometry(fixture,flat)
    for point in result["gauss"]:
        assert abs(point["ex"]-.001)<1e-12
        assert abs(point["ey"]+.002)<1e-12
        assert abs(point["engineering_shear"]-.001)<1e-12
        assert abs(point["deformed_det"]-.999)<1e-12
    assert abs(result["max_infinitesimal_rotation_radians"]-.0015)<1e-12

    bishop=json.loads(args.bishop.read_text(encoding="utf-8-sig"))
    case=next(c for c in bishop["cases"] if c["model"]["name"]=="griffiths_ex3_phi0")
    model=case["model"]; best=case["runs"][0]["slice_convergence"][-1]
    xc,yc,radius=[best[k] for k in ("circle_center_x","circle_center_y","radius")]
    entry,exit=best["entry_x"],best["exit_x"]
    # For phi=0, Bishop is exactly resisting circular-arc shear moment /
    # gravity moment. Integrate each linear terrain segment analytically.
    breaks=sorted(set([entry,exit]+[x for x in [model["ground"]["crest"],model["ground"]["crest"]+model["ground"]["run"]] if entry<x<exit]))
    drive=0.
    for a,b in zip(breaks[:-1],breaks[1:]):
        ya,yb=[float(ground(np.array(x),**model["ground"])) for x in (a,b)]
        slope=(yb-ya)/(b-a); intercept=ya-slope*a
        def primitive(x):
            z=x-xc
            return (yc-intercept-slope*xc)*z*z/2-slope*z**3/3+(max(0,radius*radius-z*z))**1.5/3
        drive+=model["gamma"]*(primitive(b)-primitive(a))
    theta=math.asin((exit-xc)/radius)-math.asin((entry-xc)/radius)
    exact=model["c"]*radius*radius*theta/drive
    approximate,_=evaluate(np.array([[entry,exit,best["center_distance"]]]),model,640)
    error=abs(float(approximate[0])-exact)/exact
    assert error<1e-4,(exact,approximate,error)
    checks=dict(status="PASS",affine_q8_gauss_points=len(result["gauss"]),
                phi0_exact_circle_fos=exact,phi0_640_slice_fos=float(approximate[0]),
                relative_integration_error=error)
    args.output.parent.mkdir(parents=True,exist_ok=True)
    args.output.write_text(json.dumps(checks,indent=2)+"\n",encoding="utf-8")
    print(checks)


if __name__=="__main__":main()
