"""Audit current-source SRM evidence without promoting numerical failure to FOS."""
import argparse
import csv
import hashlib
import json
import math
from pathlib import Path

import numpy as np
from verify_inputs import workbook_tables, equal


def read(path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def get_settings(book):
    cells=workbook_tables(book)["設定"]
    return {v:cells.get("C"+a[1:]) for a,v in cells.items()
            if a.startswith("E") and a[1:].isdigit() and isinstance(v,str) and v.isascii()}


def trial_history(log):
    rows=[]
    if not log.exists():return rows
    for row in csv.DictReader(log.read_text(encoding="cp932",errors="replace").splitlines()):
        if row.get("place")!="SRM試行" or row.get("status") not in ("PASS","FAIL"):continue
        try:fs=float(row["Fs"])
        except (ValueError,KeyError):continue
        rows.append(dict(fs_log_rounded=fs,status=row["status"],residual=float(row["relres"]),message=row.get("message","")))
    return rows


def displacement_geometry(fixture,flat):
    rows=np.array(flat,dtype=float).reshape(-1,3)
    displacement={int(r[0]):r[1:] for r in rows}
    coordinates={int(r[0]):np.array(r[1:3],dtype=float) for r in fixture["nodes"]}
    maximum=0.;min_det=math.inf;max_rotation=0.;gauss=[]
    root=1/math.sqrt(3)
    for e in fixture["elements"]:
        ids=e[1:9];xy=np.array([coordinates[n] for n in ids]);u=np.array([displacement[n] for n in ids])
        for gp,(r,s) in enumerate([(-root,-root),(root,-root),(root,root),(-root,root)],1):
            dr=np.array([.25*(1-s)*(2*r+s),.25*(1-s)*(2*r-s),.25*(1+s)*(2*r+s),.25*(1+s)*(2*r-s),-r*(1-s),.5*(1-s*s),-r*(1+s),-.5*(1-s*s)])
            ds=np.array([.25*(1-r)*(r+2*s),.25*(1+r)*(-r+2*s),.25*(1+r)*(r+2*s),.25*(1-r)*(-r+2*s),-.5*(1-r*r),-(1+r)*s,.5*(1-r*r),-(1-r)*s])
            natural=np.vstack([dr,ds]);jacobian=natural@xy
            physical=np.linalg.solve(jacobian,natural)
            gradient=(physical@u).T
            ex=gradient[0,0];ey=gradient[1,1];shear=gradient[0,1]+gradient[1,0]
            determinant=float(np.linalg.det(np.eye(2)+gradient))
            rotation=.5*(gradient[1,0]-gradient[0,1])
            maximum=max(maximum,abs(ex),abs(ey),abs(shear));min_det=min(min_det,determinant);max_rotation=max(max_rotation,abs(rotation))
            gauss.append(dict(element=e[0],gauss=gp,ex=float(ex),ey=float(ey),engineering_shear=float(shear),deformed_det=float(determinant)))
    return dict(max_component_mm=float(np.max(np.abs(rows[:,1:]))*1000),
                max_vector_mm=float(np.max(np.linalg.norm(rows[:,1:],axis=1))*1000),
                max_absolute_small_strain=float(maximum),min_deformed_gradient_det=min_det,
                max_infinitesimal_rotation_radians=max_rotation,gauss=gauss)


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("--output",type=Path,required=True)
    parser.add_argument("--fixtures",type=Path,required=True)
    parser.add_argument("--protocol",type=Path,default=Path(__file__).with_name("practical_protocol.json"))
    args=parser.parse_args();protocol=read(args.protocol)
    fixtures={f["name"]:f for f in read(args.fixtures)["fixtures"]}
    reviewed_sources={protocol["source_build"]:protocol["source_sha256"].lower()}
    reviewed_sources.update({s["build"]:s["sha256"].lower() for s in protocol.get("additional_reviewed_sources",[])})
    thresholds=protocol["project_review_thresholds"]
    report=dict(protocol=protocol,fixture_sha256=hashlib.sha256(args.fixtures.read_bytes()).hexdigest(),cases=[])
    for path in sorted(args.output.glob("*/results/*.json")):
        result=read(path)
        if "metrics" not in result:continue
        if result["fixture"] not in fixtures:continue # separate elastic harness check
        fixture=fixtures[result["fixture"]];book=path.parent.parent/"cases"/(result["case"]+".xlsm")
        settings=get_settings(book);metrics=dict(zip(result["metric_names"],result["metrics"]))
        failure=dict(zip(result["failure_names"],result["failure_values"]))
        controls=protocol["controls"]
        expected=dict(SRM_TOL=controls["srm_tolerance"],SRM_FMAX=controls["srm_fmax"],SRM_MODE=controls["srm_mode"],SRM_PSI_POLICY="KEEP",SOLVER=controls["solver"],RCM_POLICY=controls["rcm"],SRM_FIXED_FS=0,Q8_HOURGLASS_FACTOR=.05)
        control_errors={k:dict(expected=v,actual=settings.get(k)) for k,v in expected.items() if not equal(v,settings.get(k))}
        if settings.get("FLOW_POLICY")!=result["flow_policy"]:control_errors["FLOW_POLICY"]="saved/result mismatch"
        source_ok=result["source_sha256"].lower()==reviewed_sources.get(result["build"])
        role="baseline_flow_run" if result["flow_policy"]==controls["flow_policy"] else "same_material_flow_solver_control" if fixture["materials"][0][5]==0 else "different_flow_control"
        success=result["status"].startswith("PASS|") and bool(metrics["analysis_ok"])
        bracket=success and bool(metrics["fos_bracket"]) and failure["fos_interpretation"]=="NUMERICAL_BRACKET"
        geometry=displacement_geometry(fixture,result["displacements_flat"])
        gauss=geometry.pop("gauss")
        destination=path.parent.parent/"analysis";destination.mkdir(exist_ok=True)
        with (destination/(result["case"]+"_strain.csv")).open("w",encoding="utf-8-sig",newline="") as file:
            writer=csv.DictWriter(file,fieldnames=list(gauss[0]));writer.writeheader();writer.writerows(gauss)
        weight=-fixture["checks"]["body_force_sum_y"]
        balance=abs(metrics["reaction_y"]-weight)/max(1,abs(weight)) if success else None
        ref=fixture["reference"].get("fos_reference_range",fixture["reference"].get("published_pass_fail"))
        if ref is None:ref=[fixture["reference"]["fos_reported"]]*2
        fos=metrics["fos_mid"] if bracket else None
        outside=None
        if fos is not None and role!="different_flow_control":
            outside=0 if ref[0]<=fos<=ref[1] else (ref[0]-fos)/ref[0] if fos<ref[0] else (fos-ref[1])/ref[1]
        fixture_identity=result.get("fixture_sha256","").lower()==report["fixture_sha256"]
        numerical_checks=dict(source_hash_and_build=source_ok,controls_match=not control_errors,
                              fixture_identity_recorded_at_run_start=fixture_identity,
                              source_control_errors=control_errors,completed_bracket=bracket,
                              global_residual=success and metrics["relative_residual"]<=thresholds["global_residual"],
                              reaction_balance=balance is not None and balance<=thresholds["relative_reaction_imbalance"])
        if bracket:
            numerical_checks["valid_bracket_endpoints"]=(metrics["fos_fail"]>metrics["fos_pass"]>0 and
                abs(metrics["fos_width"]-(metrics["fos_fail"]-metrics["fos_pass"]))<1e-10 and
                metrics["fos_width"]<=controls["srm_tolerance"]+1e-12 and abs(metrics["fss"]-metrics["fos_pass"])<1e-10)
        elif not success:
            numerical_checks["refuses_invalid_fos"]=(metrics["fss"]==0 and metrics["fos_fail"]==0 and metrics["fos_mid"]==0 and
                not metrics["fos_bracket"] and failure["fos_interpretation"]=="UNDETERMINED")
        history=trial_history(path.parent.parent/"cases"/(result["case"]+"_out")/"run.log")
        row=dict(case=result["case"],fixture=result["fixture"],role=role,flow_policy=result["flow_policy"],
                 source_build=result["build"],
                 literature_comparison_role=fixture["reference"].get("comparison_role","secondary"),
                 status=result["status"],source_sha256=result["source_sha256"],metrics=metrics,failure=failure,
                 numerical_checks=numerical_checks,reference_interval=ref,reference=fixture["reference"],
                 fos_outside_reference_fraction=outside,reference_screen=outside is not None and outside<=thresholds["relative_fos_deviation_outside_published_interval"],
                 relative_reaction_imbalance=balance,geometry=geometry,
                 geometry_state_basis="final full-load converged lower-Fs state" if success else "partial-load failed-trial diagnostic; not equilibrium displacement",
                 trial_history=history,failed_trial_fs_log_rounded=next((r["fs_log_rounded"] for r in reversed(history) if r["status"]=="FAIL"),None),
                 elapsed_seconds=result["elapsed_seconds"],
                 small_deformation_diagnostic="large local deformation" if geometry["max_absolute_small_strain"]>.05 or geometry["min_deformed_gradient_det"]<=0 else "inspect local fields")
        if role=="different_flow_control":
            phi=fixture["materials"][0][5];tan=math.tan(math.radians(phi))
            row["davis_strength_only_equivalent_interval"]=[math.sqrt(metrics[k]**2+tan**2) for k in ["fos_pass","fos_fail"]] if bracket else None
            row["reference_screen"]=False # changed flow rule cannot certify the psi=0 benchmark
        row["reference_assessment"]=("NOT_ESTABLISHED" if not bracket else
            "FLOW_CHANGED" if role=="different_flow_control" else
            "PASS_SCREEN" if row["reference_screen"] else "OUTSIDE_SCREEN")
        report["cases"].append(row)
    report["case_count"]=len(report["cases"])
    report["source_controls_audit_ok"]=all(r["numerical_checks"]["source_hash_and_build"] and r["numerical_checks"]["controls_match"] for r in report["cases"])
    report["fixture_provenance_gaps"]=[r["case"] for r in report["cases"] if not r["numerical_checks"]["fixture_identity_recorded_at_run_start"]]
    report["audit_ok"]=report["source_controls_audit_ok"] and not report["fixture_provenance_gaps"]
    report["critical_fos_not_established_cases"]=[r["case"] for r in report["cases"] if not r["numerical_checks"]["completed_bracket"]]
    (args.output/"practical_comparison.json").write_text(json.dumps(report,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    fields=["case","flow_policy","role","status","critical_fos_mid","fos_pass","fos_fail","fos_mid","reference_low","reference_high","reference_screen","reference_assessment","failure_kind","failed_lambda","relative_residual","geometry_state_basis","max_component_mm","max_vector_mm","max_absolute_small_strain","min_deformed_gradient_det","elapsed_seconds"]
    with (args.output/"practical_comparison.csv").open("w",encoding="utf-8-sig",newline="") as file:
        writer=csv.DictWriter(file,fieldnames=fields);writer.writeheader()
        for r in report["cases"]:
            writer.writerow(dict(case=r["case"],flow_policy=r["flow_policy"],role=r["role"],status=r["status"],
                critical_fos_mid=r["metrics"]["fos_mid"] if r["numerical_checks"]["completed_bracket"] else None,
                reference_assessment=r["reference_assessment"],
                geometry_state_basis=r["geometry_state_basis"],
                **{k:r["metrics"][k] for k in ["fos_pass","fos_fail","fos_mid","relative_residual"]},
                reference_low=r["reference_interval"][0],reference_high=r["reference_interval"][1],reference_screen=r["reference_screen"],
                failure_kind=r["failure"]["failure_kind"],failed_lambda=r["failure"]["failed_lambda"],
                **{k:r["geometry"][k] for k in ["max_component_mm","max_vector_mm","max_absolute_small_strain","min_deformed_gradient_det"]},elapsed_seconds=r["elapsed_seconds"]))
    print("Audited",report["case_count"],"cases; controls/hash",report["source_controls_audit_ok"],"fixture provenance gaps",report["fixture_provenance_gaps"],"no FOS",report["critical_fos_not_established_cases"])
    if not report["audit_ok"]:raise SystemExit(1)


if __name__=="__main__":main()
