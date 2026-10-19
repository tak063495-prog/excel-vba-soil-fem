"""Create a new disposable VBA copy with hierarchical QPC profiling."""
from pathlib import Path
import argparse
import json
import re
import shutil

TARGETS = {
    "CONTROL.bas": ["P0_RunAnalysis"],
    "FEMEngine.bas": ["P3EvaluateTrialState", "P3RestoreCommittedMaterialState",
        "P3CommitMaterialState", "P3RebuildTangentFromSpmat", "P3ApplyLineSearch",
        "P3AccumulateResidualMetrics", "P3InitializeTrialState",
        "P3CaptureSrmSuccessSnapshot", "P3RestoreSrmSuccessSnapshot", "P3RunLoadStages",
        "P3TryResidualRecovery", "P3TryNewtonWatchdog", "P3ShouldRebuildTangent"],
    "FEMNativeMat.bas": ["FEMMatEval", "FEMMatTangent", "FEMMatCommonParams",
        "FEMMatBuildMaterialTable"],
    "FEMSolver.bas": ["SetTotalMat", "P6SolveNonsymmetricBand", "P6SolveResidualRecovery",
        "BandSolver", "SetBoundaryCondition", "P6SolveCSR", "P6SolveCSRGMRES", "P6BuildCSRILU"],
    "FEMIo.bas": ["P6SolverEvent", "FEMIoAppendCsv", "FEMIoOutDir", "FEMIoFlushEvents", "FEMIoEnsureDir",
        "FEMAppendRunLog", "P6PerfEmit", "P6PhysOnRebuild", "GetDiagnosticSheet"],
    "FEMCore.bas": ["SetAnalysisFailure"],
    "FEM.bas": ["SaveP0Progress"],
    "FEMNativeKern.bas": ["FEMKernResidual", "FEMKernOn", "FEMKernGmres"],
}
HEADER = re.compile(r"\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function)\s+(\w+)", re.I)

def code_mask(line):
    """Blank strings/comments while preserving positions for token edits."""
    mask=list(line); quoted=False; i=0
    while i<len(line):
        ch=line[i]
        if ch=='"':
            mask[i]=' '
            if quoted and i+1<len(line) and line[i+1]=='"':
                mask[i+1]=' '; i+=2; continue
            quoted=not quoted
        elif ch=="'" and not quoted:
            mask[i:]=[' ']*(len(line)-i); break
        elif quoted:
            mask[i]=' '
        i+=1
    return ''.join(mask)

def inject(text, names, ids):
    out=[]; current=None; header=False; found=[]
    wanted={n.casefold() for n in names}
    for line in text.splitlines(keepends=True):
        mask=code_mask(line)
        if current is None:
            match=HEADER.match(mask)
            if match and match[2].casefold() in wanted:
                current=match[2]; header=True; found.append(current)
        if current is None:
            out.append(line); continue
        if header:
            out.append(line)
            if not mask.rstrip().endswith('_'):
                out.append(f'  FEMProfileEnter {ids[current]}\n'); header=False
            continue
        if re.match(r'\s*End\s+(Sub|Function)\b',mask,re.I):
            out.append(f'  FEMProfileLeave {ids[current]}\n')
            out.append(line); current=None; continue
        for match in reversed(list(re.finditer(r'\bExit\s+(Sub|Function)\b',mask,re.I))):
            line=line[:match.start()]+f'FEMProfileLeave {ids[current]}: '+line[match.start():]
        out.append(line)
    if current is not None:
        raise ValueError(f'Unterminated procedure {current}')
    return ''.join(out),found

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--src',required=True)
    parser.add_argument('--dst',required=True)
    args=parser.parse_args()
    src=Path(args.src).resolve(); dst=Path(args.dst).resolve()
    if dst.exists() or src==dst or src in dst.parents:
        raise ValueError('Use a new disposable output directory')
    ids={name:i+1 for i,name in enumerate(name for values in TARGETS.values() for name in values)}
    prepared={}; report={}
    for filename,names in TARGETS.items():
        text,found=inject((src/filename).read_text(encoding='utf-8-sig'),names,ids)
        if set(found)!=set(names):
            raise ValueError(f'{filename}: requested {names}, found {found}')
        prepared[filename]=text; report[filename]=found
    shutil.copytree(src,dst)
    for filename,text in prepared.items():
        (dst/filename).write_bytes(text.replace('\r\n','\n').encode('utf-8'))
    mod=Path(__file__).with_name('FEMIncoProfile.bas').read_text(encoding='utf-8-sig')
    cases='\n'.join(f'      Case {number}: FEMProfileName = "{name}"' for name,number in ids.items())
    (dst/'FEMIncoProfile.bas').write_bytes(mod.replace("      ' PROFILE_NAMES",cases).encode('utf-8'))
    (dst/'profile_injection.json').write_text(json.dumps({'ids':ids,'files':report},indent=2),encoding='utf-8')
    print(json.dumps({'dst':str(dst),'procedures':len(ids)},indent=2))

if __name__=='__main__':
    main()
