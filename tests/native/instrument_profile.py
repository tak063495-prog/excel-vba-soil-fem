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
    "FEMIo.bas": ["P6SolverEvent", "FEMIoAppendCsv", "FEMIoOutDir", "FEMIoFlushEvents", "FEMIoEnsureDir", "FEMIoBufferEvent", "FEMIoNativeGmresEvents", "FEMIoEventPrefix", "FEMIoEventName", "FEMIoEventCsv", "FEMIoCsv", "P6CsvHeaderEvent",
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
    parser.add_argument('--assembly-sections',action='store_true')
    args=parser.parse_args()
    src=Path(args.src).resolve(); dst=Path(args.dst).resolve()
    if dst.exists() or src==dst or src in dst.parents:
        raise ValueError('Use a new disposable output directory')
    targets={filename:list(names) for filename,names in TARGETS.items()}
    sections=['AssemblyMeshPolicy','AssemblySymmetryScan','AssemblyZeroValues','AssemblyCopyValues'] if args.assembly_sections else []
    if args.assembly_sections:
        targets['FEMSolver.bas'] += ['P6SelectSolverMode','P6EnsureBandScatter','P6BuildCSRFromElements','P6AssembleCSRValuesDetailed',
            'P6EnsureCSRScatterList','P6CheckCSRDiagonalValues','P6EnsureCSRValueWorkspace','P6BuildCSRPattern']
        targets['FEMNativeKern.bas'] += ['FEMKernBandScatter','FEMKernCsrScatter']
        if 'Public Function FEMNativeCSRScatterList()' in (src/'FEMNative.bas').read_text(encoding='utf-8-sig'):
            targets.setdefault('FEMNative.bas',[]).append('FEMNativeCSRScatterList')
    ids={name:i+1 for i,name in enumerate([name for values in targets.values() for name in values]+sections)}
    if len(ids)>64: raise ValueError('Profile node capacity exceeded')
    prepared={}; report={}
    for filename,names in targets.items():
        original=(src/filename).read_text(encoding='utf-8-sig')
        if args.assembly_sections and filename=='FEMSolver.bas':
            start=original.index('Sub SetTotalMat()')
            end=original.index('End Sub',start)
            part=original[start:end]
            replacements=[
                (' rcmPolicy = UCase$', f' FEMProfileEnter {ids["AssemblyMeshPolicy"]}\n rcmPolicy = UCase$'),
                ('   matrixMax = 0#', f' FEMProfileLeave {ids["AssemblyMeshPolicy"]}\n FEMProfileEnter {ids["AssemblySymmetryScan"]}\n   matrixMax = 0#'),
                (' P6SelectSolverMode\n', f' FEMProfileLeave {ids["AssemblySymmetryScan"]}\n P6SelectSolverMode\n'),
            ]
            for old,new in replacements:
                if part.count(old)!=1: raise ValueError(f'Assembly marker mismatch: {old}')
                part=part.replace(old,new)
            original=original[:start]+part+original[end:]
            for name,section,first,last in [
                ('P6AssembleCSRValuesDetailed','AssemblyZeroValues','  For position = 0 To P6CSRNNZ - 1','  Next position'),
                ('P6BuildCSRFromElements','AssemblyCopyValues','  For position = 0 To P6CSRNNZ - 1','  Next position'),
            ]:
                start=original.index('Private Function '+name+'(')
                end=original.index('End Function',start)
                part=original[start:end]
                if part.count(first)!=1 or part.count(last)!=1: raise ValueError(f'{name}: section markers')
                part=part.replace(first,f'  FEMProfileEnter {ids[section]}\n'+first)
                part=part.replace(last,last+f'\n  FEMProfileLeave {ids[section]}')
                original=original[:start]+part+original[end:]
        text,found=inject(original,names,ids)
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
