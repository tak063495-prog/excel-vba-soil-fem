param([Parameter(Mandatory=$true)][string]$SourceWorkbook,
 [Parameter(Mandatory=$true)][string]$MatrixPath,
 [Parameter(Mandatory=$true)][string]$VectorsPath,
 [Parameter(Mandatory=$true)][string]$OutputRoot,
 [int]$DofCount=1282,[int]$Bandwidth=105,
 [switch]$RequireRefinement,[switch]$RequireMultiplePivots)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputRoot|Out-Null
$before=(Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash
$tmp=Join-Path ([IO.Path]::GetTempPath()) ('matrix_replay_'+[guid]::NewGuid().ToString('N')+'.xlsm')
Copy-Item -LiteralPath $SourceWorkbook -Destination $tmp
$xl=New-Object -ComObject Excel.Application
$xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1
$wb=$null
try {
 $wb=$xl.Workbooks.Open($tmp,0,$false)
 $cm=$wb.VBProject.VBComponents.Item('FEMSolver').CodeModule
 $s=$cm.Lines(1,$cm.CountOfLines)
 $s=$s.Replace('Option Explicit',"Option Explicit`r`nPrivate mReplayPivots As Long, mReplayRefinements As Long")
 $s=$s.Replace('  P6SolveNonsymmetricBand = False',"  P6SolveNonsymmetricBand = False`r`n  mReplayPivots = 0: mReplayRefinements = 0")
 $s=$s.Replace('    pivotRows(k) = pivotRow',"    pivotRows(k) = pivotRow`r`n    If pivotRow <> k Then mReplayPivots = mReplayPivots + 1")
 $s=$s.Replace('      For refineIter = 1 To 5',"      For refineIter = 1 To 5`r`n        mReplayRefinements = mReplayRefinements + 1")
 $cm.DeleteLines(1,$cm.CountOfLines);$cm.AddFromString($s)
 $cm.AddFromString(@'
Public Function MatrixReplay(ByVal matrixFile As String, ByVal vectorFile As String, ByVal n As Long, ByVal bw As Long) As Variant
  Dim f As Integer, line As String, a As Variant, i As Long, row As Long, pos As Long, ok As Boolean, norm As Double, normB As Double, rv() As Double
  On Error GoTo Failed
  SetP0SilentMode True
  nDof=n: lastDof=n-1: BandWidth=bw: P6KrylovLast=lastDof
  P6SolverMemoryLimitBytes=0#: P6IterativeTolerance=0.00000001: P6CSRNNZ=0
  f=FreeFile: Open matrixFile For Input As #f: Line Input #f,line
  Do While Not EOF(f): Line Input #f,line: P6CSRNNZ=P6CSRNNZ+1: Loop
  Close #f
  ReDim P6CSRRowPtr(0 To n): ReDim P6CSRColumnIndex(0 To P6CSRNNZ-1): ReDim P6CSRValues(0 To P6CSRNNZ-1)
  ReDim Force(0 To lastDof): ReDim Disp(0 To lastDof): ReDim rv(0 To lastDof)
  f=FreeFile: Open matrixFile For Input As #f: Line Input #f,line
  pos=0: row=0
  Do While Not EOF(f)
    Line Input #f,line: a=Split(line,",")
    Do While row<CLng(a(0)): row=row+1: P6CSRRowPtr(row)=pos: Loop
    P6CSRColumnIndex(pos)=CLng(a(1)): P6CSRValues(pos)=CDbl(a(2)): pos=pos+1
  Loop
  Close #f: P6CSRRowPtr(n)=pos
  f=FreeFile: Open vectorFile For Input As #f: Line Input #f,line
  Do While Not EOF(f): Line Input #f,line: a=Split(line,","): Force(CLng(a(0)))=CDbl(a(1)): Loop
  Close #f
  ok=P6SolveNonsymmetricBand()
  norm=P6CSRMatVecResidualNorm(Disp,Force,rv,P6CSRValues)
  For i=0 To lastDof: normB=normB+Force(i)^2: Next i
  MatrixReplay=Array(ok,P6IterativeLastResidual,norm/Sqr(normB),P6BandRejectReason,mReplayPivots,mReplayRefinements): Exit Function
Failed:
  MatrixReplay=Array(False,1E+100,1E+100,CStr(Err.Number) & ":" & Err.Description,0,0)
  On Error Resume Next: Close #f
End Function
'@)
 $r=@($xl.Run("'"+$wb.Name+"'!MatrixReplay",[IO.Path]::GetFullPath($MatrixPath),[IO.Path]::GetFullPath($VectorsPath),$DofCount,$Bandwidth))
 $result=[ordered]@{status='FAIL';source_sha256=$before;matrix_sha256=(Get-FileHash $MatrixPath).Hash;vector_sha256=(Get-FileHash $VectorsPath).Hash;dof=$DofCount;bandwidth=$Bandwidth;solver_ok=[bool]$r[0];reported_residual=$r[1];independent_csr_residual=$r[2];reason=$r[3];pivots=$r[4];refinement_iterations=$r[5]}
 if([bool]$r[0] -and [double]$r[2] -le 1e-8){$result.status='PASS'}
 if($RequireRefinement -and $r[5] -lt 1){$result.status='FAIL'}
 if($RequireMultiplePivots -and $r[4] -lt 2){$result.status='FAIL'}
 [IO.File]::WriteAllText((Join-Path $OutputRoot 'check_failure_matrix.json'),($result|ConvertTo-Json -Depth 4),[Text.Encoding]::UTF8)
 if($result.status -ne 'PASS'){throw ($result|ConvertTo-Json -Compress)}
 Write-Output ('PASS failure matrix native replay: residual='+$r[2])
} finally {
 if($wb){$wb.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}
if((Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash -ne $before){throw 'Source workbook changed'}
