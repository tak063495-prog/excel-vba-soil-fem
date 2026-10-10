param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputRoot|Out-Null
$hash=(Get-FileHash -LiteralPath $SourceWorkbook).Hash
$tmp=Join-Path ([IO.Path]::GetTempPath()) ('cache_test_'+[guid]::NewGuid().ToString('N')+'.xlsm')
Copy-Item -LiteralPath $SourceWorkbook -Destination $tmp
$xl=New-Object -ComObject Excel.Application;$xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1;$wb=$null
try {
 $wb=$xl.Workbooks.Open($tmp,0,$false);$cm=$wb.VBProject.VBComponents.Item('FEMSolver').CodeModule
 $cm.AddFromString(@'
Public Function CacheTest(ByVal stage As Long) As Variant
  Dim a As Variant, x As Variant, i As Long, j As Long, p As Long, ok As Boolean, errValue As Double
  On Error GoTo Failed
  SetP0SilentMode True
  nDof=5: lastDof=4: BandWidth=1: P6KrylovLast=4: P6CSRReady=True
  P6SolverMemoryLimitBytes=0#: P6CSRStorageBytes=0#: P6IterativeTolerance=0.00000001
  ResultStatus=RESULT_NOT_RUN: AnalysisErrorNumber=0
  If stage=1 Then P6ClearElasticRecoveryCSR: P6FactorizationCount=0: P6FactorizationReuseCount=0
  a=Array(0.01,2#,0#,0#,0#,3#,4#,1#,0#,0#,0#,-2#,5#,1#,0#,0#,0#,-3#,4#,1#,0#,0#,0#,-1#,6#)
  x=Array(1#,2#,-1#,3#,-2#)
  If stage=2 Then x=Array(-2#,1#,4#,-1#,0.5)
  If stage>=3 Then a(0)=5.01
  P6CSRNNZ=25: If stage>=4 Then P6CSRNNZ=23
  ReDim P6CSRRowPtr(0 To 5): ReDim P6CSRColumnIndex(0 To P6CSRNNZ-1): ReDim P6CSRValues(0 To P6CSRNNZ-1)
  ReDim Force(0 To 4): ReDim Disp(0 To 4)
  p=0
  For i=0 To 4
    P6CSRRowPtr(i)=p
    For j=0 To 4
      If Not (stage>=4 And i=0 And j>=3) Then
        P6CSRColumnIndex(p)=j: P6CSRValues(p)=a(i*5+j)
        If stage=5 Then P6CSRValues(p)=0#
        p=p+1
      End If
      Force(i)=Force(i)+a(i*5+j)*x(j)
    Next j
  Next i
  P6CSRRowPtr(5)=p
  If stage=6 Then P6SolverMemoryLimitBytes=1#
  ok=P6SolveElasticRecoveryCSR()
  For i=0 To 4: If Abs(Disp(i)-x(i))>errValue Then errValue=Abs(Disp(i)-x(i))
  Next i
  CacheTest=Array(ok,P6IterativeLastResidual,errValue,P6FactorizationCount,P6FactorizationReuseCount,ResultStatus): Exit Function
Failed: CacheTest=Array(False,1E+100,1E+100,-1,-1,CStr(Err.Number) & ":" & Err.Description)
End Function
'@)
 $records=@()
 foreach($stage in 1..6){
  $r=@($xl.Run("'"+$wb.Name+"'!CacheTest",$stage))
  if($stage -le 4 -and (-not $r[0] -or $r[1] -gt 1e-8 -or $r[2] -gt 1e-10)){throw ('Incorrect cached solve stage '+$stage+': '+($r -join '|'))}
  if($stage -eq 2 -and ($r[3] -ne 1 -or $r[4] -ne 1)){throw 'Repeated RHS was not reused'}
  if($stage -eq 3 -and ($r[3] -ne 2 -or $r[4] -ne 1)){throw 'Changed coefficient was reused'}
  if($stage -eq 4 -and ($r[3] -ne 3 -or $r[4] -ne 1)){throw 'Changed CSR pattern was reused'}
  if($stage -ge 5 -and $r[0]){throw 'Invalid matrix or capacity limit accepted'}
  if($stage -eq 6 -and $r[5] -ne 'CAPACITY_ERROR'){throw 'Capacity status lost'}
  $records+=@{stage=$stage;values=$r}
 }
 @{status='PASS';source_sha256=$hash;records=$records}|ConvertTo-Json -Depth 6|Set-Content (Join-Path $OutputRoot 'check_elastic_recovery_solver.json') -Encoding UTF8
 Write-Output 'PASS cached pivot LU: repeated RHS, coefficient/pattern invalidation, singular and capacity rejection'
} finally {if($wb){$wb.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)};$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue}
if((Get-FileHash -LiteralPath $SourceWorkbook).Hash -ne $hash){throw 'Source changed'}
