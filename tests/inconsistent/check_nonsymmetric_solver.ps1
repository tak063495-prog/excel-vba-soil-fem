param(
  [Parameter(Mandatory=$true)][string]$SourceWorkbook,
  [Parameter(Mandatory=$true)][string]$OutputRoot
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$out = Join-Path $OutputRoot 'check_nonsymmetric_solver.json'
New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null
$temp = Join-Path ([IO.Path]::GetTempPath()) ('ns_solver_' + [guid]::NewGuid().ToString('N') + '.xlsm')
$sourceHashBefore = (Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash
Copy-Item -LiteralPath $SourceWorkbook -Destination $temp -Force
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
$wb = $null; $cm = $null
try {
  $wb = $xl.Workbooks.Open($temp, 0, $false)
  $cm = $wb.VBProject.VBComponents.Item('FEMSolver').CodeModule
  $cm.AddFromString(@'
Public Function NSHarness(ByVal caseId As Long, ByVal callPublic As Boolean) As Variant
  On Error GoTo HarnessFailed
  SetP0SilentMode True
  Dim i As Long, j As Long, p As Long, n As Long, x As Variant, b() As Double, rp() As Long, ci() As Long, a As Variant, r(0 To 3) As Variant
  Select Case caseId
    Case 1: n = 1: a = Array(2#): x = Array(3#)
    Case 2: n = 2: a = Array(1#, 2#, 3#, 1#): x = Array(2#, -1#)
    Case 3: n = 4: a = Array(4#, 1#, 0#, 0#, -2#, 5#, 1#, 0#, 0#, 2#, 3#, 1#, 1#, 0#, -1#, 4#): x = Array(1#, 2#, -1#, 3#)
    Case 4: n = 3: a = Array(2#, 1#, 0#, 0#, 3#, 1#, 1#, 0#, 4#): x = Array(0#, 0#, 0#)
    Case 5: n = 2: a = Array(1#, 2#, 2#, 4#): x = Array(1#, 1#)
    Case 6: n = 3: a = Array(1E9, 2#, 0#, 0#, 1E9, 3#, 4#, 0#, 1E9): x = Array(0.5, -0.25, 2#)
    Case 7: n = 2: a = Array(2#, 1#, 0#, 3#): x = Array(2#, -1#)
    Case 8: n = 2: a = Array(2#, 1#, 0#, 3#): x = Array(2E-9, -1E-9)
    Case 9: n = 5: a = Array(0.01, 2#, 0#, 0#, 0#, 3#, 4#, 1#, 0#, 0#, 0#, -2#, 5#, 1#, 0#, 0#, 0#, -3#, 4#, 1#, 0#, 0#, 0#, -1#, 6#): x = Array(1#, 2#, -1#, 3#, -2#)
    Case 10, 11, 12, 13: n = 3: a = Array(4#, 1#, 2#, -2#, 5#, 1#, 1#, -1#, 3#): x = Array(2#, -1#, 1#)
    Case Else: NSHarness = Array(False, 0#, 0#): Exit Function
  End Select
  ReDim b(0 To n - 1): ReDim rp(0 To n): ReDim ci(0 To n * n - 1): ReDim P6CSRValues(0 To n * n - 1): ReDim P6CSRDiagonalPosition(0 To n - 1)
  p = 0: rp(0) = 0
  For i = 0 To n - 1
    For j = 0 To n - 1
      ci(p) = j: If caseId = 4 Then b(i) = 0# Else b(i) = b(i) + a(i*n+j) * x(j)
      P6CSRValues(p) = a(i*n+j): p = p + 1
    Next j
    rp(i + 1) = p
  Next i
  If caseId = 5 Then b(1) = 7#
  nDof = n: lastDof = n - 1: BandWidth = n - 1
  If caseId = 9 Then BandWidth = 1
  ReDim Force(0 To n - 1): ReDim Disp(0 To n - 1)
  For i = 0 To n - 1: Force(i) = b(i): Disp(i) = 77# + i: Next i
  For i = 0 To n - 1: P6CSRDiagonalPosition(i) = i*n+i: Next i
  P6CSRRowPtr = rp: P6CSRColumnIndex = ci: P6CSRNNZ = p: P6CSRValueCapacity = p
  P6CSRReady = True: P6CSRILUReady = False: P6CSRDiagonalInverseReady = False
  P6IterativeTolerance = 1E-10: P6IterativeMaxIterations = 200
  P6GMRESRestart = 3: P6GMRESRestartRequested = 3: P6KrylovLast = n - 1
  If caseId = 7 And callPublic Then P6SolverMemoryLimitBytes = 1# Else P6SolverMemoryLimitBytes = 100000000#
  P6IterativeLastResidual = 0#: P6IterativeLastIterations = 0
  P6MixedUP = False: P6MixedCoupledKrylov = False: P6CSRStorageBytes = 0#
  If caseId >= 10 Then
    ReDim NodeCond(0 To n - 1): NodeCond(0) = 1: Disp(0) = 2#
    ReDim P6CSROriginalValues(0 To n*n - 1)
    For i = 0 To n*n - 1: P6CSROriginalValues(i) = P6CSRValues(i): Next i
    P6UseCSR = True: P6CSRBoundaryApplied = False: P6CSRBoundaryConstraintVersion = -1
    P6InvalidateCSRConstraintCache
    If caseId = 13 Then NodeCond(0) = 0
    SetBoundaryCondition
    If caseId >= 11 Then
      ' Populate old extraction, then change the actual stage support mask.
      NodeCond(0) = 0: NodeCond(2) = 0
      If caseId = 11 Or caseId = 13 Then NodeCond(2) = 1
      If caseId = 13 Then NodeCond(0) = 1
      For i = 0 To n - 1: Force(i) = b(i): Disp(i) = 0#: Next i
      If NodeCond(0) <> 0 Then Disp(0) = 2#
      If NodeCond(2) <> 0 Then Disp(2) = 1#
      P3BumpConstraintGeneration
      SetBoundaryCondition
    End If
  End If
  If callPublic Then r(0) = P6SolveNonsymmetricBand() Else r(0) = P6SolveCSRGMRES()
  r(1) = P6IterativeLastResidual: r(2) = Disp: r(3) = Force: NSHarness = r
  Exit Function
HarnessFailed:
  NSHarness = Array("ERROR", CStr(Err.Number) & " " & Err.Description)
End Function
'@)
  $cases = @(
    @{id=1; x=@(3.0); A=@(@(2.0)); b=@(6.0); pass=$true},
    @{id=2; x=@(2.0,-1.0); A=@(@(1.0,2.0),@(3.0,1.0)); b=@(0.0,5.0); pass=$true},
    @{id=3; x=@(1.0,2.0,-1.0,3.0); A=@(@(4,1,0,0),@(-2,5,1,0),@(0,2,3,1),@(1,0,-1,4)); b=@(6,7,4,14); pass=$true},
    @{id=4; x=@(0,0,0); A=@(@(2,1,0),@(0,3,1),@(1,0,4)); b=@(0,0,0); pass=$true},
    @{id=5; x=@(1,1); A=@(@(1,2),@(2,4)); b=@(3,7); pass=$false},
    @{id=6; x=@(0.5,-0.25,2); A=@(@(1e9,2,0),@(0,1e9,3),@(4,0,1e9)); b=@(499999999.5,-249999994,2000000002); pass=$true},
    @{id=7; x=@(2,-1); A=@(@(2,1),@(0,3)); b=@(3,-3); pass=$false},
    @{id=8; x=@(2e-9,-1e-9); A=@(@(2,1),@(0,3)); b=@(3e-9,-3e-9); pass=$true},
    @{id=9; x=@(1,2,-1,3,-2); A=@(@(0.01,2,0,0,0),@(3,4,1,0,0),@(0,-2,5,1,0),@(0,0,-3,4,1),@(0,0,0,-1,6)); b=@(4.01,10,-6,13,-15); pass=$true},
    @{id=10; x=@(2,-1,1); A=@(@(1,0,0),@(0,5,1),@(0,-1,3)); b=@(2,-4,4); pass=$true},
    @{id=11; x=@(2,-1,1); A=@(@(4,1,0),@(-2,5,0),@(0,0,1)); b=@(7,-9,1); pass=$true},
    @{id=12; x=@(2,-1,1); A=@(@(4,1,2),@(-2,5,1),@(1,-1,3)); b=@(9,-8,6); pass=$true},
    @{id=13; x=@(2,-1,1); A=@(@(1,0,0),@(0,5,0),@(0,0,1)); b=@(2,-5,1); pass=$true}
  )
  $results = @(); $prefix = "'" + $wb.Name + "'!"
  foreach($c in $cases) {
    foreach($public in @($true,$false)) {
      Write-Host ('CASE '+$c.id+' public='+$public)
      $v = @($xl.Run($prefix+'NSHarness',$c.id,$public))
      if([string]$v[0] -ceq "ERROR"){throw $v[1]}
      $ok = [bool]$v[0]; $res = [double]$v[1]; $d = @($v[2]); $dense = 0.0
      for($i=0;$i -lt $c.x.Count;$i++){ $sum=0.0; for($j=0;$j -lt $c.x.Count;$j++){$sum += $c.A[$i][$j]*$d[$j]}; $dense += ($sum-$c.b[$i])*($sum-$c.b[$i]) }
      $dense = [math]::Sqrt($dense); $expected = $c.pass
      if($c.id -eq 7 -and -not $public){$expected=$true}
      if($ok -ne $expected){throw "case $($c.id) $(if($public){'public'}else{'private'}) status $ok expected $expected"}
      $bnorm=[math]::Sqrt(($c.b|ForEach-Object {$_*$_}|Measure-Object -Sum).Sum)
      if($ok -and $dense -gt 1e-9*[math]::Max(1e-30,$bnorm)){throw "case $($c.id) dense residual $dense"}
      if($ok){for($q=0;$q -lt $d.Count;$q++){
        if([math]::Abs($d[$q]-$c.x[$q]) -gt 1e-8*[math]::Max(1e-12,[math]::Abs($c.x[$q]))){throw "case $($c.id) incorrect displacement"}
        if([math]::Abs($v[3][$q]-$c.b[$q]) -gt 1e-10*[math]::Max(1e-30,[math]::Abs($c.b[$q]))){throw "case $($c.id) incorrect boundary RHS"}
      }}
      if(-not $ok){for($q=0;$q -lt $d.Count;$q++){if([math]::Abs($d[$q]-(77+$q)) -gt 1e-12){throw "case $($c.id) changed Disp on failure"}}}
      $results += [ordered]@{case=$c.id; path=$(if($public){'public'}else{'private'}); ok=$ok; iterative_residual=$res; dense_residual=$dense; displacement=$d}
    }
  }
  $json = [ordered]@{source=$SourceWorkbook; source_sha256=$sourceHashBefore; results=$results; status='PASS'} | ConvertTo-Json -Depth 8
  [IO.File]::WriteAllText($out,$json)
}
catch { [IO.File]::WriteAllText($out,([ordered]@{status='FAIL';error=$_.Exception.Message}|ConvertTo-Json)); throw }
finally { if($wb){$wb.Close($false)}; $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null; Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
if((Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash -ne $sourceHashBefore){ throw 'SourceWorkbook changed' }
