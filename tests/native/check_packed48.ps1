param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot,[switch]$WithoutDll)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be idle.'}
$principal48=[Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
if($principal48.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'Normal user token required.'}
if(Test-Path -LiteralPath $OutputRoot){throw 'Preserve completed and failed evidence.'}
$SourceWorkbook=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$OutputRoot=[IO.Path]::GetFullPath($OutputRoot)
$sourceHash48=(Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash
$casesPath48=Join-Path $PSScriptRoot '../inconsistent/regularized_cases.json'
$cases48=@(Get-Content -LiteralPath $casesPath48 -Raw -Encoding utf8 | ConvertFrom-Json -DateKind String)
if($cases48.Count -ne 28){throw 'Expected 28 independent regularized cases.'}
New-Item -ItemType Directory -Path $OutputRoot | Out-Null
$copy48=Join-Path $OutputRoot 'packed48_test.xlsm';Copy-Item -LiteralPath $SourceWorkbook -Destination $copy48
if(-not $WithoutDll){foreach($dll48 in 'soilfem_native_x64.dll','soilfem_mat_x64.dll'){Copy-Item -LiteralPath (Join-Path (Split-Path $SourceWorkbook) $dll48) -Destination $OutputRoot}}
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class Packed48TestIdentity {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint processId);
}
'@
$harness48=@'
Public Function Packed48Harness(ByVal n As Long, ByVal mu As Double, ByVal bw As Long, ByVal budget As Double, ByVal route As String) As Variant
 On Error GoTo Failed48
 SetP0SilentMode True
 Dim i As Long, j As Long, p As Long, ok As Boolean, same As Boolean, sheet48 As Worksheet
 Set sheet48=ThisWorkbook.Worksheets("packed48_test")
 nDof=n:lastDof=n-1:BandWidth=bw:P6CSRNNZ=n*n:P6CSRStorageBytes=0#
 NumberOfElement=1:ReDim Elem(0)
 ReDim P6CSRRowPtr(n):ReDim P6CSRColumnIndex(n*n-1):ReDim P6CSRValues(n*n-1):ReDim P6CSROriginalValues(n*n-1)
 ReDim Force(n-1):ReDim Disp(n-1):ReDim NodeCond(n-1)
 For i=0 To n-1
  P6CSRRowPtr(i)=i*n:Force(i)=sheet48.Cells(i+1,n+1).Value2:NodeCond(i)=sheet48.Cells(i+1,n+2).Value2:Disp(i)=77#+i
  For j=0 To n-1
   p=i*n+j:P6CSRColumnIndex(p)=j:P6CSROriginalValues(p)=sheet48.Cells(i+1,j+1).Value2:P6CSRValues(p)=P6CSROriginalValues(p)
  Next j
 Next i
 P6CSRRowPtr(n)=n*n:P6CSRReady=True:P6SolverMemoryLimitBytes=budget:P6IterativeTolerance=1E-8
 P6IterativeLastResidual=.123:P6IterativeLastIterations=47:P6FactorizationCount=0
 P3UserCancel=False:P3SearchFatal=False
 FEMKernResetCounts
 If route="LEGACY" Then
  FEMNativeSetMode "OFF"
  ok=P6SolveResidualRecoveryLegacy48(mu)
 Else
  FEMNativeSetMode route
  ok=P6SolveResidualRecovery(mu)
 End If
 same=True
 For p=0 To n*n-1
  If P6CSRValues(p)<>P6CSROriginalValues(p) Then same=False
 Next p
 Packed48Harness=Array(ok,Disp,same,P6IterativeLastResidual,P6IterativeLastIterations,P6FactorizationCount,FEMKernRegCalls,FEMMatStatusText())
 Exit Function
Failed48:
 Packed48Harness=Array("ERROR",Err.Number,Err.Description,P3UserCancel,P3SearchFatal)
End Function
Public Function Packed48IndexHarness(ByVal n As Long, ByVal bw As Long) As Boolean
 Packed48IndexHarness=P6RecoveryPackedIndexable48(n,bw)
End Function
'@
function Check48($condition48,$message48){if(-not $condition48){throw $message48}}
function Bits48($values48){@($values48 | ForEach-Object {[BitConverter]::DoubleToInt64Bits([double]$_).ToString('X16')})}
function Run48($case48,$route48){
 $n48=$case48.rhs.Count
 for($i48=0;$i48 -lt $n48;$i48++){
  for($j48=0;$j48 -lt $n48;$j48++){$sheet48.Cells.Item($i48+1,$j48+1).Value2=[double]$case48.A[$i48][$j48]}
  $sheet48.Cells.Item($i48+1,$n48+1).Value2=[double]$case48.rhs[$i48]
  $sheet48.Cells.Item($i48+1,$n48+2).Value2=[double]([int]($case48.fixed -contains $i48))
 }
 $v48=@($app48.Run("'"+$book48.Name+"'!Packed48Harness",$n48,[double]$case48.damping,[int]$case48.band,[double]$case48.budget,$route48))
 return ,$v48
}
$app48=$null;$book48=$null;$identity48=$null;$records48=[Collections.Generic.List[object]]::new();$ok48=$false;$failure48=$null
try{
 $app48=New-Object -ComObject Excel.Application
 [uint32]$pid48=0;[void][Packed48TestIdentity]::GetWindowThreadProcessId([IntPtr]$app48.Hwnd,[ref]$pid48)
 $p48=Get-Process -Id $pid48 -ErrorAction Stop
 $identity48=@{pid=$pid48;birth_utc_ticks=$p48.StartTime.ToUniversalTime().Ticks.ToString();path=$p48.Path;binary_sha256=(Get-FileHash -LiteralPath $p48.Path -Algorithm SHA256).Hash}
 $identity48 | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutputRoot 'owned_excel.json') -Encoding utf8
 if(@(Get-Process EXCEL).Count -ne 1){throw 'Foreign Excel appeared; no test dispatched.'}
 $app48.Visible=$false;$app48.EnableEvents=$false;$app48.DisplayAlerts=$false;$app48.AutomationSecurity=1
 $book48=$app48.Workbooks.Open($copy48,0,$false)
 $sheet48=$book48.Worksheets.Add();$sheet48.Name='packed48_test'
 $cm48=$book48.VBProject.VBComponents.Item('FEMSolver').CodeModule
 $production48=$cm48.Lines(1,$cm48.CountOfLines)
 $cm48.AddFromString($harness48)
 $caseIndex48=0
 foreach($case48 in $cases48){
  $caseIndex48++
  $control48=Run48 $case48 'LEGACY'
  Check48 ([string]$control48[0] -cne 'ERROR') ('Legacy exception '+$case48.name)
  foreach($route48 in 'OFF','BAND','MAT','AUTO'){
   $v48=Run48 $case48 $route48
   Check48 ([string]$v48[0] -cne 'ERROR') ('Exception '+$case48.name+'/'+$route48+': '+($v48 -join '|'))
   Check48 ($v48.Count -eq 8 -and @($v48[1]).Count -eq $case48.rhs.Count) 'Invalid test shape'
   Check48 ([bool]$v48[0] -eq [bool]$case48.ok -and [bool]$v48[0] -eq [bool]$control48[0]) 'Status mismatch'
   Check48 (((Bits48 $v48[1]) -join ',') -ceq ((Bits48 $control48[1]) -join ',')) ('Double bit mismatch '+$case48.name+'/'+$route48)
   Check48 ([bool]$v48[2] -and [double]$v48[3] -eq .123 -and [int]$v48[4] -eq 47) 'Primary operator/diagnostic changed'
   Check48 ([int]$v48[5] -eq [int]$control48[5]) 'Factorization counted twice'
   # MAT accelerates material kernels; only AUTO enables Phase C regularized.
   $nativeExpected48=(-not $WithoutDll -and $case48.ok -and $route48 -eq 'AUTO')
   Check48 ([int]$v48[6] -eq [int]$nativeExpected48) ('DLL engagement mismatch '+$route48+': '+$v48[7])
   $gap48=0.;$norm48=0.
   for($i48=0;$i48 -lt $case48.rhs.Count;$i48++){
    $diff48=[double]$v48[1][$i48]-[double]$case48.expected[$i48]
    $gap48+=$diff48*$diff48;$norm48+=[double]$case48.expected[$i48]*[double]$case48.expected[$i48]
   }
   $gap48=[Math]::Sqrt($gap48/[Math]::Max(1e-30,$norm48))
   Check48 ($gap48 -le 2e-7) 'Independent expected solution mismatch'
   $records48.Add(@{case_index=$caseIndex48;name=$case48.name;damping=$case48.damping;route=$route48;accepted=[bool]$v48[0];bits=(Bits48 $v48[1]);relative_gap=$gap48;native_calls=[int]$v48[6]})
  }
 }
 $boundaryCount48=0
 foreach($bound48 in @(@(2147483647,0,$true),@(1073741823,1,$true),@(1073741824,1,$false),@(65535,32767,$true),@(65536,32767,$false),@(46340,46339,$true),@(46341,46340,$false),@(1,0,$true))){
  $fit48=$app48.Run("'"+$book48.Name+"'!Packed48IndexHarness",[int]$bound48[0],[int]$bound48[1])
  Check48 ([bool]$fit48 -eq [bool]$bound48[2]) 'Long boundary mismatch';$boundaryCount48++
 }
 $hook48='  ReDim a(0 To lastDof * stride48 + bw)'
 Check48 ($production48.Split(@($hook48),[StringSplitOptions]::None).Count -eq 2) 'Allocation hook must be unique'
 foreach($fault48 in 7,18){
  $testSource48=$production48.Replace($hook48,('  Err.Raise '+$fault48+"`r`n"+$hook48))
  $cm48.DeleteLines(1,$cm48.CountOfLines);$cm48.AddFromString($testSource48+"`r`n"+$harness48)
  $v48=Run48 $cases48[0] 'OFF'
  if($fault48 -eq 7){
   $control48=Run48 $cases48[0] 'LEGACY'
   Check48 ([string]$v48[0] -cne 'ERROR' -and [bool]$v48[0]) 'Error7 fallback failed'
   Check48 (((Bits48 $v48[1]) -join ',') -ceq ((Bits48 $control48[1]) -join ',')) 'Error7 fallback bit mismatch'
   Check48 ([int]$v48[5] -eq 1 -and [bool]$v48[2] -and [double]$v48[3] -eq .123 -and [int]$v48[4] -eq 47) 'Error7 counter/operator/diagnostic mismatch'
  }else{Check48 ([string]$v48[0] -ceq 'ERROR' -and [int]$v48[1] -eq 18 -and [bool]$v48[3] -and [bool]$v48[4]) 'Error18 was swallowed or retried'}
 }
 $ok48=$true
}catch{$failure48=$_.Exception.ToString();throw}finally{
 foreach($com48 in @($cm48,$sheet48)){if($null -ne $com48 -and [Runtime.InteropServices.Marshal]::IsComObject($com48)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com48)}}
 $cm48=$null;$sheet48=$null
 if($null -ne $book48){$book48.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book48);$book48=$null}
 if($null -ne $app48){$app48.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app48);$app48=$null}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 for($wait48=0;$wait48 -lt 20;$wait48++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){break};Start-Sleep -Milliseconds 500}
 $live48=$false
 if($null -ne $identity48){$p48=Get-Process -Id $identity48.pid -ErrorAction SilentlyContinue;if($null -ne $p48){$live48=($p48.StartTime.ToUniversalTime().Ticks.ToString() -ceq $identity48.birth_utc_ticks)}}
 $record48=@{status=$(if($ok48 -and -not $live48){'PASS'}else{'FAILED_PRESERVED'});failure=$failure48;without_dll=[bool]$WithoutDll;case_count=$cases48.Count;mode_checks=$records48.Count;capacity_boundaries=$boundaryCount48;allocation_error7_fallback=$ok48;error18_propagated=$ok48;source_sha256=$sourceHash48;cases_sha256=(Get-FileHash -LiteralPath $casesPath48 -Algorithm SHA256).Hash;owned_excel_ended=(-not $live48);results=$records48.ToArray()}
 $record48 | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath (Join-Path $OutputRoot 'result.json') -Encoding utf8
}
if($live48){throw 'Owned Excel still present; no further test or forced termination.'}
Check48 ((Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash -ceq $sourceHash48) 'Source workbook changed'
Write-Output ('PASS: '+$records48.Count+' bitwise/independent checks, 8 boundaries, Error7 fallback, Error18 propagation; without_dll='+[bool]$WithoutDll)
