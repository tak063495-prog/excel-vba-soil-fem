param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$OutputDirectory,[string]$NativeDll='',[switch]$LegacyDll)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$Book=(Resolve-Path -LiteralPath $Book).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve prior evidence.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$copy=Join-Path $OutputDirectory 'dispatch.xlsm'
Copy-Item -LiteralPath $Book -Destination $copy
if(-not $NativeDll){$NativeDll=Join-Path (Split-Path $Book) 'soilfem_native_x64.dll'}
Copy-Item -LiteralPath $NativeDll -Destination (Join-Path $OutputDirectory 'soilfem_native_x64.dll')
Copy-Item -LiteralPath (Join-Path (Split-Path $Book) 'soilfem_mat_x64.dll') -Destination $OutputDirectory
$harness=@'
Public Function DispatchTest(ByVal enabled As Long, ByVal precision As String) As Variant
 Dim rp(0 To 2) As Long, ci(0 To 3) As Long, a(0 To 3) As Double, ilu(0 To 3) As Double
 Dim rhs(0 To 1) As Double, x(0 To 1) As Double, ev() As Double
 Dim rc As Long, residual As Double, iterations As Long, events As Long, ws As Worksheet, r As Long
 NumberOfElement=2:ReDim Elem(0 To 1)
 FEMNativeSetMode "AUTO"
 FEMWriteTextSetting "NATIVE_PRECISION",precision
 If enabled < 0 Then
  Set ws=ThisWorkbook.Worksheets("設定")
  For r=1 To ws.Cells(ws.Rows.Count,5).End(xlUp).Row
   If UCase$(CStr(ws.Cells(r,5).Value2))="NATIVE_GMRES_SIMD" Then ws.Cells(r,5).ClearContents:ws.Cells(r,3).ClearContents
  Next r
 Else
  FEMWriteNumericSetting "NATIVE_GMRES_SIMD",CDbl(enabled)
 End If
 FEMInvalidateSettingCache
 ReDim P6CSRDiagonalPosition(0 To 1)
 P6CSRDiagonalPosition(0)=0:P6CSRDiagonalPosition(1)=3
 rp(0)=0:rp(1)=2:rp(2)=4
 ci(0)=0:ci(1)=1:ci(2)=0:ci(3)=1
 a(0)=4#:a(1)=1#:a(2)=1#:a(3)=3#
 ilu(0)=4#:ilu(1)=1#:ilu(2)=0.25:ilu(3)=2.75
 rhs(0)=1#:rhs(1)=2#:x(0)=987#:x(1)=987#
 rc=FEMKernGmres(2,rp,ci,a,ilu,rhs,x,2,20,0.000000000001,residual,iterations,ev,events)
 DispatchTest=Array(rc,residual,iterations,events,x(0),x(1),FEMNativeGmresCalls,FEMReadSetting("NATIVE_GMRES_SIMD",-999#),FEMNativeStatusText(),FEMMatStatusText())
End Function
'@
$app=$null;$wb=$null;$component=$null;$results=@()
try{
 $app=New-Object -ComObject Excel.Application
 $app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $wb=$app.Workbooks.Open($copy,0,$false)
 $component=$wb.VBProject.VBComponents.Add(1);$component.Name='DispatchHarness';$component.CodeModule.AddFromString($harness)
 foreach($case in @(@(-1,'AUTO'),@(0,'AUTO'),@(1,'STRICT'),@(1,'SSE2'),@(1,'AUTO'))){
  $v=@($app.Run("'"+$wb.Name+"'!DispatchTest",[int]$case[0],[string]$case[1]))
  $expected=if(-not $LegacyDll -and $case[0] -eq 1 -and $case[1] -ne 'STRICT'){1}else{0}
  if([int]$v[0] -ne 1 -or [double]$v[1] -gt 1e-12 -or [int]$v[6] -ne $expected){throw ('Dispatch mismatch: '+($case -join '/')+' '+($v -join '/'))}
  if([Math]::Abs([double]$v[4]-1.0/11) -gt 1e-13 -or [Math]::Abs([double]$v[5]-7.0/11) -gt 1e-13){throw 'Independent solution mismatch.'}
  if($case[0] -eq -1 -and [double]$v[7] -ne -999){throw 'Absent-key test did not remove the setting.'}
  $results+=@{setting=$case[0];precision=$case[1];rc=$v[0];residual=$v[1];iterations=$v[2];events=$v[3];x=@($v[4],$v[5]);simd_calls=$v[6];read_setting=$v[7];native_status=$v[8]}
 }
 $results|ConvertTo-Json -Depth 5|Set-Content (Join-Path $OutputDirectory 'result.json') -Encoding UTF8
 Write-Output ('PASS: 5 GMRES dispatch cases; legacy DLL='+[bool]$LegacyDll)
 $wb.Close($false)
}finally{
 if($null -ne $component){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($component)}
 if($null -ne $wb){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}
 if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
