param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$OutputDirectory,[string]$NativeDll='',[switch]$LegacyDll)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$Book=(Resolve-Path -LiteralPath $Book).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve prior evidence.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$copy=Join-Path $OutputDirectory 'assembly_test.xlsm'
Copy-Item -LiteralPath $Book -Destination $copy
if(-not $NativeDll){$NativeDll=Join-Path (Split-Path $Book) 'soilfem_native_x64.dll'}
Copy-Item -LiteralPath $NativeDll -Destination (Join-Path $OutputDirectory 'soilfem_native_x64.dll')
Copy-Item -LiteralPath (Join-Path (Split-Path $Book) 'soilfem_mat_x64.dll') -Destination $OutputDirectory
$harness=@'
Private Declare PtrSafe Function AssemblyCompareMemory Lib "ntdll.dll" Alias "RtlCompareMemory" (ByVal a As LongPtr, ByVal b As LongPtr, ByVal size As LongPtr) As LongPtr
Public Function AssemblyHarness(ByVal n As Long, ByVal style As Long, ByVal enabled As Long, ByVal route As String) As Variant
 Dim k As Long, i As Long, j As Long, s As Long, expectedCount As Long, r As Long, beforeReuse As Long
 Dim savedStart() As Long, savedRow() As Long, savedCol() As Long, savedPos() As Long, savedIndex() As Long
 Dim oldElem() As Element_Data, same As Boolean, ws As Worksheet
 NumberOfElement=n:ReDim Elem(0 To n-1):ReDim oldElem(0 To n-1)
 ReDim P6CSRElementPosition(0 To n-1,0 To 15,0 To 15)
 ReDim P3ElementActive(0 To n-1):P3ElementActiveReady=True
 P6CSRNNZ=17:P6RCMCacheSignature="AssemblyHarness":P6TangentGeneration=2:P6MaterialGeneration=3
 For k=0 To n-1
  P3ElementActive(k)=(k Mod 2=0)
  For i=0 To 15:For j=0 To 15
   P6CSRElementPosition(k,i,j)=(k*7+i*3+j) Mod 17
   If style=2 Or (style=1 And (k+i*3+j*7) Mod 5=0) Then Elem(k).kmat(i,j)=-(1#+k*256+i*16+j)
  Next j:Next i
  oldElem(k)=Elem(k)
 Next k
 FEMNativeSetMode route
 FEMWriteNumericSetting "NATIVE_ASSEMBLY",0
 P6CSRScatterElementCount=-1
 If Not P6EnsureCSRScatterList() Then Err.Raise vbObjectError+6201,"AssemblyHarness","reference failed"
 expectedCount=P6CSRScatterCount:savedStart=P6CSRScatterStart
 If expectedCount>0 Then
  savedRow=P6CSRScatterLocalRow:savedCol=P6CSRScatterLocalColumn
  savedPos=P6CSRScatterPosition:savedIndex=P6CSRScatterValueIndex
 End If
 If enabled<0 Then
  Set ws=ThisWorkbook.Worksheets("設定")
  For r=1 To ws.Cells(ws.Rows.Count,5).End(xlUp).Row
   If UCase$(CStr(ws.Cells(r,5).Value2))="NATIVE_ASSEMBLY" Then ws.Cells(r,5).ClearContents:ws.Cells(r,3).ClearContents
  Next r
  FEMInvalidateSettingCache
 Else
  FEMWriteNumericSetting "NATIVE_ASSEMBLY",CDbl(enabled)
 End If
 P6CSRScatterElementCount=-1
 If Not P6EnsureCSRScatterList() Then Err.Raise vbObjectError+6202,"AssemblyHarness","candidate failed"
 same=(P6CSRScatterCount=expectedCount)
 For k=0 To n
  If savedStart(k)<>P6CSRScatterStart(k) Then same=False
 Next k
 For s=0 To expectedCount-1
  If savedRow(s)<>P6CSRScatterLocalRow(s) Or savedCol(s)<>P6CSRScatterLocalColumn(s) Or savedPos(s)<>P6CSRScatterPosition(s) Or savedIndex(s)<>P6CSRScatterValueIndex(s) Then same=False
 Next s
 For k=0 To n-1
  If AssemblyCompareMemory(VarPtr(Elem(k)),VarPtr(oldElem(k)),LenB(Elem(k)))<>LenB(Elem(k)) Then same=False
 Next k
 beforeReuse=FEMNativeAssemblyCalls
 If Not P6EnsureCSRScatterList() Then same=False
 If FEMNativeAssemblyCalls<>beforeReuse Then same=False
 AssemblyHarness=Array(same,FEMNativeAssemblyCalls,expectedCount,P6CSRScatterCount)
End Function
'@
$app=$null;$wb=$null;$results=@()
try{
 $app=New-Object -ComObject Excel.Application
 $app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $wb=$app.Workbooks.Open($copy,0,$false)
 $wb.VBProject.VBComponents.Item('FEMSolver').CodeModule.AddFromString($harness)
 $compile=$app.VBE.CommandBars.FindControl(1,578)
 if($compile.Enabled){$compile.Execute()}
 if($app.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Assembly harness did not compile.'}
 foreach($n in 1,2,20){foreach($style in 0,1,2){foreach($enabled in -1,0,1){foreach($route in 'AUTO','OFF'){
  $v=@($app.Run("'"+$wb.Name+"'!AssemblyHarness",$n,$style,$enabled,$route))
  $expected=if(-not $LegacyDll -and $enabled -eq 1 -and $route -eq 'AUTO'){1}else{0}
  if(-not [bool]$v[0] -or [int]$v[1] -ne $expected){throw ('Assembly mismatch: '+$n+'/'+$style+'/'+$enabled+'/'+$route+' '+($v -join '/'))}
  $results+=@{n=$n;style=$style;enabled=$enabled;route=$route;exact_arrays_and_udt=[bool]$v[0];native_calls=[int]$v[1];count=[int]$v[2]}
 }}}}
 $results|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $OutputDirectory 'result.json') -Encoding UTF8
 Write-Output ('PASS '+$results.Count+' actual VBA array/UDT comparisons; legacy='+$LegacyDll)
}finally{
 if($null -ne $wb){$wb.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}
 if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
