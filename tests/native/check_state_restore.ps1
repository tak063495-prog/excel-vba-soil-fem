param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be idle.'}
if(Test-Path -LiteralPath $OutputRoot){throw 'Preserve existing evidence.'}
$SourceWorkbook=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$OutputRoot=[IO.Path]::GetFullPath($OutputRoot)
New-Item -ItemType Directory -Path $OutputRoot | Out-Null
$copy=Join-Path $OutputRoot 'state_restore_test.xlsm'
Copy-Item -LiteralPath $SourceWorkbook -Destination $copy
foreach($dll in 'soilfem_native_x64.dll','soilfem_mat_x64.dll'){
 Copy-Item -LiteralPath (Join-Path (Split-Path $SourceWorkbook) $dll) -Destination $OutputRoot
}
$harness=@'
Private Declare PtrSafe Function StateCompareMemory Lib "ntdll.dll" Alias "RtlCompareMemory" (ByVal a As LongPtr, ByVal b As LongPtr, ByVal size As LongPtr) As LongPtr

Public Function StateRestoreHarness(ByVal countElements As Long, ByVal style As Long, ByVal route As String) As Variant
 Dim oldElem() As Element_Data, referenceElem() As Element_Data
 Dim oldActive() As Boolean, referenceActive() As Boolean
 Dim i As Long, j As Long, k As Long, gp As Long, oldCount As Long, referenceCount As Long
 Dim settingsSheet As Worksheet, settingRow As Long
 Dim recordBytes As LongPtr, same As Boolean, nativeCalls As Long
 NumberOfElement=countElements:lastDof=0:P6MixedUP=False
 P3StateWorkspaceElemN=-1:P3StateWorkspaceDofN=-2
 ReDim Elem(0 To countElements-1)
 P3EnsureStateWorkspace
 If style=3 Then ReDim P3CommittedPlasticStrain(0 To 4,0 To 3,0 To countElements-1)
 If style=1 Then
  ReDim P3ElementActive(1 To 3)
  P3ElementActive(1)=False:P3ElementActive(2)=True:P3ElementActive(3)=False
 Else
  ReDim P3ElementActive(0 To countElements-1)
  For k=0 To countElements-1:P3ElementActive(k)=(k Mod 2=0):Next k
 End If
 P3ElementActiveReady=(style<>2)
 P3ActivePlasticPointCount=0
 For k=0 To countElements-1
  Elem(k).area=111#+k:Elem(k).kmat(0,0)=4321#+k:Elem(k).IsJoint=(k Mod 3=0)
  For gp=0 To 3
   For i=0 To 16
    Elem(k).mStmat(i,gp)=100000#*k+1000#*gp+i
    Elem(k).Stmat(i,gp)=-Elem(k).mStmat(i,gp)-1#
   Next i
   For i=0 To 3
    P3CommittedPlasticStrain(i,gp,k)=200000#*k+1000#*gp+i
    Elem(k).P2PlasticStrain(i,gp)=-1#-P3CommittedPlasticStrain(i,gp,k)
   Next i
   P3CommittedPlasticMultiplier(gp,k)=30000#+100#*k+gp
   P3CommittedYieldFunction(gp,k)=40000#+100#*k+gp
   P3CommittedYielded(gp,k)=((gp+k) Mod 3<>0)
   P3CommittedSpmatValid(gp,k)=((gp+k) Mod 2=0)
   Elem(k).P2PlasticMultiplier(gp)=-10#:Elem(k).P2YieldFunction(gp)=-20#
   Elem(k).P2Yielded(gp)=Not P3CommittedYielded(gp,k)
   Elem(k).SpmatValid(gp)=Not P3CommittedSpmatValid(gp,k)
   P3ActivePlasticPoint(gp,k)=((gp+k) Mod 2<>0)
   If P3ActivePlasticPoint(gp,k) Then P3ActivePlasticPointCount=P3ActivePlasticPointCount+1
   For i=0 To 2
    P3CommittedPrincipalStress(i,gp,k)=50000#+1000#*k+10#*gp+i
    Elem(k).P2PrincipalStress(i,gp)=-30#
    For j=0 To 15
     P3CommittedSpmat(i,j,gp,k)=60000#+1000#*k+100#*gp+10#*j+i
     Elem(k).Spmat(i,j,gp)=-40#
    Next j
    For j=0 To 2
     P3CommittedDpmat(i,j,gp,k)=70000#+1000#*k+100#*gp+10#*j+i
     Elem(k).Dpmat(i,j,gp)=-50#
    Next j
   Next i
  Next gp
 Next k
 ReDim oldElem(0 To countElements-1):ReDim referenceElem(0 To countElements-1)
 For k=0 To countElements-1:oldElem(k)=Elem(k):Next k
 oldActive=P3ActivePlasticPoint:oldCount=P3ActivePlasticPointCount
 FEMNativeSetMode route
 FEMWriteNumericSetting "NATIVE_STATE_COPY",0
 P3RestoreCommittedMaterialState False
 For k=0 To countElements-1:referenceElem(k)=Elem(k):Next k
 referenceActive=P3ActivePlasticPoint:referenceCount=P3ActivePlasticPointCount
 For k=0 To countElements-1:Elem(k)=oldElem(k):Next k
 P3ActivePlasticPoint=oldActive:P3ActivePlasticPointCount=oldCount
 FEMWriteNumericSetting "NATIVE_STATE_COPY",1
 If style=4 Then
  Set settingsSheet=ThisWorkbook.Worksheets("設定")
  For settingRow=1 To settingsSheet.Cells(settingsSheet.Rows.Count,5).End(xlUp).Row
   If UCase$(CStr(settingsSheet.Cells(settingRow,5).Value2))="NATIVE_STATE_COPY" Then
    settingsSheet.Cells(settingRow,5).ClearContents:settingsSheet.Cells(settingRow,3).ClearContents
   End If
  Next settingRow
  FEMInvalidateSettingCache
  If FEMReadSetting("NATIVE_STATE_COPY",-999#)<>-999# Then Err.Raise vbObjectError+6100,"StateRestoreHarness","setting still present"
 End If
 P3RestoreCommittedMaterialState False
 recordBytes=LenB(Elem(0)):same=True
 For k=0 To countElements-1
  If StateCompareMemory(VarPtr(Elem(k)),VarPtr(referenceElem(k)),recordBytes)<>recordBytes Then same=False
 Next k
 If StateCompareMemory(VarPtr(P3ActivePlasticPoint(0,0)),VarPtr(referenceActive(0,0)),8*countElements)<>8*countElements Then same=False
 If P3ActivePlasticPointCount<>referenceCount Then same=False
 nativeCalls=FEMNativeStateCopies
 StateRestoreHarness=Array(same,nativeCalls,referenceCount,P3ActivePlasticPointCount,CLng(recordBytes))
End Function
'@
$app=$null;$book=$null;$results=[Collections.Generic.List[object]]::new()
try{
 $app=New-Object -ComObject Excel.Application
 $app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $book=$app.Workbooks.Open($copy,0,$false)
 $book.VBProject.VBComponents.Item('FEMEngine').CodeModule.AddFromString($harness)
 $compile=$app.VBE.CommandBars.FindControl(1,578)
 if($compile.Enabled){$compile.Execute()}
 if($app.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Integration harness did not compile.'}
 foreach($n in 1,2,7){foreach($style in 0,1,2,3,4){foreach($route in 'AUTO','OFF'){
  $v=@($app.Run("'"+$book.Name+"'!StateRestoreHarness",$n,$style,$route))
  $expectedCalls=if($route -eq 'AUTO' -and $style -ne 3 -and $style -ne 4){1}else{0}
  if(-not [bool]$v[0] -or [int]$v[1] -ne $expectedCalls){throw ('State mismatch/fallback: '+$n+'/'+$style+'/'+$route+'/'+($v -join ','))}
  $results.Add(@{n=$n;style=$style;route=$route;all_udt_bytes_equal=[bool]$v[0];native_calls=[int]$v[1];reference_active_count=[int]$v[2];actual_active_count=[int]$v[3];element_bytes=[int]$v[4]})
 }}}
 $results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputRoot 'result.json') -Encoding UTF8
 Write-Output ('PASS: '+$results.Count+' real VBA UDT/full-state comparisons')
}finally{
 if($null -ne $book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
