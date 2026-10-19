param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$Book=(Resolve-Path -LiteralPath $Book).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve existing evidence.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$copy=Join-Path $OutputDirectory 'event_replay.xlsm'
Copy-Item -LiteralPath $Book -Destination $copy
foreach($dll in 'soilfem_native_x64.dll','soilfem_mat_x64.dll'){Copy-Item -LiteralPath (Join-Path (Split-Path $Book) $dll) -Destination $OutputDirectory}
$harness=@'
Private Declare PtrSafe Function EventLockOpen Lib "kernel32" Alias "CreateFileW" (ByVal path As LongPtr, ByVal access As Long, ByVal share As Long, ByVal security As LongPtr, ByVal creation As Long, ByVal flags As Long, ByVal template As LongPtr) As LongPtr
Private Declare PtrSafe Function EventLockClose Lib "kernel32" Alias "CloseHandle" (ByVal handle As LongPtr) As Long
Public Function EventReplayHarness(ByVal enabled As Boolean, ByVal destination As String) As String
 Dim events() As Double, original(0 To 23) As Double, codes As Variant, lockHandle As LongPtr, replayState As String
 Dim i As Long, iteration As Long, folder As String, path As String, handled As Boolean
 On Error GoTo Failed
 NumberOfElement=2:ReDim Elem(0 To 1):FEMNativeSetMode "AUTO"
 mEventFast=enabled:mEvUsed=0:mEvBuf="":mEvCount=0:mEvPath=""
 mEventPrefixValid=False:mEventName="":mEventNameCsv=""
 mIoFast=True:mIoCachedDir="":mIoBookIdentity="":mIoSep=Application.PathSeparator
 femIoMode="STAGE":folder=FEMIoOutDir():path=folder & mIoSep & "solver_events.csv"
 If Dir$(path)<>"" Then Kill path
 ReDim events(0 To 23)
 codes=Array(1,2,3,1,4,5,6,99)
 For i=0 To 7
  events(3*i)=codes(i):events(3*i+1)=30*i:events(3*i+2)=CDbl(i+1)/7#
 Next i
 For i=0 To 23:original(i)=events(i):Next i
 P3SrmTrialRunning=True:P3SrmLogStageNo=3:P3SrmLogKind=""
 P3CurrentStrengthFactor=1.23456:CurrentIncrement=5:replayState="ILU0"
 handled=FEMIoNativeGmresEvents(events,8,30,replayState)
 If handled<>enabled Then Err.Raise vbObjectError+6501,,"dispatch mismatch"
 If Not handled Then
  For i=0 To 7
   iteration=CLng(events(3*i+1))
   Select Case CLng(events(3*i))
    Case 1:P6SolverEvent "GMRES_RESTART","restart=" & CStr(iteration \ 30) & ";iterations=" & CStr(iteration) & ";preconditioner=" & replayState & ";true_relres=" & Format$(events(3*i+2),"0.000E+00")
    Case 2:P6SolverEvent "GMRES_BREAKDOWN","reason=nonfinite_true_residual;iterations=" & CStr(iteration)
    Case 3:replayState="ILU0_APPLY_FAILED":P6SolverEvent "GMRES_BREAKDOWN","reason=preconditioner_apply;preconditioner=" & replayState & ";iterations=" & CStr(iteration)
    Case 4:P6SolverEvent "GMRES_BREAKDOWN","reason=nonfinite_krylov_norm;iterations=" & CStr(iteration)
    Case 5:P6SolverEvent "GMRES_BREAKDOWN","reason=zero_givens_rotation;preconditioner=" & replayState & ";iterations=" & CStr(iteration)
    Case 6:P6SolverEvent "GMRES_BREAKDOWN","reason=zero_upper_triangular_diagonal;preconditioner=" & replayState & ";iterations=" & CStr(iteration)
   End Select
  Next i
 End If
 For i=0 To 23:If events(i)<>original(i) Then Err.Raise vbObjectError+6502,,"event array changed"
 Next i
 If replayState<>"ILU0_APPLY_FAILED" Then Err.Raise vbObjectError+6503,,"code-3 status not replayed"
 Call FEMIoFlushEvents
 lockHandle=EventLockOpen(StrPtr(path),&HC0000000,0,0,3,128,0)
 If lockHandle=-1 Then Err.Raise vbObjectError+6507,,"cannot lock batch output"
 ReDim events(0 To 3299)
 For i=0 To 1099
  events(3*i)=1:events(3*i+1)=30*i:events(3*i+2)=CDbl(i+1)/7#
 Next i
 events(15)=3:replayState="ILU0"
 handled=FEMIoNativeGmresEvents(events,1100,30,replayState)
 If Not handled Then
  For i=0 To 1099
   iteration=CLng(events(3*i+1))
   If events(3*i)=3 Then
    replayState="ILU0_APPLY_FAILED"
    P6SolverEvent "GMRES_BREAKDOWN","reason=preconditioner_apply;preconditioner=" & replayState & ";iterations=" & CStr(iteration)
   Else
    P6SolverEvent "GMRES_RESTART","restart=" & CStr(iteration \ 30) & ";iterations=" & CStr(iteration) & ";preconditioner=" & replayState & ";true_relres=" & Format$(events(3*i+2),"0.000E+00")
   End If
  Next i
 End If
 If mEvCount<>1100 Or replayState<>"ILU0_APPLY_FAILED" Then Err.Raise vbObjectError+6508,,"locked batch lost rows/state"
 EventLockClose lockHandle:lockHandle=0
 Call FEMIoFlushEvents
 If mEvCount<>0 Then Err.Raise vbObjectError+6509,,"locked batch retry failed"
 FileCopy path,destination
 If enabled Then
  If Not FEMIoNativeGmresEvents(events,0,30,replayState) Then Err.Raise vbObjectError+6504,,"empty batch not accepted"
  replayState="UNCHANGED":femIoMode="OFF"
  If FEMIoNativeGmresEvents(events,8,30,replayState) Or replayState<>"UNCHANGED" Then Err.Raise vbObjectError+6505,,"OFF dispatch/state changed"
  femIoMode="STAGE":FEMNativeSetMode "OFF"
  If FEMIoNativeGmresEvents(events,8,30,replayState) Or replayState<>"UNCHANGED" Then Err.Raise vbObjectError+6506,,"non-native dispatch/state changed"
 End If
 EventReplayHarness="PASS|codes1-6,unknown,zero,OFF,native-OFF,array,status"
 Exit Function
Failed:
 EventReplayHarness="ERROR|" & CStr(Err.Number) & "|" & Err.Description
 On Error Resume Next
 If lockHandle<>0 And lockHandle<>-1 Then EventLockClose lockHandle
End Function
'@
$app=$null;$wb=$null;$module=$null;$runs=@()
try{
 $app=New-Object -ComObject Excel.Application
 $app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $wb=$app.Workbooks.Open($copy,0,$false)
 $module=$wb.VBProject.VBComponents.Item('FEMIo').CodeModule
 $module.AddFromString($harness)
 $compile=$app.VBE.CommandBars.FindControl(1,578);if($compile.Enabled){$compile.Execute()}
 if($app.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Event harness failed to compile.'}
 foreach($enabled in 0,1){
  $destination=Join-Path $OutputDirectory "events${enabled}.csv"
  $result=[string]$app.Run("'"+$wb.Name+"'!EventReplayHarness",[bool]$enabled,$destination)
  if($result -notlike 'PASS|*'){throw "Event test failed: $result"}
  $runs+=@{event_fast=$enabled;result=$result;sha256=(Get-FileHash -LiteralPath $destination).Hash;bytes=(Get-Item -LiteralPath $destination).Length}
 }
 if($runs[0].sha256 -cne $runs[1].sha256 -or $runs[0].bytes -ne $runs[1].bytes){throw 'Event output bytes differ.'}
 $raw=[Text.Encoding]::GetEncoding(28591).GetString([IO.File]::ReadAllBytes((Join-Path $OutputDirectory 'events1.csv')))
 if([regex]::Matches($raw,"`r`n").Count -ne 1108){throw 'Missing or extra event rows.'}
 if([regex]::Matches($raw,'ver,stage,kind,fs,inc,event,detail').Count -ne 1){throw 'Duplicate/missing header.'}
 @{status='PASS';runs=$runs;cases=@('native event codes1-6','unknown event ignored','zero records','OFF mode fallback','non-native fallback','input array unchanged','code3 preconditioner state','exact CSV bytes','1100 locked batch records and unlock retry')} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'result.json') -Encoding UTF8
 Write-Output 'PASS: native GMRES event codes1-6/unknown/empty/OFF/non-native, exact bytes/status/array'
 $wb.Close($false)
}finally{
 if($null -ne $module){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($module)}
 if($null -ne $wb){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}
 if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
