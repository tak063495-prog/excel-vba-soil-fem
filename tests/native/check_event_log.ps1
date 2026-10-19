param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$Book=(Resolve-Path -LiteralPath $Book).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve existing evidence.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$copy=Join-Path $OutputDirectory 'event_test.xlsm'
Copy-Item -LiteralPath $Book -Destination $copy
foreach($dll in 'soilfem_native_x64.dll','soilfem_mat_x64.dll'){Copy-Item -LiteralPath (Join-Path (Split-Path $Book) $dll) -Destination $OutputDirectory}
$harness=@'
Private Declare PtrSafe Function EventLockOpen Lib "kernel32" Alias "CreateFileW" (ByVal path As LongPtr, ByVal access As Long, ByVal share As Long, ByVal security As LongPtr, ByVal creation As Long, ByVal flags As Long, ByVal template As LongPtr) As LongPtr
Private Declare PtrSafe Function EventLockClose Lib "kernel32" Alias "CloseHandle" (ByVal handle As LongPtr) As Long
Public Function EventLogHarness(ByVal enabled As Boolean, ByVal destination As String) As String
 Dim i As Long, text As String, name As String, folder As String, path As String
 Dim lockHandle As LongPtr, expected As String, flag As Boolean, header As String, f As Integer
 On Error GoTo Failed
 NumberOfElement=2:ReDim Elem(0 To 1):FEMNativeSetMode "AUTO"
 mEventFast=enabled:mEvUsed=0:mEvBuf="":mEvCount=0:mEvPath=""
 mEventPrefixValid=False:mEventName="":mEventNameCsv=""
 mIoFast=True:mIoCachedDir="":mIoBookIdentity="":mIoSep=Application.PathSeparator
 femIoMode="STAGE":folder=FEMIoOutDir():path=folder & mIoSep & "solver_events.csv"
 If Dir$(path)<>"" Then Kill path
 For i=0 To 2199
  P3SrmTrialRunning=(i Mod 7<>0):P3SrmLogStageNo=i Mod 3:P3RunLogStageNo=i Mod 5
  P3SrmLogKind=Choose(1+i Mod 3,"","SRM","S,R" & vbCrLf)
  P3RunLogKind=Choose(1+i Mod 3,"LOAD","地盤","L" & vbCr & "F" & vbLf)
  P3CurrentStrengthFactor=Choose(1+i Mod 6,1#,1.0001,1.0004,1.0006,2.2125,-0.0011)
  CurrentIncrement=(i \ 31) Mod 9
  name=Choose(1+(i \ 17) Mod 4,"","GMRES_RESTART","a,b" & vbCrLf,"地盤")
  text=Choose(1+i Mod 5,"","plain", "a,b" & vbCr & vbLf & "c",Chr$(0) & "nul", "地盤," & Chr$(34) & "引用" & Chr$(34))
  If i=511 Or i=1023 Then text=String$(70000,"X") & "終端"
  If i Mod 113=0 Then femIoMode="OFF" Else femIoMode="STAGE"
  P6SolverEvent name,text
 Next i
 Call FEMIoFlushEvents
 If mEvCount<>0 Or mEvUsed<>0 Then Err.Raise vbObjectError+6301,,"flush did not clear logical buffer"
 femIoMode="STAGE"
 ' Long line followed by a short flush must not emit the reserved tail.
 P6SolverEvent "LONG",String$(140000,"Q")
 Call FEMIoFlushEvents
 P6SolverEvent "",""
 Call FEMIoFlushEvents
 ' Retain more than 512 records under a real OS sharing violation, grow, retry.
 lockHandle=EventLockOpen(StrPtr(path),&HC0000000,0,0,3,128,0)
 If lockHandle=-1 Then Err.Raise vbObjectError+6302,,"exclusive lock failed"
 For i=1 To 1100
  FEMIoAppendCsv "solver_events.csv",flag,"ver,stage,kind,fs,inc,event,detail","locked_" & CStr(i) & String$(80,"z")
  expected=expected & "locked_" & CStr(i) & String$(80,"z") & vbCrLf
 Next i
 If mEvCount<>1100 Then Err.Raise vbObjectError+6303,,"locked buffer lost records"
 If enabled Then
  If mEvUsed<>Len(expected) Or Left$(mEvBuf,mEvUsed)<>expected Then Err.Raise vbObjectError+6304,,"invalid capacity/content"
 Else
  If mEvBuf<>expected Then Err.Raise vbObjectError+6304,,"invalid legacy buffer"
 End If
 EventLockClose lockHandle:lockHandle=0
 Call FEMIoFlushEvents
 If mEvCount<>0 Then Err.Raise vbObjectError+6305,,"unlock retry failed"
 ' A second flush must not duplicate a header, long row, or locked records.
 Call FEMIoFlushEvents
 FileCopy path,destination
 EventLogHarness="PASS|2200|1100|" & CStr(Len(mEvBuf))
 Exit Function
Failed:
 EventLogHarness="ERROR|" & CStr(Err.Number) & "|" & Err.Description
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
  $result=[string]$app.Run("'"+$wb.Name+"'!EventLogHarness",[bool]$enabled,$destination)
  if($result -notlike 'PASS|*'){throw "Event test failed: $result"}
  $runs+=@{event_fast=$enabled;result=$result;sha256=(Get-FileHash -LiteralPath $destination).Hash;bytes=(Get-Item -LiteralPath $destination).Length}
 }
 if($runs[0].sha256 -cne $runs[1].sha256 -or $runs[0].bytes -ne $runs[1].bytes){throw 'Event output bytes differ.'}
 $raw=[Text.Encoding]::GetEncoding(28591).GetString([IO.File]::ReadAllBytes((Join-Path $OutputDirectory 'events1.csv')))
 if([regex]::Matches($raw,"`r`n").Count -ne 3283){throw 'Missing or extra event rows.'}
 if([regex]::Matches($raw,'ver,stage,kind,fs,inc,event,detail').Count -ne 1){throw 'Duplicate/missing header.'}
 @{status='PASS';runs=$runs;cases=@('2200 changing contexts and escaped fields','512-row boundaries','70000/140000-character rows','OFF logging','1100 locked rows with capacity growth','unlock retry','empty flush without duplication')} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'result.json') -Encoding UTF8
 Write-Output 'PASS: mixed fields/context/boundaries/long rows/1100-row locked retry, exact event bytes'
 $wb.Close($false)
}finally{
 if($null -ne $module){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($module)}
 if($null -ne $wb){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}
 if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
