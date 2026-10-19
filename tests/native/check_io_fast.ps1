param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$Book=(Resolve-Path -LiteralPath $Book).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve existing evidence.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$copy=Join-Path $OutputDirectory 'io_test.xlsm'
Copy-Item -LiteralPath $Book -Destination $copy
foreach($dll in 'soilfem_native_x64.dll','soilfem_mat_x64.dll'){Copy-Item -LiteralPath (Join-Path (Split-Path $Book) $dll) -Destination $OutputDirectory}
$harness=@'
Private Declare PtrSafe Function IoLockOpen Lib "kernel32" Alias "CreateFileW" (ByVal path As LongPtr, ByVal access As Long, ByVal share As Long, ByVal security As LongPtr, ByVal creation As Long, ByVal flags As Long, ByVal template As LongPtr) As LongPtr
Private Declare PtrSafe Function IoLockClose Lib "kernel32" Alias "CloseHandle" (ByVal handle As LongPtr) As Long
Private Sub IoAssert(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError+6200,"IoFastHarness",message
End Sub
Private Function IoBytes(ByVal path As String) As String
 Dim f As Integer, text As String
 f=FreeFile:Open path For Binary Access Read As #f
 text=Space$(LOF(f)):If Len(text)>0 Then Get #f,,text
 Close #f:IoBytes=text
End Function
Public Function IoFastHarness(ByVal enabled As Boolean, ByVal renamedPath As String) As Variant
 Dim folder As String, legacy As String, oldPath As String, newFolder As String, newPath As String
 Dim flag As Boolean, f As Integer, expected As String, pending As String, checks As Long, beginError As Long
 Dim lockHandle As LongPtr
 On Error GoTo HarnessFailed
 NumberOfElement=2:ReDim Elem(0 To 1):FEMNativeSetMode "AUTO"
 femIoMode="STAGE":mEvBuf="":mEvCount=0:mEvPath=""
 mIoCachedDir="":mIoBookIdentity="":mIoSep=Application.PathSeparator
 mIoFast=False:legacy=FEMIoOutDir()
 mIoFast=enabled:folder=FEMIoOutDir()
 IoAssert folder=legacy,"canonical folder differs":checks=checks+1
 If Len(Dir$(folder & Application.PathSeparator & "perf_trace.csv"))>0 Then Kill folder & Application.PathSeparator & "perf_trace.csv"
 RmDir folder
 FEMIoAppendCsv "solver_events.csv",flag,"header","first"
 IoAssert mEvCount=1,"event not buffered":checks=checks+1
 If enabled Then IoAssert Dir$(folder,vbDirectory)="","buffered append touched directory"
 FEMIoFlushEvents
 oldPath=folder & Application.PathSeparator & "solver_events.csv"
 expected="header" & vbCrLf & "first" & vbCrLf
 IoAssert mEvCount=0 And IoBytes(oldPath)=expected,"deleted-folder recovery/content":checks=checks+1
 lockHandle=IoLockOpen(StrPtr(oldPath),&HC0000000,0,0,3,128,0)
 IoAssert lockHandle<>-1,"cannot acquire exclusive OS lock"
 FEMIoAppendCsv "solver_events.csv",flag,"header","second"
 pending=mEvBuf:Call FEMIoFlushEvents
 IoAssert mEvCount=1 And mEvBuf=pending And mEvPath=oldPath,"locked flush lost pending bytes":checks=checks+1
 FEMIoAppendCsv "perf_trace.csv",flag,"trace_header","must_not_overtake"
 IoAssert mEvCount=1 And Dir$(folder & Application.PathSeparator & "perf_trace.csv")="","direct CSV overtook pending events":checks=checks+1
 On Error Resume Next
 FEMIoBegin
 beginError=Err.Number:Err.Clear
 On Error GoTo HarnessFailed
 IoAssert beginError=vbObjectError+6210 And mEvCount=1 And mEvBuf=pending,"new run mixed with pending log":checks=checks+1
 IoAssert IoLockClose(lockHandle)<>0,"OS lock did not close":lockHandle=0
 Call FEMIoFlushEvents
 expected=expected & "second" & vbCrLf
 IoAssert mEvCount=0 And IoBytes(oldPath)=expected,"retry lost/duplicated rows or header":checks=checks+1
 FEMIoAppendCsv "solver_events.csv",flag,"header","old_pending"
 pending=mEvBuf
 lockHandle=IoLockOpen(StrPtr(oldPath),&HC0000000,0,0,3,128,0)
 IoAssert lockHandle<>-1,"cannot acquire second exclusive OS lock"
 ThisWorkbook.SaveAs renamedPath,52
 newFolder=FEMIoOutDir()
 IoAssert newFolder<>folder,"SaveAs did not invalidate cache":checks=checks+1
 FEMIoAppendCsv "solver_events.csv",flag,"header","must_retry"
 IoAssert mEvPath=oldPath And mEvBuf=pending And mEvCount=1,"path change contaminated pending log":checks=checks+1
 IoLockClose lockHandle:lockHandle=0:Call FEMIoFlushEvents
 expected=expected & "old_pending" & vbCrLf
 IoAssert IoBytes(oldPath)=expected,"pending log used new folder":checks=checks+1
 FEMIoAppendCsv "solver_events.csv",flag,"header","new_row":Call FEMIoFlushEvents
 newPath=newFolder & Application.PathSeparator & "solver_events.csv"
 IoAssert IoBytes(newPath)="header" & vbCrLf & "new_row" & vbCrLf,"new path/header/order":checks=checks+1
 Kill newPath:RmDir newFolder
 FEMIoAppendCsv "perf_trace.csv",flag,"trace_header","trace_row"
 IoAssert IoBytes(newFolder & Application.PathSeparator & "perf_trace.csv")="trace_header" & vbCrLf & "trace_row" & vbCrLf,"direct CSV did not recreate folder":checks=checks+1
 IoFastHarness=Array(checks,folder,newFolder,mEvCount)
 Exit Function
HarnessFailed:
 IoFastHarness=Array(-1,Err.Number,Err.Description,mEvCount,mEvBuf,mEvPath,checks)
 On Error Resume Next
 If lockHandle<>0 And lockHandle<>-1 Then IoLockClose lockHandle
End Function
'@
$app=$null;$wb=$null;$module=$null;$result=@()
try{
 $app=New-Object -ComObject Excel.Application
 $app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $wb=$app.Workbooks.Open($copy,0,$false)
 $module=$wb.VBProject.VBComponents.Item('FEMIo').CodeModule
 $module.AddFromString($harness)
 $compile=$app.VBE.CommandBars.FindControl(1,578);if($compile.Enabled){$compile.Execute()}
 if($app.VBE.CommandBars.FindControl(1,578).Enabled){throw 'I/O harness failed to compile.'}
 foreach($enabled in 0,1){
  $target=Join-Path $OutputDirectory "renamed${enabled}.xlsm"
  $v=@($app.Run("'"+$wb.Name+"'!IoFastHarness",[bool]$enabled,$target))
  if([int]$v[0] -ne 12 -or [int]$v[3] -ne 0){throw ('I/O test failed: '+($v -join '|'))}
  $result+=@{enabled=$enabled;checks=[int]$v[0];old_folder=$v[1];new_folder=$v[2];pending_after=[int]$v[3];status='PASS'}
 }
 $result|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $OutputDirectory 'result.json') -Encoding UTF8
 Write-Output 'PASS: 24 actual I/O assertions (OFF/ON, delete, lock, retry, SaveAs, next run, bytes)'
 $wb.Close($false)
}finally{
 if($null -ne $module){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($module)}
 if($null -ne $wb){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}
 if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
