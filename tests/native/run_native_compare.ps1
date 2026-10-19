# -MeshPath: smoke系のfixture。未指定時はリポジトリ追跡の tests/fixtures/smoke_mesh.json（スクリプト基準）。
# 相対パス指定はカレントディレクトリ基準。ブック隣へのフォールバックはしない。
param([string]$Book,[string]$Suite='smoke',[string]$Out,[double]$FixedFs=1.0,[string]$Modes='OFF,AUTO',[string]$Flow='',[string]$MeshPath='',[ValidateSet('STRICT','SSE2','AUTO','AVX2')][string]$Precision='STRICT',[string]$CaseName='',[ValidateSet('off','adaptive')][string[]]$Variants=@('off','adaptive'),[switch]$Profile,[ValidateSet(0,1)][int]$StateCopy=0,[ValidateSet(0,1)][int]$GmresSimd=0,[ValidateSet(0,1)][int]$IoFast=0,[ValidateSet(0,1)][int]$EventFast=0,[ValidateSet(0,1)][int]$AssemblyFast=0,[switch]$NativeProfile,[ValidateSet(0,1)][int]$GmresTuned=0)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $Out){throw "Preserve prior evidence: $Out"}
if(-not (Test-Path -LiteralPath $Book -PathType Leaf)){throw "Book not found: $Book"}
$Book=(Resolve-Path -LiteralPath $Book).Path
$root=Split-Path -Parent $Book
$mesh=$null
if($Suite -like 'smoke*'){
 if($MeshPath -eq ''){$MeshPath=Join-Path $PSScriptRoot '../fixtures/smoke_mesh.json'}
 if(-not (Test-Path -LiteralPath $MeshPath -PathType Leaf)){throw "smoke fixture not found: $([IO.Path]::GetFullPath($MeshPath))"}
 $MeshPath=(Resolve-Path -LiteralPath $MeshPath).Path
 try{$mesh=Get-Content -LiteralPath $MeshPath -Raw -Encoding UTF8|ConvertFrom-Json}catch{throw "smoke fixture is not valid JSON: $MeshPath ($($_.Exception.Message))"}
 if($null -eq $mesh.nodes -or @($mesh.nodes).Count -eq 0){throw "smoke fixture has no nodes: $MeshPath"}
 if($null -eq $mesh.elements -or @($mesh.elements).Count -eq 0){throw "smoke fixture has no elements: $MeshPath"}
 foreach($n in $mesh.nodes){if(@($n).Count -ne 9){throw "smoke fixture node row must have 9 columns: $MeshPath"}}
 foreach($e in $mesh.elements){if(@($e).Count -ne 10){throw "smoke fixture element row must have 10 columns: $MeshPath"}}
 Write-Output "mesh=$MeshPath nodes=$(@($mesh.nodes).Count) elements=$(@($mesh.elements).Count) sha256=$((Get-FileHash -LiteralPath $MeshPath -Algorithm SHA256).Hash)"
}
$before=@(Get-Process EXCEL -ErrorAction SilentlyContinue|%{$_.Id})
if($before.Count -gt 0){throw "Another Excel is running: $before"}
$copy=Join-Path $root ("test_"+$Suite+"_"+(Get-Date -Format yyyyMMdd_HHmmss)+".xlsm")
Copy-Item -LiteralPath $Book -Destination $copy
$harness=@'
Public Function TRun(ByVal modeText As String) As String
  Dim t As Double
  On Error GoTo Failed
  FEMNativeSetMode modeText
  t = Timer
  P0_RunAnalysis
  TRun = CStr(Timer - t) & "|" & ResultStatus & "|" & CStr(AnalysisOK) & "|" & CStr(RelativeResidualFree) & "|" & CStr(P3SuccessfulIncrementCount) & "|" & CStr(P6FactorizationCount) & "|" & CStr(P6FactorizationReuseCount) & "|" & CStr(P3ActivePlasticPointCount) & "|" & FEMNativeStatusText() & "|" & CStr(P6ProfFactorMs) & "|" & CStr(P6ProfSolveMs) & "|" & CStr(P6ProfEvalMs) & "|" & CStr(P6ProfTangentMs) & "|" & AnalysisMessage & "|" & FEMMatStatusText()
  Exit Function
Failed:
  TRun = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function TFos() As String
  Dim ws As Worksheet, r As Long
  On Error GoTo Failed
  Set ws = ThisWorkbook.Worksheets("ステージ結果")
  For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    TFos = TFos & CStr(ws.Cells(r, 2).Value2) & ":" & CStr(ws.Cells(r, 3).Value2) & ":" & CStr(ws.Cells(r, 5).Value2) & ":" & CStr(ws.Cells(r, 10).Value2) & ";"
  Next r
  Exit Function
Failed:
  TFos = "ERR"
End Function
Public Function TDispV() As Variant
  TDispV = TDisp
End Function
Public Function TTunedCalls() As Long
  TTunedCalls=FEMNativeGmresTunedCalls
End Function
Public Function TAssemblyV() As Variant
  TAssemblyV = Array(FEMNativeAssemblyCalls, P6CSRScatterCount, MatrixSymmetryError)
End Function
Public Function TBracketV() As Variant
  TBracketV = Array(P3FosPass, P3FosFail, P3FosMid, P3FosWidth, P3FosBracket, P3SrmTrialCount)
End Function
'@
if($Profile){$harness=$harness.Replace('  FEMNativeSetMode modeText',"  FEMProfileStart`r`n  FEMNativeSetMode modeText")}
if($NativeProfile){
 $harness=@'
Private Declare PtrSafe Sub NativeProfileReset Lib "soilfem_native_x64.dll" Alias "sfn_gmres_profile_reset" ()
Private Declare PtrSafe Function NativeProfileRead Lib "soilfem_native_x64.dll" Alias "sfn_gmres_profile_read" (ByVal data As LongPtr, ByVal n As Long) As Long
Public Function TNativeProfile() As Variant
 Dim d(0 To 31) As Double
 If NativeProfileRead(VarPtr(d(0)),32)<>0 Then Err.Raise vbObjectError+6301,"TNativeProfile","Unbalanced native profiler"
 TNativeProfile=d
End Function
'@ + "`r`n" + $harness.Replace('  FEMNativeSetMode modeText',"  FEMNativeSetMode modeText`r`n  If Not FEMNativeAvailable() Then Err.Raise vbObjectError+6302`r`n  NativeProfileReset")
}
function SetS($wb,$k,$v){$m=$(if($v -is [string]){'FEMWriteTextSetting'}else{'FEMWriteNumericSetting'});$xl.Run("'"+$wb.Name+"'!"+$m,$k,$v)}
function SetRows($wb,$sheet,$rows,$cols){$ws=$wb.Worksheets.Item($sheet);$ws.Range($ws.Cells(2,1),$ws.Cells($ws.UsedRange.Rows.Count,$cols)).ClearContents();$d=[object[,]]::new($rows.Count,$cols);for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $cols;$c++){$d[$r,$c]=$rows[$r][$c]}};$ws.Range($ws.Cells(2,1),$ws.Cells($rows.Count+1,$cols)).Value2=$d}
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class NativeCompareIdentity {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
 public static string NonGmresDigest(string path, out int nonGmres, out int gmres, bool normalize) {
  byte[] data=System.IO.File.ReadAllBytes(path); nonGmres=0; gmres=0;
  using(var kept=new System.IO.MemoryStream()) {
   int start=0;
   for(int end=0;end<=data.Length;end++) {
    if(end<data.Length && data[end]!=10) continue;
    if(end==start) { start=end+1;continue; }
    int e=start,commas=0;
    while(e<end && commas<5) { if(data[e++]==44) commas++; }
    bool isGmres=commas==5 && end-e>=6 && data[e]==71 && data[e+1]==77 && data[e+2]==82 && data[e+3]==69 && data[e+4]==83 && data[e+5]==95;
    if(isGmres) gmres++; else {
     int length=end-start+(end<data.Length?1:0),ee=e;
     while(ee<end && data[ee]!=44) ee++;
     string eventName=System.Text.Encoding.ASCII.GetString(data,e,ee-e);
     if(normalize && eventName=="REGULARIZED_CORRECTION_REQUEST") {
      var encoding=System.Text.Encoding.GetEncoding(28591);
      string line=encoding.GetString(data,start,length);
      line=System.Text.RegularExpressions.Regex.Replace(line,@"(?<=;linear_relres=)[^;\r\n]*","<gmres_residual>");
      byte[] normalized=encoding.GetBytes(line);kept.Write(normalized,0,normalized.Length);
     } else kept.Write(data,start,length);
     nonGmres++;
    }
    start=end+1;
   }
   nonGmres--; // Header is included in the digest, excluded from the row count.
   using(var sha=System.Security.Cryptography.SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(kept.ToArray())).Replace("-","");
  }
 }
}
'@
$xl=New-Object -ComObject Excel.Application
[uint32]$ownedNativePid=0
[void][NativeCompareIdentity]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd,[ref]$ownedNativePid)
$xpid=[int]$ownedNativePid
$excelNativeProcess=Get-Process -Id $xpid -ErrorAction Stop
$excelNativeVersion=$excelNativeProcess.MainModule.FileVersionInfo.FileVersion
Set-Content "$copy.running" "pid=$xpid"
$res=[Collections.Generic.List[object]]::new()
try{
 $xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false)
 if($null -eq $wb){throw "open failed"}
 $nativeModule=$wb.VBProject.VBComponents.Item('FEMNative').CodeModule
 if($nativeModule.Lines(1,$nativeModule.CountOfLines) -notmatch 'Public FEMNativeAssemblyCalls As Long'){
  if($AssemblyFast -ne 0){throw 'Assembly acceleration requires the new wrapper.'}
  $harness=$harness.Replace('FEMNativeAssemblyCalls','0')
 }
 if($nativeModule.Lines(1,$nativeModule.CountOfLines) -notmatch 'Public FEMNativeGmresTunedCalls As Long'){
  if($GmresTuned -ne 0){throw 'Tuned GMRES requires the new wrapper.'}
  $harness=$harness.Replace('FEMNativeGmresTunedCalls','0')
 }
 $h=$wb.VBProject.VBComponents.Add(1);$h.Name='TestHarness2';$h.CodeModule.AddFromString($harness)
 SetS $wb 'EXPORT_MODE' 'OFF'; SetS $wb 'EXPORT_LOAD' ''
 SetS $wb 'NATIVE_PRECISION' $Precision
 SetS $wb 'NATIVE_STATE_COPY' $StateCopy
 SetS $wb 'NATIVE_GMRES_SIMD' $GmresSimd
 SetS $wb 'IO_FAST_PATH' $IoFast
 SetS $wb 'EVENT_LOG_FAST' $EventFast
 SetS $wb 'NATIVE_ASSEMBLY' $AssemblyFast
 SetS $wb 'NATIVE_GMRES_TUNED' $GmresTuned
 if($Flow -ne ''){SetS $wb 'FLOW_POLICY' $Flow}
 $runs=@()
 if($Suite -like 'smoke*'){
  SetRows $wb '要素データ' $mesh.elements 10
  $wb.Worksheets.Item('載荷').Range('A2:F100').ClearContents()
  SetS $wb 'SOLVER_MAX_ITERATIONS' 2000
  foreach($c in @(@('Case1',10.0,5.0,0.5),@('Case2',10.0,25.0,1.1),@('Case3',20.0,5.0,0.9),@('Case4',20.0,35.0,2.0))){
   If($CaseName -ne '' -and $c[0] -ne $CaseName){continue}
   foreach($v in $Variants){ $runs+=,@($c,$v) } }
 if($runs.Count -eq 0){throw 'No matching smoke cases.'}
 } else { $runs+=,@($null,'default') }
 foreach($r in $runs){
  $c=$r[0];$v=$r[1]
  if($c){
   $sc=$(if($c[0] -eq 'Case1'){0.25}else{1.0})
   $nodes=[Collections.Generic.List[object]]::new();foreach($n in $mesh.nodes){$a=@($n);$a[1]=[double]$a[1]*$sc;$a[2]=[double]$a[2]*$sc;$nodes.Add($a)}
   SetRows $wb '節点データ' $nodes 9
   $wm=$wb.Worksheets.Item('材料データ');$wm.Cells(2,6).Value2=$c[2];$wm.Cells(2,7).Value2=$c[1]
   SetS $wb 'SRM_FIXED_FS' $(if($Suite -eq 'smokesrm'){0}else{$c[3]})
   foreach($k in 'ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V2B_COST','ACCEL_V3_COST_SEARCH','ACCEL_V4_ANDERSON','ACCEL_V5_GMRES_LU'){SetS $wb $k 0}
   SetS $wb 'ACCEL_ADAPTIVE' 0
   if($v -eq 'adaptive'){SetS $wb 'ACCEL_ADAPTIVE' 1;foreach($k in 'ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V4_ANDERSON'){SetS $wb $k 1}}
   $name=$c[0]
  } else { SetS $wb 'SRM_FIXED_FS' $FixedFs; $name='builtin' }
  foreach($mode in $Modes.Split(',')){
   $liveNativePids=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
   if($liveNativePids.Count -ne 1 -or $liveNativePids[0] -ne $xpid){throw 'Another Excel appeared; no test dispatched.'}
   $xl.Run("'"+$wb.Name+"'!FEMInvalidateSettingCache")
   Set-Content "$copy.running" "pid=$xpid case=$name/$v/$mode start=$(Get-Date -Format s)"
   $wallClock=[Diagnostics.Stopwatch]::StartNew()
   $t=$xl.Run("'"+$wb.Name+"'!TRun",$mode)
   $wallClock.Stop()
   $d=@($xl.Run("'"+$wb.Name+"'!TDispV"))
   $fos=$xl.Run("'"+$wb.Name+"'!TFos")
   $assembly=@($xl.Run("'"+$wb.Name+"'!TAssemblyV"))
   $bracket=@($xl.Run("'"+$wb.Name+"'!TBracketV"))
   $tunedCalls=[int]$xl.Run("'"+$wb.Name+"'!TTunedCalls")
   $nativeProfileData=$null
   if($NativeProfile){$nativeProfileData=@($xl.Run("'"+$wb.Name+"'!TNativeProfile"))}
   $liveNativePids=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
   $exclusiveExcel=($liveNativePids.Count -eq 1 -and $liveNativePids[0] -eq $xpid)
   $profilePath=$null
   if($Profile){
    $profilePath=$Out+'.'+$name+'.'+$v+'.'+$mode+'.profile.csv'
    [string]$profileText=$xl.Run("'"+$wb.Name+"'!FEMProfileCsv")
    $profileText | Set-Content -LiteralPath $profilePath -Encoding UTF8
   }
   $dllHash=$null; $nativePath=Join-Path $root 'soilfem_native_x64.dll'
   if(Test-Path -LiteralPath $nativePath){$dllHash=(Get-FileHash -LiteralPath $nativePath).Hash}
   $runOutput=Join-Path $root ([IO.Path]::GetFileNameWithoutExtension($copy)+'_out')
   $eventFile=Join-Path $runOutput 'solver_events.csv';$eventHash=$null;$eventBytes=0
   $nonGmresHash=$null;$decisionHash=$null;[int]$nonGmresRows=0;[int]$gmresRows=0
   if(Test-Path -LiteralPath $eventFile){$eventHash=(Get-FileHash -LiteralPath $eventFile).Hash;$eventBytes=(Get-Item -LiteralPath $eventFile).Length}
   if(Test-Path -LiteralPath $eventFile){
    $nonGmresHash=[NativeCompareIdentity]::NonGmresDigest($eventFile,[ref]$nonGmresRows,[ref]$gmresRows,$false)
    $decisionHash=[NativeCompareIdentity]::NonGmresDigest($eventFile,[ref]$nonGmresRows,[ref]$gmresRows,$true)
   }
   $res.Add([PSCustomObject]@{Case=$name;Variant=$v;Mode=$mode;Precision=$Precision;StateCopy=$StateCopy;GmresSimd=$GmresSimd;IoFast=$IoFast;EventFast=$EventFast;AssemblyFast=$AssemblyFast;Assembly=$assembly;Flow=$Flow;GmresTuned=$GmresTuned;TunedCalls=$tunedCalls;NativeProfile=$nativeProfileData;ProfileCsv=$profilePath;NativeDllSha256=$dllHash;ExcelPid=$xpid;ExcelVersion=$excelNativeVersion;ExclusiveExcel=$exclusiveExcel;WallSeconds=$wallClock.Elapsed.TotalSeconds;SolverEventsSha256=$eventHash;SolverEventsBytes=$eventBytes;DecisionEventsSha256=$decisionHash;NonGmresSolverEventsSha256=$nonGmresHash;NonGmresSolverEventsRows=$nonGmresRows;GmresEventRows=$gmresRows;RunOutputDirectory=$runOutput;Run=$t;Fos=$fos;Bracket=$bracket;Disp=$d})
   $res|ConvertTo-Json -Depth 4|Set-Content $Out -Encoding UTF8
   Write-Output "$name $v $mode $t FOS=$fos"
   if(-not $exclusiveExcel){throw 'Another Excel appeared during this test; preserve result but do not accept its timing.'}
  }
 }
 $wb.Close($false)
} finally {
 if($null -ne $nativeModule){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($nativeModule);$nativeModule=$null}
 if($null -ne $h){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($h);$h=$null}
 if($null -ne $wb){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb);$wb=$null}
 $xl.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);$xl=$null
 [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
 for($closeWait=0;$closeWait -lt 15;$closeWait++){
  if(-not (Get-Process -Id $xpid -ErrorAction SilentlyContinue)){break}
  Start-Sleep -Seconds 1
 }
 # Quit only this COM instance; do not terminate a PID that may have been reused.
 Remove-Item "$copy.running" -ErrorAction SilentlyContinue }
