param([string]$Book,[string]$Suite='smoke',[string]$Out,[double]$FixedFs=1.0,[string]$Modes='OFF,AUTO',[string]$Flow='')
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $Book
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
  TRun = CStr(Timer - t) & "|" & ResultStatus & "|" & CStr(AnalysisOK) & "|" & CStr(RelativeResidualFree) & "|" & CStr(P3SuccessfulIncrementCount) & "|" & CStr(P6FactorizationCount) & "|" & CStr(P6FactorizationReuseCount) & "|" & CStr(P3ActivePlasticPointCount) & "|" & FEMNativeStatusText() & "|" & CStr(P6ProfFactorMs) & "|" & CStr(P6ProfSolveMs) & "|" & CStr(P6ProfEvalMs) & "|" & CStr(P6ProfTangentMs) & "|" & AnalysisMessage
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
'@
function SetS($wb,$k,$v){$m=$(if($v -is [string]){'FEMWriteTextSetting'}else{'FEMWriteNumericSetting'});$xl.Run("'"+$wb.Name+"'!"+$m,$k,$v)}
function SetRows($wb,$sheet,$rows,$cols){$ws=$wb.Worksheets.Item($sheet);$ws.Range($ws.Cells(2,1),$ws.Cells($ws.UsedRange.Rows.Count,$cols)).ClearContents();$d=[object[,]]::new($rows.Count,$cols);for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $cols;$c++){$d[$r,$c]=$rows[$r][$c]}};$ws.Range($ws.Cells(2,1),$ws.Cells($rows.Count+1,$cols)).Value2=$d}
$xl=New-Object -ComObject Excel.Application
$xpid=@(Get-Process EXCEL|%{$_.Id}|?{$before -notcontains $_})[0]
Set-Content "$copy.running" "pid=$xpid"
$res=[Collections.Generic.List[object]]::new()
try{
 $xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false)
 if($null -eq $wb){throw "open failed"}; $h=$wb.VBProject.VBComponents.Add(1);$h.Name='TestHarness2';$h.CodeModule.AddFromString($harness)
 SetS $wb 'EXPORT_MODE' 'OFF'; SetS $wb 'EXPORT_LOAD' ''
 if($Flow -ne ''){SetS $wb 'FLOW_POLICY' $Flow}
 $runs=@()
 if($Suite -like 'smoke*'){
  $mesh=Get-Content (Join-Path $root 'smoke_mesh.json') -Raw|ConvertFrom-Json
  SetRows $wb '要素データ' $mesh.elements 10
  $wb.Worksheets.Item('載荷').Range('A2:F100').ClearContents()
  SetS $wb 'SOLVER_MAX_ITERATIONS' 2000
  foreach($c in @(@('Case1',10.0,5.0,0.5),@('Case2',10.0,25.0,1.1),@('Case3',20.0,5.0,0.9),@('Case4',20.0,35.0,2.0))){
   foreach($v in 'off','adaptive'){ $runs+=,@($c,$v) } }
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
  } else { if($FixedFs -gt 0){SetS $wb 'SRM_FIXED_FS' $FixedFs}; $name='builtin' }
  foreach($mode in $Modes.Split(',')){
   $xl.Run("'"+$wb.Name+"'!FEMInvalidateSettingCache")
   Set-Content "$copy.running" "pid=$xpid case=$name/$v/$mode start=$(Get-Date -Format s)"
   $t=$xl.Run("'"+$wb.Name+"'!TRun",$mode)
   $d=@($xl.Run("'"+$wb.Name+"'!TDispV"))
   $fos=$xl.Run("'"+$wb.Name+"'!TFos")
   $res.Add([PSCustomObject]@{Case=$name;Variant=$v;Mode=$mode;Run=$t;Fos=$fos;Disp=$d})
   $res|ConvertTo-Json -Depth 4|Set-Content $Out -Encoding UTF8
   Write-Output "$name $v $mode $t FOS=$fos"
  }
 }
 $wb.Close($false)
} finally { $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null; Start-Sleep 3
 if(Get-Process -Id $xpid -ErrorAction SilentlyContinue){Stop-Process -Id $xpid -Force}
 Remove-Item "$copy.running" -ErrorAction SilentlyContinue }
