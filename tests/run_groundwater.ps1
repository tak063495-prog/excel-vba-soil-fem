param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out,[switch]$DryOnly)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'An Excel session is running; do not touch user work.'}
if(Test-Path -LiteralPath $Out){throw 'Preserve prior evidence; choose a new output directory.'}
$Book=(Resolve-Path -LiteralPath $Book).Path
$Out=[IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory -Path $Out | Out-Null
$copy=Join-Path $Out 'water_tests.xlsm'
Copy-Item -LiteralPath $Book -Destination $copy
Copy-Item -LiteralPath (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') -Destination $Out
$xl=New-Object -ComObject Excel.Application
$result=[Collections.Generic.List[object]]::new()
try {
 $xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false)
 $component=$wb.VBProject.VBComponents.Add(1);$component.Name='WaterTestDriver'
 $source=Get-Content (Join-Path $PSScriptRoot 'native/WaterHarness.bas') -Raw -Encoding UTF8
 $source += (Get-Content (Join-Path $PSScriptRoot 'native/WaterGeometryHarness.bas') -Raw -Encoding UTF8).Replace('Option Explicit','')
 $source += @'

Public Function GWRunMechanicalTest() As String
  On Error GoTo Failed
  SuppressUserMessages = True
  P0_RunAnalysis
  GWRunMechanicalTest = ResultStatus & "|" & AnalysisMessage & "|" & CStr(RelativeResidualFree)
  Exit Function
Failed:
  GWRunMechanicalTest = "ERROR|" & Err.Description
End Function
Public Function GWTestDisp() As Variant
  GWTestDisp = TDisp
End Function
Public Function GWTestBracket() As Variant
  GWTestBracket = Array(P3FosPass,P3FosFail,P3FosWidth)
End Function
Public Function GWTestPreview() As String
  On Error GoTo Failed
  FEMInvalidateSettingCache
  GWPreviewCompute
  ThisWorkbook.Worksheets("水条件").Range("B7").Value2 = GWInputSignature()
  GWTestPreview = "PASS"
  Exit Function
Failed:
  GWTestPreview = "ERROR|" & Err.Description
End Function
'@
 $component.CodeModule.AddFromString($source)
 function Macro([string]$name){$xl.Run("'"+$wb.Name+"'!"+$name)}
 function Setting([string]$key,$value){if($value -is [string]){$null=$xl.Run("'"+$wb.Name+"'!FEMWriteTextSetting",$key,$value)}else{$null=$xl.Run("'"+$wb.Name+"'!FEMWriteNumericSetting",$key,[double]$value)}}
 if(-not $DryOnly){
  $null=Macro 'FEMWaterInstallUi'
  foreach($name in 'GWTestDarcy','GWTestDarcyCases','GWTestWetGeometry','GWTestDarcyPatches'){
   $value=Macro $name
   $result.Add(@{case=$name;result=$value});$result | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $Out 'results.json') -Encoding UTF8
   if($value -notlike 'PASS*'){throw "$name : $value"}
  }
 }
 # Straight-sided regular Q8 mesh, width=2, height=2, 2 by 2 elements.
 $coords=[Collections.Generic.List[object]]::new();$lookup=@{};$elements=[Collections.Generic.List[object]]::new()
 foreach($iy in 0..1){foreach($ix in 0..1){
  $nodes=[Collections.Generic.List[int]]::new()
  foreach($offset in @(@(0,0),@(2,0),@(2,2),@(0,2),@(1,0),@(2,1),@(1,2),@(0,1))){
   $xx=2*$ix+$offset[0];$yy=2*$iy+$offset[1];$key="$xx,$yy"
   if(-not $lookup.ContainsKey($key)){$lookup[$key]=$coords.Count+1;$coords.Add(@(([double]$xx/2),([double]$yy/2)))}
   $nodes.Add($lookup[$key])
  }
  $elements.Add(@($elements.Count+1)+@($nodes)+@(1))
 }}
 $wn=$wb.Worksheets.Item('節点データ');$we=$wb.Worksheets.Item('要素データ');$wm=$wb.Worksheets.Item('材料データ');$st=$wb.Worksheets.Item('ステージ')
 $wn.Range('A2:N10000').ClearContents();$we.Range('A2:J10000').ClearContents();$wm.Range('A2:M1000').ClearContents();$st.Range('A2:F1000').ClearContents();$wb.Worksheets.Item('載荷').Range('A2:J1000').ClearContents()
 for($i=0;$i -lt $coords.Count;$i++){
  $xy=$coords[$i];$row=@(($i+1),$xy[0],$xy[1],1,[int]($xy[1] -eq 0),0,0,0,0)
  for($j=0;$j -lt 9;$j++){$wn.Cells.Item($i+2,$j+1).Value2=[double]$row[$j]}
 }
 for($i=0;$i -lt $elements.Count;$i++){for($j=0;$j -lt 10;$j++){$we.Cells.Item($i+2,$j+1).Value2=[double]$elements[$i][$j]}}
 $material=@(1,10000,0.3,1,20,0,1000000,0,1,1,'SOIL',0,0)
 $matValues=[object[,]]::new(1,13)
 for($j=0;$j -lt 13;$j++){$matValues[0,$j]=$material[$j]}
 $wm.Range("A2:M2").Value2=$matValues
 $st.Cells.Item(2,1).Value2=1.;$st.Cells.Item(2,2).Value2='GRAVITY';$st.Cells.Item(2,5).Value2=1.
 Setting 'EXPORT_MODE' 'OFF';Setting 'EXPORT_LOAD' '';Setting 'OUTPUT_STAGE_MODE' 'FINAL';Setting 'FLOW_POLICY' 'INCONSISTENT';Setting 'NATIVE_DLL' 'AUTO'
 $cases=if($DryOnly){@('dry')}else{@('dry','hydro','submerged','steady','total')}
 foreach($case in $cases){
  if(-not $DryOnly){
   $water=$wb.Worksheets.Item('水条件');$water.Range('B7').ClearContents();$water.Range('A12:D211').ClearContents();$water.Range('F12:J211').ClearContents();$water.Range('A216:F415').ClearContents()
   $water.Range('B3').Value2=if($case -eq 'dry'){'NONE'}elseif($case -eq 'steady'){'STEADY'}else{'HYDROSTATIC'}
   $water.Range('B4').Value2=9.80665;$water.Range('B5').Value2='外水';$water.Range('B6').Value2='AUTO'
   $head=if($case -eq 'submerged' -or $case -eq 'steady'){3.}else{2.}
   $water.Range('A12').Value2='外水';$water.Range('B12').Value2=0.;$water.Range('C12').Value2=$head
   $water.Range('A13').Value2='外水';$water.Range('B13').Value2=2.;$water.Range('C13').Value2=$head
   $water.Range('F12').Value2=1.;$water.Range('G12').Value2=if($case -eq 'total'){'TOTAL'}else{'EFFECTIVE'};$water.Range('H12').Value2=0.00001
   if($case -eq 'steady'){
    $row=216
    foreach($boundary in @(@(1,4),@(3,4),@(2,2),@(4,2))){$water.Cells.Item($row,1).Value2=[double]$boundary[0];$water.Cells.Item($row,2).Value2=[double]$boundary[1];$water.Cells.Item($row,3).Value2='HEAD';$water.Cells.Item($row,4).Value2='外水';$row++}
   }
   if($case -ne 'dry'){$preview=Macro 'GWTestPreview';if($preview -ne 'PASS'){throw "$case preview: $preview"}}
  }
  $timer=[Diagnostics.Stopwatch]::StartNew();$status=Macro 'GWRunMechanicalTest';$timer.Stop()
  $disp=@(Macro 'GWTestDisp');$entry=@{case=$case;status=$status;seconds=$timer.Elapsed.TotalSeconds;disp=$disp;bracket=@(Macro 'GWTestBracket')}
  $result.Add($entry);$result | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $Out 'results.json') -Encoding UTF8
  Write-Output "$case $status seconds=$($timer.Elapsed.TotalSeconds)"
  if($status -notlike 'PASS|*'){throw "$case : $status"}
  if(-not $DryOnly -and $case -ne 'dry'){$wb.Worksheets.Item('水圧結果').UsedRange.Value2 | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $Out ($case+'_water.json')) -Encoding UTF8}
 }
 if(-not $DryOnly){
  $invalid=Macro 'GWTestInvalidInputs';$result.Add(@{case='invalid_inputs';result=$invalid})
  $result | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $Out 'results.json') -Encoding UTF8
  if($invalid -notlike 'PASS*'){throw $invalid}
 }
 $wb.Save();$wb.Close($false)
}catch{
 $_.ScriptStackTrace | Write-Output
 throw
}finally{
 if($null -ne $wb){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
