param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Fixture,[Parameter(Mandatory=$true)][string]$Out,[switch]$Resume,[string]$Flow="INCONSISTENT",[string]$Native="AUTO")
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; do not dispatch.'}
if(Test-Path -LiteralPath $Out){throw 'Preserve previous evidence.'}
$Book=(Resolve-Path -LiteralPath $Book).Path;$Fixture=(Resolve-Path -LiteralPath $Fixture).Path
$Out=[IO.Path]::GetFullPath($Out);New-Item -ItemType Directory -Path $Out | Out-Null
$data=Get-Content -LiteralPath $Fixture -Raw -Encoding UTF8 | ConvertFrom-Json
$copy=Join-Path $Out ($data.name+'.xlsm');Copy-Item -LiteralPath $Book -Destination $copy
Copy-Item -LiteralPath (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') -Destination $Out
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false)
 $mod=$wb.VBProject.VBComponents.Add(1);$mod.Name='WaterFixtureDriver'
 $mod.CodeModule.AddFromString(@'
Public Function WFPreview() As String
 On Error GoTo Failed
 SuppressUserMessages = True
 FEMInvalidateSettingCache
 If UCase$(Trim$(CStr(ThisWorkbook.Worksheets("水条件").Range("B3").Value2))) <> "NONE" Then GWPreviewCompute
 ThisWorkbook.Worksheets("水条件").Range("B7").Value2 = GWInputSignature()
 WFPreview = "PASS"
 Exit Function
Failed:
 WFPreview = "ERROR|" & Err.Description & "|" & Err.Source
End Function
Public Function WFRun() As String
 On Error GoTo Failed
 SuppressUserMessages = True
 P0_RunAnalysis
 WFRun = ResultStatus & "|" & AnalysisMessage
 Exit Function
Failed:
 WFRun = "ERROR|" & Err.Description
End Function
Public Function WFMetrics() As Variant
 WFMetrics = Array(TotalSelfWeight,P6ReactionSumX,P6ReactionSumY,RelativeResidualFree,FSS,P3FosPass,P3FosFail,P3FosWidth)
End Function
Public Function WFNative() As String
 WFNative = FEMNativeStatusText() & "|" & FEMMatStatusText()
End Function
Public Function WFBuild() As String
 WFBuild = FEM_BUILD_STAMP
End Function
Public Function WFStress() As Variant
 Dim v() As Double, e As Long, g As Long, j As Long, n As Long
 ReDim v(NumberOfElement * 4 * 17 - 1)
 For e = 0 To NumberOfElement - 1
  For g = 0 To 3
   For j = 0 To 16
    v(n) = Elem(e).Stmat(j, g): n = n + 1
   Next j
  Next g
 Next e
 WFStress = v
End Function
Public Function WFDisplacements() As Variant
 WFDisplacements = TDisp
End Function
'@)
 if($data.checkTangent){
  $testModule=$wb.VBProject.VBComponents.Add(1);$testModule.Name='JointTangentDriver'
  $testModule.CodeModule.AddFromString([IO.File]::ReadAllText("$PSScriptRoot/joint_tangent_driver.bas"))
 }
 function Call([string]$name){$xl.Run("'"+$wb.Name+"'!"+$name)}
 function SetSetting([string]$key,$value){if($value -is [string]){$null=$xl.Run("'"+$wb.Name+"'!FEMWriteTextSetting",$key,$value)}else{$null=$xl.Run("'"+$wb.Name+"'!FEMWriteNumericSetting",$key,[double]$value)}}
 function Rows($ws,[int]$firstRow,[int]$firstColumn,$rows,[int]$cols){
  if($null -eq $rows -or $rows.Count -eq 0){return}
  $v=[object[,]]::new($rows.Count,$cols)
  for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $cols;$c++){$v[$r,$c]=$rows[$r][$c]}}
  $ws.Range($ws.Cells($firstRow,$firstColumn),$ws.Cells(($firstRow+$rows.Count-1),($firstColumn+$cols-1))).Value2=$v
 }
 $null=Call 'FEMEnsureStageSheets'
 foreach($name in '節点データ','要素データ','材料データ','ステージ','載荷'){$wb.Worksheets.Item($name).Range('A2:N10000').ClearContents()}
 $wb.Worksheets.Item('接合').Range('A2:G10000').ClearContents()
 Rows $wb.Worksheets.Item('節点データ') 2 1 $data.nodes 9
 Rows $wb.Worksheets.Item('要素データ') 2 1 $data.elements 10
 Rows $wb.Worksheets.Item('材料データ') 2 1 $data.materials 13
 Rows $wb.Worksheets.Item('ステージ') 2 1 $data.stages 6
 if($data.loads){Rows $wb.Worksheets.Item('載荷') 2 1 $data.loads 10}
 if($data.jointPairs){Rows $wb.Worksheets.Item('接合') 2 1 $data.jointPairs 7}
 $null=Call 'FEMWaterInstallUi'
 $water=$wb.Worksheets.Item('水条件')
 $water.Range('B3').Value2=[string]$data.water.mode;$water.Range('B4').Value2=[double]$data.water.gamma
 $water.Range('B5').Value2=[string]$data.water.defaultProfile;$water.Range('B6').Value2=[string]$data.water.auto
 Rows $water 12 1 $data.water.profiles 4
 Rows $water 12 6 $data.water.materials 5
 Rows $water 216 1 $data.water.boundaries 6
 SetSetting 'EXPORT_MODE' $(if($Resume){'STAGE'}else{'OFF'});SetSetting 'EXPORT_LOAD' '';SetSetting 'FLOW_POLICY' $Flow;SetSetting 'NATIVE_DLL' $Native;SetSetting 'OUTPUT_STAGE_MODE' 'ALL'
 if($data.settings){foreach($property in $data.settings.PSObject.Properties){SetSetting $property.Name $property.Value}}
 if($Resume){SetSetting 'EXPORT_MODE' 'STAGE'}
 $preview=Call 'WFPreview'
 if($data.expectedPreviewError){
  $expected=[string]$data.expectedPreviewError
  if($preview -notlike 'ERROR|*' -or -not $preview.Contains($expected)){throw "Unexpected preview validation: $preview"}
  @{status='PASS';expected_error=$expected;actual_error=$preview;book_sha256=(Get-FileHash $Book).Hash;fixture_sha256=(Get-FileHash $Fixture).Hash}|ConvertTo-Json|Set-Content (Join-Path $Out 'input_check.json') -Encoding UTF8
  $wb.Close($false)
  return
 }
 if($preview -ne 'PASS'){throw "Preview: $preview"}
 $timer=[Diagnostics.Stopwatch]::StartNew();$status=Call 'WFRun';$timer.Stop()
 # Failed input/setup may leave no elements allocated. Do not mask its message
 # with a second error while collecting an unallocated stress/DOF array.
 if($status -notlike 'PASS|*'){
  @{case=$data.name;status=$status;seconds=$timer.Elapsed.TotalSeconds}|ConvertTo-Json|Set-Content (Join-Path $Out 'failure.json') -Encoding UTF8
  throw "Analysis: $status"
 }
 $record=@{case=$data.name;status=$status;seconds=$timer.Elapsed.TotalSeconds;metrics=@(Call 'WFMetrics');disp=@(Call 'WFDisplacements');stress=@(Call 'WFStress');book_sha256=(Get-FileHash $Book).Hash;fixture_sha256=(Get-FileHash $Fixture).Hash;excel_version=$xl.Version;excel_build=$xl.Build;build=(Call 'WFBuild');native=(Call 'WFNative');requested_native=$Native;flow=$Flow}
 $record|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Out 'result.json') -Encoding UTF8
 Write-Output ($record|ConvertTo-Json -Depth 2 -Compress)
 if($status -notlike 'PASS|*'){throw "Analysis: $status"}
 foreach($name in '水圧結果','結果節点','ステージ結果','水圧確認','接合結果'){
  try{$ws=$wb.Worksheets.Item($name)}catch{if($data.water.mode -eq 'NONE' -and $name -in '水圧結果','水圧確認'){continue};throw}
  $array=$ws.UsedRange.Value2;$values=[Collections.Generic.List[object]]::new()
  for($r=1;$r -le $array.GetLength(0);$r++){$row=[Collections.Generic.List[object]]::new();for($c=1;$c -le $array.GetLength(1);$c++){$row.Add($array[$r,$c])};$values.Add(@($row))}
  ConvertTo-Json -InputObject @($values) -Depth 8 | Set-Content (Join-Path $Out ($name+'.json')) -Encoding UTF8
 }
 $wb.Save()
 if($Resume){
  # Analysis startup rewrites run/performance logs. Retain the initial solve
  # separately before checkpoint-resume verification starts.
  $initialLogs=Join-Path $Out 'initial_run_logs'
  New-Item -ItemType Directory -Path $initialLogs | Out-Null
  Get-ChildItem $Out -Directory | Where-Object {$_.Name -like '*_out'} | ForEach-Object {
   $destination=Join-Path $initialLogs $_.Name
   New-Item -ItemType Directory -Path $destination | Out-Null
   Get-ChildItem $_.FullName -File | Where-Object {$_.Extension -in '.csv','.log','.txt','.json'} | Copy-Item -Destination $destination
  }
  $oldDisp=@($record.disp);$oldMetrics=@($record.metrics)
  SetSetting 'EXPORT_LOAD' 'AUTO'
  $timer.Restart();$status=Call 'WFRun';$timer.Stop()
  $resumed=@{case=$data.name;status=$status;seconds=$timer.Elapsed.TotalSeconds;metrics=@(Call 'WFMetrics');disp=@(Call 'WFDisplacements');stress=@(Call 'WFStress')}
  $resumed|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Out 'resume.json') -Encoding UTF8
  $ws=$wb.Worksheets.Item('接合結果');$array=$ws.UsedRange.Value2;$values=[Collections.Generic.List[object]]::new()
  for($r=1;$r -le $array.GetLength(0);$r++){$row=[Collections.Generic.List[object]]::new();for($c=1;$c -le $array.GetLength(1);$c++){$row.Add($array[$r,$c])};$values.Add(@($row))}
  ConvertTo-Json -InputObject @($values) -Depth 8 | Set-Content (Join-Path $Out 'resume_接合結果.json') -Encoding UTF8
  if($status -notlike 'PASS|*'){throw "Resume: $status"}
  for($i=4;$i -lt 8;$i++){if([Math]::Abs($oldMetrics[$i]-$resumed.metrics[$i]) -gt 1e-10){throw "Resume Fs mismatch index=$i"}}
  # Principal directions are undefined for equal eigenvalues. Large yield-function
  # sentinels use the existing VBA CSV formatter (about 15 significant digits).
  $maxStress=0.;for($i=0;$i -lt $record.stress.Count;$i++){
   $delta=[Math]::Abs($record.stress[$i]-$resumed.stress[$i])
   if($i % 17 -eq 9){
    $yieldTolerance=1e-6+5e-15*[Math]::Max([Math]::Abs($record.stress[$i]),[Math]::Abs($resumed.stress[$i]))
    if($delta -gt $yieldTolerance){throw "Resume yield diagnostic mismatch index=$i delta=$delta tolerance=$yieldTolerance"}
   }elseif($i % 17 -ne 6){$maxStress=[Math]::Max($maxStress,$delta)}
  }
  if($maxStress -gt 1e-6){throw "Resume stress mismatch max=$maxStress"}
  $maxDu=0.;for($i=0;$i -lt $oldDisp.Count;$i++){$maxDu=[Math]::Max($maxDu,[Math]::Abs($oldDisp[$i]-$resumed.disp[$i]))}
  if($maxDu -gt 1e-9 -or [double]$resumed.metrics[3] -gt 1e-7 -or [Math]::Abs($oldMetrics[2]-$resumed.metrics[2]) -gt 1e-6){throw "Resume mismatch du=$maxDu metrics=$($resumed.metrics)"}
 }
 if($data.checkTangent){
  $tangent=@(Call 'JTCheck')
  if($tangent.Count -ne 4 -or $tangent[0] -like 'ERROR|*'){throw "Tangent check: $tangent"}
  $tangent|ConvertTo-Json|Set-Content (Join-Path $Out 'tangent.json') -Encoding UTF8
 }
 # A completed run must retain node/element tables for normal result drawing.
 foreach($sheetName in '結果節点','結果要素'){
  $resultPage=$wb.Worksheets.Item($sheetName)
  if(-not [string]$resultPage.Range('A2').Value2){throw "Completed result table is empty: $sheetName"}
 }
 # Resume writes a diagnostic start snapshot; persist the completed results afterward.
 $wb.Save()
 $wb.Close($false)
}catch{
 $_.ScriptStackTrace | Write-Output
 throw
}finally{
 foreach($reference in @($testModule,$mod,$water,$ws,$wb)){if($null -ne $reference -and [Runtime.InteropServices.Marshal]::IsComObject($reference)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference)}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
