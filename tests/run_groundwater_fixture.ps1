param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Fixture,[Parameter(Mandatory=$true)][string]$Out,[switch]$Resume,[string]$Flow="INCONSISTENT",[string]$Native="AUTO",[switch]$FrontUi,[string]$ResumeState='',[switch]$ReopenResume)
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
 $xl.DisplayAlerts=$false;$xl.EnableEvents=[bool]$data.shape;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false);$xl.EnableEvents=$false
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
Public Function WFConvergenceLimit() As Double
 WFConvergenceLimit=P3_ENGINEERING_RESIDUAL
End Function
Public Function WFNative() As String
 WFNative = FEMNativeStatusText() & "|" & FEMMatStatusText()
End Function
Public Function WFBuild() As String
 WFBuild = FEM_BUILD_STAMP
End Function
Public Function WFStress() As Variant
 Dim v() As Double, e As Long, g As Long, j As Long, n As Long
 On Error GoTo Unavailable
 If NumberOfElement < 1 Then WFStress = Array(): Exit Function
 ReDim v(NumberOfElement * 4 * 17 - 1)
 For e = 0 To NumberOfElement - 1
  For g = 0 To 3
   For j = 0 To 16
    v(n) = Elem(e).Stmat(j, g): n = n + 1
   Next j
  Next g
 Next e
 WFStress = v
Unavailable:
End Function
Public Function WFDisplacements() As Variant
 WFDisplacements = TDisp
End Function
'@)
 & {
  $mod.CodeModule.AddFromString(@'
Public Function WFCoupledState() As Variant
 WFCoupledState = Array(FSActive,FSTime,FSAccepted,FSRunningStage,FSStageElapsed,FSResumePending)
End Function
Public Function WFPressure() As Variant
 Dim result() As Double,e As Long,g As Long
 On Error GoTo Unavailable
 If NumberOfElement < 1 Then WFPressure = Array(): Exit Function
 ReDim result(0 To NumberOfElement*4-1)
 For e=0 To NumberOfElement-1
  For g=0 To 3: result(e*4+g)=GWNewPressure(e,g): Next g
 Next e
 WFPressure=result
Unavailable:
End Function
Public Function WFOriginalNodes() As Variant
 Dim result() As Double,n As Long,k As Long
 On Error GoTo Unavailable
 If NumberOfFreeNode < 1 Then Exit Function
 ReDim result(0 To NumberOfFreeNode-1,0 To 2)
 For n=0 To NumberOfFreeNode-1
  k=P6GetInternalFreeNode(n)
  result(n,0)=n+1: result(n,1)=TDisp(2*k): result(n,2)=TDisp(2*k+1)
 Next n
 WFOriginalNodes=result
Unavailable:
End Function
Public Function WFReferenceCoordinates() As Variant
 Dim result() As Double,n As Long,k As Long
 On Error GoTo Unavailable
 If NumberOfFreeNode < 1 Then Exit Function
 ReDim result(0 To NumberOfFreeNode-1,0 To 2)
 For n=0 To NumberOfFreeNode-1
  k=P6GetInternalFreeNode(n)
  result(n,0)=n+1: result(n,1)=XXX(2*k): result(n,2)=XXX(2*k+1)
 Next n
 WFReferenceCoordinates=result
Unavailable:
End Function
'@)
 }
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
 Rows $wb.Worksheets.Item('要素データ') 2 1 $data.elements $(if($data.elements[0].Count -ge 11){11}else{10})
 Rows $wb.Worksheets.Item('材料データ') 2 1 $data.materials 13
 Rows $wb.Worksheets.Item('ステージ') 2 1 $data.stages 6
 if($data.loads){Rows $wb.Worksheets.Item('載荷') 2 1 $data.loads $(if($data.loads[0].Count -ge 11){11}else{10})}
 if($data.jointPairs){Rows $wb.Worksheets.Item('接合') 2 1 $data.jointPairs 7}
 if($data.shape){
  $shapePage=$wb.Worksheets.Item('形状入力')
  $shapePage.Range('B8:D2000').ClearContents();$shapePage.Range('G8:J2000').ClearContents()
  Rows $shapePage 8 2 $data.shape.points 3
  Rows $shapePage 8 7 $data.shape.boundaries 4
 }
 $null=Call 'FEMWaterInstallUi'
 $water=$wb.Worksheets.Item('水条件')
 $water.Range('B3').Value2=[string]$data.water.mode;$water.Range('B4').Value2=[double]$data.water.gamma
 $water.Range('B5').Value2=[string]$data.water.defaultProfile;$water.Range('B6').Value2=[string]$data.water.auto
 Rows $water 12 1 $data.water.profiles 4
 Rows $water 12 6 $data.water.materials 5
 Rows $water 216 1 $data.water.boundaries 6
 if($data.water.stagePlans){Rows $water 12 12 $data.water.stagePlans 9}
 if($data.water.coupledMaterials){Rows $water 12 22 $data.water.coupledMaterials $(if($data.water.coupledMaterials[0].Count -ge 7){7}else{6})}
 if($data.water.timeOptions){Rows $water 12 29 $data.water.timeOptions 10}
 if($data.water.drainSets){Rows $water 216 22 $data.water.drainSets 7}
 if($data.water.timeCurves){Rows $water 420 22 $data.water.timeCurves 3}
 SetSetting 'EXPORT_MODE' $(if($Resume){'STAGE'}else{'OFF'});SetSetting 'EXPORT_LOAD' '';SetSetting 'FLOW_POLICY' $Flow;SetSetting 'NATIVE_DLL' $Native;SetSetting 'OUTPUT_STAGE_MODE' 'ALL'
 if($data.settings){foreach($property in $data.settings.PSObject.Properties){SetSetting $property.Name $property.Value}}
 if($Resume){SetSetting 'EXPORT_MODE' 'STAGE'}
 if($FrontUi){
  # Perturb canonical values, then restore through real front-sheet events.
  $null=Call 'FEM24Install'
  $mat=$wb.Worksheets.Item('材料データ');$young=$mat.Range('B2').Value2
  $mat.Range('B2').Value2=[double]$young+1.0
  $null=Call 'FEM24Materials';$xl.EnableEvents=$true
  $wb.Worksheets.Item('材料・接合').Range('B8').Value2=$young
  if($mat.Range('B2').Value2 -ne $young){throw 'Front material write-through failed.'}
  $xl.EnableEvents=$false
  $level=$water.Range('C12').Value2;$water.Range('C12').Value2=[double]$level+0.125
  $null=Call 'FEM24Water';$xl.EnableEvents=$true
  $wb.Worksheets.Item('境界・水条件').Range('C4').Value2='水位線'
  $wb.Worksheets.Item('境界・水条件').Range('C8').Value2=$level
  if($water.Range('C12').Value2 -ne $level){throw 'Front water write-through failed.'}
  $xl.EnableEvents=$false
  @{status='PASS';material_young=$young;water_level=$level;flow=$Flow;events='real Workbook_SheetChange'}|ConvertTo-Json|Set-Content (Join-Path $Out 'front_inputs.json') -Encoding UTF8
 }
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
 if($data.expectedAnalysisError){
  if($status -like 'PASS|*' -or -not $status.Contains([string]$data.expectedAnalysisError)){throw "Expected rejection was not observed: $status"}
  $v=$xl.Run("'"+$wb.Name+"'!WFOriginalNodes");$nodes=[Collections.Generic.List[object]]::new()
  if($null -ne $v){for($n=0;$n -lt $v.GetLength(0);$n++){$nodes.Add(@($v[$n,0],$v[$n,1],$v[$n,2]))}}
  $failure=@{case=$data.name;status=$status;expected=[string]$data.expectedAnalysisError;nodes=@($nodes);book_sha256=(Get-FileHash $Book).Hash;fixture_sha256=(Get-FileHash $Fixture).Hash}
  if($data.name -eq 'reset_t6_nonaffine' -or $data.checkResetPreservation){
   $v=$xl.Run("'"+$wb.Name+"'!WFReferenceCoordinates");$coords=[Collections.Generic.List[object]]::new()
   for($n=0;$n -lt $v.GetLength(0);$n++){$coords.Add(@($v[$n,0],$v[$n,1],$v[$n,2]))}
   $failure.reference_coordinates=@($coords);$failure.stress=@(Call 'WFStress')
  }
  if($data.water.coupledMaterials){$failure.coupled=@(Call 'WFCoupledState');$failure.pressure=@(Call 'WFPressure');$failure.stress=@(Call 'WFStress');$failure.profile=Call 'GWViewDefaultProfile'}
  $failure|ConvertTo-Json -Depth 7|Set-Content (Join-Path $Out 'expected_failure.json') -Encoding UTF8
  $wb.Close($false)
  return
 }
 if($status -notlike 'PASS|*'){
  @{case=$data.name;status=$status;seconds=$timer.Elapsed.TotalSeconds}|ConvertTo-Json|Set-Content (Join-Path $Out 'failure.json') -Encoding UTF8
  throw "Analysis: $status"
 }
 $record=@{case=$data.name;status=$status;seconds=$timer.Elapsed.TotalSeconds;metrics=@(Call 'WFMetrics');disp=@(Call 'WFDisplacements');stress=@(Call 'WFStress');book_sha256=(Get-FileHash $Book).Hash;fixture_sha256=(Get-FileHash $Fixture).Hash;excel_version=$xl.Version;excel_build=$xl.Build;build=(Call 'WFBuild');native=(Call 'WFNative');requested_native=$Native;flow=$Flow}
 $record|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Out 'result.json') -Encoding UTF8
 Write-Output ($record|ConvertTo-Json -Depth 2 -Compress)
 if($status -notlike 'PASS|*'){throw "Analysis: $status"}
 $outputSheets=@('水圧結果','結果節点','結果要素','結果増分','結果塑性','ステージ結果','水圧確認','接合結果')
 if($data.water.coupledMaterials){$outputSheets+='時間履歴'}
 foreach($name in $outputSheets){
  try{$ws=$wb.Worksheets.Item($name)}catch{if($data.water.mode -eq 'NONE' -and $name -in '水圧結果','水圧確認'){continue};throw}
  $array=$ws.UsedRange.Value2;$values=[Collections.Generic.List[object]]::new()
  for($r=1;$r -le $array.GetLength(0);$r++){$row=[Collections.Generic.List[object]]::new();for($c=1;$c -le $array.GetLength(1);$c++){$row.Add($array[$r,$c])};$values.Add(@($row))}
  ConvertTo-Json -InputObject @($values) -Depth 8 | Set-Content (Join-Path $Out ($name+'.json')) -Encoding UTF8
 }
 $wb.Save()
 if($Resume){
  $wb.SaveCopyAs((Join-Path $Out 'fresh_results.xlsm'))
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
  $resumeFolder='AUTO'
  if($ResumeState){
   $matches=@(Get-ChildItem -LiteralPath $Out -Recurse -Directory|Where-Object {$_.Name -eq $ResumeState -and $_.FullName -notlike '*initial_run_logs*'})
   if($matches.Count -ne 1){throw "ResumeState must match exactly one checkpoint: $ResumeState"}
   $resumeFolder=$matches[0].FullName
  }
  if($ReopenResume){$wb.Close($false);$wb=$xl.Workbooks.Open($copy,0,$false);$xl.EnableEvents=$false}
  SetSetting 'EXPORT_LOAD' $resumeFolder
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
  $resumeResidualLimit=1e-7
  if($data.checkResumeAgainstSolverTolerance){
   # Near an SRM numerical boundary, verify the solver's unchanged acceptance
   # contract and equality of the accepted residual, rather than a toy-case gate.
   $resumeResidualLimit=[double](Call 'WFConvergenceLimit')
   if([double]::IsNaN($resumeResidualLimit) -or [double]::IsInfinity($resumeResidualLimit) -or $resumeResidualLimit -le 0){throw 'Invalid solver convergence limit'}
   if([double]$oldMetrics[3] -gt $resumeResidualLimit -or [Math]::Abs($oldMetrics[3]-$resumed.metrics[3]) -gt 1e-9){throw 'Accepted residual changed on resume'}
   @{status='PASS';source='P3_ENGINEERING_RESIDUAL';limit=$resumeResidualLimit;fresh=$oldMetrics[3];resumed=$resumed.metrics[3];delta=[Math]::Abs($oldMetrics[3]-$resumed.metrics[3]);production_tolerance_changed=$false}|ConvertTo-Json|Set-Content (Join-Path $Out 'resume_residual_contract.json') -Encoding UTF8
  }
  if($maxDu -gt 1e-9 -or [double]$resumed.metrics[3] -gt $resumeResidualLimit -or [Math]::Abs($oldMetrics[2]-$resumed.metrics[2]) -gt 1e-6){throw "Resume mismatch du=$maxDu metrics=$($resumed.metrics)"}
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
 $xl.EnableEvents=$false
 if($wb){try{$wb.Close($false)}catch{}}
 foreach($reference in @($testModule,$mod,$water,$ws,$wb)){if($null -ne $reference -and [Runtime.InteropServices.Marshal]::IsComObject($reference)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference)}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
