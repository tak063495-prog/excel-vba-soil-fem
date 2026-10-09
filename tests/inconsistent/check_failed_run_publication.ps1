param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'; New-Item -ItemType Directory -Force $OutputRoot|Out-Null
$sourceTask=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$hashTask=(Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash
$fixtureTask=(Get-Content (Join-Path $PSScriptRoot '../literature/fixtures.json') -Raw -Encoding UTF8|ConvertFrom-Json).fixtures|Where-Object name -eq 'confined_gravity'
function RowsTask($book,$sheet,$rows,$columns){
 $ws=$book.Worksheets.Item($sheet);$ws.Range($ws.Cells(2,1),$ws.Cells(2000,$columns)).ClearContents();if($rows.Count -eq 0){return}
 $data=[object[,]]::new($rows.Count,$columns);for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $columns;$c++){$v=$rows[$r][$c];if($null -ne $v -and $v -isnot [string]){$v=[double]$v};$data[$r,$c]=$v}}
 $ws.Range($ws.Cells(2,1),$ws.Cells($rows.Count+1,$columns)).Value2=$data
}
$cases=@(
 @{name='gravity_stagnation';kind='GRAVITY';fault=1;stages=@(@(1,'GRAVITY',$null,$null,1,'gravity'),@(2,'SRM',$null,$null,1,'srm'));srm=$true},
 @{name='gravity_capacity';kind='GRAVITY';fault=2;stages=@(@(1,'GRAVITY',$null,$null,1,'gravity'),@(2,'SRM',$null,$null,1,'srm'));srm=$true},
 @{name='gravity_material';kind='GRAVITY';fault=3;stages=@(@(1,'GRAVITY',$null,$null,1,'gravity'),@(2,'SRM',$null,$null,1,'srm'));srm=$true},
 @{name='gravity_cancel';kind='GRAVITY';fault=4;stages=@(@(1,'GRAVITY',$null,$null,1,'gravity'),@(2,'SRM',$null,$null,1,'srm'));srm=$true},
 @{name='load_failure';kind='APPLY';fault=1;stages=@(@(1,'GRAVITY',$null,$null,1,'gravity'),@(2,'SRM',$null,$null,1,'srm'),@(3,'LOAD',1,$null,1,'load'));srm=$true},
 @{name='ordinary_gravity';kind='GRAVITY';fault=1;stages=,@(1,'GRAVITY',$null,$null,1,'gravity');srm=$false},
 @{name='disabled_srm';kind='GRAVITY';fault=1;stages=@(@(1,'GRAVITY',$null,$null,1,'gravity'),@(2,'SRM',$null,$null,0,'srm'));srm=$false},
 @{name='successful_srm';kind='NONE';fault=0;stages=@(@(1,'GRAVITY',$null,$null,1,'gravity'),@(2,'SRM',$null,$null,1,'srm'));srm=$true}
)
$recordsTask=@()
foreach($caseTask in $cases){
 $copyTask=Join-Path $OutputRoot ('failed_publication_'+$caseTask.name+'.xlsm');Copy-Item -LiteralPath $sourceTask -Destination $copyTask -Force
 $xlTask=New-Object -ComObject Excel.Application;$xlTask.Visible=$false;$xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1;$wbTask=$null
 try{
  $wbTask=$xlTask.Workbooks.Open([IO.Path]::GetFullPath($copyTask),0,$false);$cmTask=$wbTask.VBProject.VBComponents.Item('FEMEngine').CodeModule;$sTask=$cmTask.Lines(1,$cmTask.CountOfLines)
  $sTask=$sTask.Replace('Option Explicit',"Option Explicit`r`nPublic IncoPubFault As Long`r`nPublic IncoPubKind As String")
  $needleTask='Private Function P3ExecuteStage(ByVal stageIndex As Long) As Boolean';if(-not $sTask.Contains($needleTask)){throw 'stage hook absent'}
  $hookTask=@'
__P3_STAGE_NEEDLE__
  If IncoPubFault > 0 And (IncoPubKind = "" Or UCase$(P3StageKind(stageIndex)) = UCase$(IncoPubKind)) Then
    If IncoPubFault = 1 Then SetAnalysisFailure RESULT_NONCONVERGED, "Injected global stagnation", vbObjectError + 3701, -1, -1, 1, 0: P3FailureKind = "GLOBAL_STAGNATION"
    If IncoPubFault = 2 Then SetAnalysisFailure RESULT_CAPACITY_ERROR, "Injected capacity failure", vbObjectError + 3702, -1, -1, 1, 0: P3FailureKind = "CAPACITY"
    If IncoPubFault = 3 Then SetAnalysisFailure RESULT_MATERIAL_ERROR, "Injected material failure", vbObjectError + 3703, -1, -1, 1, 0: P3FailureKind = "MATERIAL"
    If IncoPubFault = 4 Then SetAnalysisFailure RESULT_NONCONVERGED, "Injected cancellation", 18, -1, -1, 1, 0: P3FailureKind = "CANCELLED"
    P3ExecuteStage = False: Exit Function
  End If
'@
  $hookTask=$hookTask.Replace('__P3_STAGE_NEEDLE__',$needleTask)
  $sTask=$sTask.Replace($needleTask,($hookTask -replace "`r?`n","`r`n"));$cmTask.DeleteLines(1,$cmTask.CountOfLines);$cmTask.AddFromString($sTask)
  $cmTask.AddFromString(@'
Public Function IncoPublicationRun(ByVal faultId As Long, ByVal selectedKind As String) As Variant
  SetP0SilentMode True: IncoPubFault = faultId: IncoPubKind = selectedKind: P0_RunAnalysis
  IncoPublicationRun = Array(ResultStatus, AnalysisErrorNumber, AnalysisMessage, P3FosInterpretation, P3FailureKind, P3SrmTrialClassification, P3FosPass, P3FosFail, P3FosMid, P3FosWidth, P3FosBracket, FSS, P3FosDisplayFs(), AnalysisOK)
End Function
'@)
  RowsTask $wbTask '材料データ' $fixtureTask.materials 13;RowsTask $wbTask '節点データ' $fixtureTask.nodes 9;RowsTask $wbTask '要素データ' $fixtureTask.elements 10;RowsTask $wbTask '載荷' $fixtureTask.loads 10;RowsTask $wbTask '接合' @() 5;RowsTask $wbTask 'ステージ' $caseTask.stages 6
  $prefixTask="'"+$wbTask.Name+"'!";foreach($k in @('ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V2B_COST','ACCEL_V3_COST_SEARCH','ACCEL_V4_ANDERSON','ACCEL_V5_GMRES_LU','ACCEL_ADAPTIVE','ACCEL_STEP_RECOVERY')){$xlTask.Run($prefixTask+'FEMWriteNumericSetting',$k,0)}
  $xlTask.Run($prefixTask+'FEMWriteNumericSetting','SRM_FMAX',1.1);$xlTask.Run($prefixTask+'FEMWriteTextSetting','DEBUG_MODE','STAGE');$xlTask.Run($prefixTask+'FEMWriteTextSetting','EXPORT_MODE','OFF')
  $vTask=@($xlTask.Run($prefixTask+'IncoPublicationRun',$caseTask.fault,$caseTask.kind));$outDir=Join-Path $OutputRoot (($copyTask|Split-Path -Leaf) -replace '\.xlsm$','_out');$perf=Join-Path $outDir 'perf_summary.txt'
  if($caseTask.fault -gt 0 -and $vTask[0] -eq 'PASS'){throw "$($caseTask.name): injected failure was reported as PASS"}
  if($caseTask.fault -gt 0){
    $expectedStatusTask=@('','NONCONVERGED','CAPACITY_ERROR','MATERIAL_ERROR','NONCONVERGED')[$caseTask.fault]
    $expectedKindTask=@('','GLOBAL_STAGNATION','CAPACITY','MATERIAL','CANCELLED')[$caseTask.fault]
    $expectedErrorTask=if($caseTask.fault -eq 4){18}else{-2147221504+3700+$caseTask.fault}
    if($vTask[0] -cne $expectedStatusTask -or $vTask[1] -ne $expectedErrorTask -or $vTask[4] -cne $expectedKindTask -or $vTask[13]){throw "$($caseTask.name): stop reason lost: $($vTask -join '|')"}
  }
  $failed=($caseTask.fault -gt 0 -and $caseTask.srm);$passOk=($caseTask.name -eq 'load_failure' -and $vTask[6] -gt 0) -or ($caseTask.name -ne 'load_failure' -and $vTask[6] -eq 0)
  if($failed -and ($vTask[3] -ne 'UNDETERMINED' -or -not $passOk -or $vTask[7] -ne 0 -or $vTask[8] -ne 0 -or $vTask[9] -ne 0 -or $vTask[10] -or $vTask[11] -ne 0 -or $vTask[12] -ne 0)){throw "$($caseTask.name): stale FOS publication: $($vTask -join '|')"}
  if(-not $failed -and $vTask[3] -eq 'UNDETERMINED'){throw "$($caseTask.name): control was incorrectly invalidated: $($vTask -join '|')"}
  $eventsTask=Join-Path $outDir 'solver_events.csv'
  $eventTextTask=if(Test-Path $eventsTask){[IO.File]::ReadAllText($eventsTask,[Text.Encoding]::GetEncoding(932))}else{''}
  if($failed -and $eventTextTask -notmatch 'SRM_RUN_INVALID'){throw "$($caseTask.name): invalidation event missing"}
  if(-not $failed -and $eventTextTask -match 'SRM_RUN_INVALID'){throw "$($caseTask.name): control invalidation event"}
  if($caseTask.name -in @('ordinary_gravity','disabled_srm') -and ($vTask[11] -ne 1 -or $vTask[3] -cne 'NOT_EVALUATED')){throw "$($caseTask.name): ordinary-run semantics changed: $($vTask -join '|')"}
  if($caseTask.name -eq 'successful_srm' -and ($vTask[0] -cne 'PASS' -or -not $vTask[13] -or [math]::Abs($vTask[11]-1.1) -gt 1e-12 -or [math]::Abs($vTask[6]-1.1) -gt 1e-12)){throw 'successful SRM control failed'}
  if($failed -and (-not (Test-Path $perf) -or ((Get-Content -LiteralPath $perf -Raw) -notmatch 'FOS_MID=NA') -or ((Get-Content -LiteralPath $perf -Raw) -notmatch 'FOS=UNDETERMINED'))){throw "$($caseTask.name): missing perf summary publication"}
  $recordsTask+=@{name=$caseTask.name;values=$vTask}
 }finally{if($wbTask){$wbTask.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbTask)};$xlTask.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)}
}
if((Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash -ne $hashTask){throw 'Source changed'}
@{status='PASS';source_sha256=$hashTask;checks=$recordsTask}|ConvertTo-Json -Depth 10|Set-Content (Join-Path $OutputRoot 'failed_run_publication_checks.json') -Encoding UTF8
Write-Host 'PASS: failed stage publication clears SRM results only when enabled SRM exists'
