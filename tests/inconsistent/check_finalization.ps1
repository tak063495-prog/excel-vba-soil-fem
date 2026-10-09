param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force $OutputRoot|Out-Null
$sourceTask=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$hashTask=(Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash
$fixtureTask=(Get-Content (Join-Path $PSScriptRoot '../literature/fixtures.json') -Raw -Encoding UTF8|ConvertFrom-Json).fixtures|Where-Object name -eq 'confined_gravity'
function RowsTask($book,$sheet,$rows,$columns){
 $ws=$book.Worksheets.Item($sheet);$ws.Range($ws.Cells(2,1),$ws.Cells(2000,$columns)).ClearContents()
 if($rows.Count -eq 0){return}
 $data=[object[,]]::new($rows.Count,$columns)
 for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $columns;$c++){ $value=$rows[$r][$c];if($null -ne $value -and $value -isnot [string]){$value=[double]$value};$data[$r,$c]=$value}}
 $ws.Range($ws.Cells(2,1),$ws.Cells($rows.Count+1,$columns)).Value2=$data
}
$recordsTask=@()
foreach($faultTask in @(1,2)){
 $copyTask=Join-Path $OutputRoot ('finalization_'+$faultTask+'.xlsm')
 Copy-Item -LiteralPath $sourceTask -Destination $copyTask -Force
 $xlTask=New-Object -ComObject Excel.Application
 $xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1;$wbTask=$null
 try{
  $wbTask=$xlTask.Workbooks.Open([IO.Path]::GetFullPath($copyTask),0,$false)
  $cmTask=$wbTask.VBProject.VBComponents.Item('FEMEngine').CodeModule
  $sTask=$cmTask.Lines(1,$cmTask.CountOfLines)
  $sTask=$sTask.Replace('Option Explicit',"Option Explicit`r`nPublic IncoFinalFault As Long`r`nPrivate IncoAtFinalize As Boolean")
  $needleTask='Private Function P3TryStrengthFactor(ByVal strengthFactor As Double) As Boolean'
  $injectionTask=@'
Private Function P3TryStrengthFactor(ByVal strengthFactor As Double) As Boolean
  If IncoFinalFault > 0 Then
    If IncoAtFinalize Then
      SetAnalysisFailure RESULT_CAPACITY_ERROR, "Injected final replay capacity error", vbObjectError + 3511, -1, -1, 1, 0
      P3SrmTrialClassification = "NONCONVERGENCE_BOUNDARY"
      P3TryStrengthFactor = False
      Exit Function
    ElseIf strengthFactor > 1.01 Then
      SetAnalysisFailure RESULT_NONCONVERGED, "Synthetic boundary only for control test", vbObjectError + 3202, -1, -1, 1, 0
      P3SrmTrialClassification = "NONCONVERGENCE_BOUNDARY"
      P3FailureKind = "GLOBAL_ITERATION_LIMIT": P3FailurePlasticPoints = 1
      P3TryStrengthFactor = False
      Exit Function
    End If
  End If
'@
  if(-not $sTask.Contains($needleTask)){throw 'Trial hook absent'}
  $sTask=$sTask.Replace($needleTask,($injectionTask -replace "`r?`n","`r`n"))
  $needleTask='  If P3SrmSnapReady And Abs(P3SrmSnapFs - fLow) <= 0.000000001 Then'
  $injectionTask=@'
  If IncoFinalFault > 0 Then
    IncoAtFinalize = True
    FSS = 9#: P3FosFail = 9#: P3FosMid = 9#: P3FosWidth = 9#: P3FosBracket = True
    If IncoFinalFault = 2 Then P3SrmSnapReady = False
  End If
'@
  if(-not $sTask.Contains($needleTask)){throw 'Finalization hook absent'}
  $sTask=$sTask.Replace($needleTask,(($injectionTask -replace "`r?`n","`r`n")+"`r`n"+$needleTask))
  $needleTask='Private Function P3RestoreSrmSuccessSnapshot() As Boolean'
  if(-not $sTask.Contains($needleTask)){throw 'Snapshot hook absent'}
  $sTask=$sTask.Replace($needleTask,($needleTask+"`r`n  If IncoFinalFault = 1 And IncoAtFinalize Then P3RestoreSrmSuccessSnapshot = False: Exit Function"))
  $cmTask.DeleteLines(1,$cmTask.CountOfLines);$cmTask.AddFromString($sTask)
  $cmTask.AddFromString(@'
Public Function IncoFinalRun(ByVal faultId As Long) As Variant
  SetP0SilentMode True
  IncoFinalFault = faultId: IncoAtFinalize = False
  P0_RunAnalysis
  IncoFinalRun = Array(ResultStatus, P3FosInterpretation, P3SrmTrialClassification, P3FailureKind, P3FosPass, P3FosFail, P3FosMid, P3FosWidth, P3FosBracket, FSS, IncoAtFinalize, P3FosDisplayFs())
End Function
'@)
  RowsTask $wbTask '材料データ' $fixtureTask.materials 13
  RowsTask $wbTask '節点データ' $fixtureTask.nodes 9
  RowsTask $wbTask '要素データ' $fixtureTask.elements 10
  RowsTask $wbTask '載荷' @() 10;RowsTask $wbTask '接合' @() 5
  RowsTask $wbTask 'ステージ' @(@(1,'GRAVITY',$null,$null,1,'gravity'),@(2,'SRM',$null,$null,1,'srm')) 6
  $prefixTask="'"+$wbTask.Name+"'!"
  foreach($keyTask in @('ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V2B_COST','ACCEL_V3_COST_SEARCH','ACCEL_V4_ANDERSON','ACCEL_V5_GMRES_LU','ACCEL_ADAPTIVE','ACCEL_STEP_RECOVERY','SRM_FIXED_FS')){$xlTask.Run($prefixTask+'FEMWriteNumericSetting',$keyTask,0)}
  $xlTask.Run($prefixTask+'FEMWriteNumericSetting','SRM_TOL',.0125)
  $xlTask.Run($prefixTask+'FEMWriteTextSetting','FLOW_POLICY','INCONSISTENT')
  $xlTask.Run($prefixTask+'FEMWriteTextSetting','EXPORT_MODE','OFF')
  $vTask=@($xlTask.Run($prefixTask+'IncoFinalRun',$faultTask))
  $expectedStatusTask=$(if($faultTask -eq 1){'NONCONVERGED'}else{'CAPACITY_ERROR'})
  $expectedKindTask=$(if($faultTask -eq 1){'STATE_RESTORE'}else{'FINAL_VERIFICATION'})
  if($vTask[0] -ne $expectedStatusTask -or $vTask[1] -ne 'UNDETERMINED' -or $vTask[2] -ne 'NUMERICAL_FAILURE' -or $vTask[3] -ne $expectedKindTask -or $vTask[5] -ne 0 -or $vTask[6] -ne 0 -or $vTask[7] -ne 0 -or $vTask[8] -or $vTask[9] -ne 0 -or -not $vTask[10] -or $vTask[11] -ne 0){throw ('Unsafe finalization: '+($vTask -join '|'))}
  $recordsTask+=@{fault=$faultTask;values=$vTask}
 }finally{if($wbTask){$wbTask.Close($false)};$xlTask.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)}
}
if((Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash -ne $hashTask){throw 'Source changed'}
@{status='PASS';source_sha256=$hashTask;records=$recordsTask}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $OutputRoot 'finalization_checks.json') -Encoding UTF8
Write-Host 'PASS: snapshot and final replay failure clear stale FOS and preserve capacity status'
