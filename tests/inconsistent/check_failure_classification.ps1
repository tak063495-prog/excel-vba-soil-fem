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
foreach($fixedTask in @($false,$true)){
foreach($faultTask in @(1,2,3)){
 $pathTask=Join-Path $OutputRoot ('fault_'+$fixedTask+'_'+$faultTask+'.xlsm')
 Copy-Item -LiteralPath $sourceTask -Destination $pathTask -Force
 $xlTask=New-Object -ComObject Excel.Application;$xlTask.DisplayAlerts=$false;$xlTask.EnableEvents=$false;$xlTask.AutomationSecurity=1;$wbTask=$null
 try{
  $wbTask=$xlTask.Workbooks.Open([IO.Path]::GetFullPath($pathTask),0,$false)
  $cmTask=$wbTask.VBProject.VBComponents.Item('FEMSolver').CodeModule
  $textTask=$cmTask.Lines(1,$cmTask.CountOfLines)
  $textTask=$textTask.Replace('Option Explicit',"Option Explicit`r`nPublic IncoFault As Long")
  $needleTask='Function BandSolver() As Boolean'
  if(-not $textTask.Contains($needleTask)){throw 'Solver hook absent'}
  $injectionTask=@'
Function BandSolver() As Boolean
  If IncoFault > 0 And P3SrmTrialRunning And P3CurrentStrengthFactor > 1.05 Then
    If IncoFault = 1 Then
      SetAnalysisFailure RESULT_NONCONVERGED, "Injected linear solver rejection", vbObjectError + 3202, -1, -1, CurrentIncrement, CurrentIteration
    ElseIf IncoFault = 2 Then
      SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "Injected singular tangent", vbObjectError + 3301, -1, -1, CurrentIncrement, CurrentIteration
    Else
      SetAnalysisFailure RESULT_CAPACITY_ERROR, "Injected capacity failure", vbObjectError + 3511, -1, -1, CurrentIncrement, CurrentIteration
    End If
    BandSolver = False
    Exit Function
  End If
'@
  $textTask=$textTask.Replace($needleTask,($injectionTask -replace "`r?`n","`r`n"))
  $cmTask.DeleteLines(1,$cmTask.CountOfLines);$cmTask.AddFromString($textTask)
  $cmTask.AddFromString(@'
Public Sub IncoSetFault(ByVal value As Long)
  IncoFault = value
End Sub
'@)
  $hTask=$wbTask.VBProject.VBComponents.Add(1);$hTask.Name='IncoFailureHarness'
  $hTask.CodeModule.AddFromString(@'
Public Function IncoFailureRun() As Variant
  SetP0SilentMode True
  P0_RunAnalysis
  IncoFailureRun = Array(ResultStatus, P3FosInterpretation, P3SrmTrialClassification, P3FailureKind, P3FosPass, P3FosFail, P3FosBracket, FSS, P3SrmTrialCount)
End Function
'@)
  RowsTask $wbTask '材料データ' $fixtureTask.materials 13
  RowsTask $wbTask '節点データ' $fixtureTask.nodes 9
  RowsTask $wbTask '要素データ' $fixtureTask.elements 10
  RowsTask $wbTask '載荷' @() 10;RowsTask $wbTask '接合' @() 5
  RowsTask $wbTask 'ステージ' @(@(1,'GRAVITY',$null,$null,1,'gravity'),@(2,'SRM',$null,$null,1,'srm')) 6
  $prefixTask="'"+$wbTask.Name+"'!"
  foreach($keyTask in @('ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V2B_COST','ACCEL_V3_COST_SEARCH','ACCEL_V4_ANDERSON','ACCEL_V5_GMRES_LU','ACCEL_ADAPTIVE','ACCEL_STEP_RECOVERY','SRM_FIXED_FS')){$xlTask.Run($prefixTask+'FEMWriteNumericSetting',$keyTask,0)}
  $xlTask.Run($prefixTask+'FEMWriteTextSetting','FLOW_POLICY','INCONSISTENT')
  $xlTask.Run($prefixTask+'FEMWriteTextSetting','DEBUG_MODE','STAGE')
  $xlTask.Run($prefixTask+'FEMWriteTextSetting','EXPORT_MODE','OFF')
  if($fixedTask){$xlTask.Run($prefixTask+'FEMWriteNumericSetting','SRM_FIXED_FS',1.2)}
  $xlTask.Run($prefixTask+'IncoSetFault',$faultTask)
  $vTask=@($xlTask.Run($prefixTask+'IncoFailureRun'))
  $expectedPassTask=$(if($fixedTask){0}else{1});$expectedTrialsTask=$(if($fixedTask){1}else{2})
  if($vTask[0] -eq 'PASS' -or $vTask[1] -ne 'UNDETERMINED' -or $vTask[2] -ne 'NUMERICAL_FAILURE' -or $vTask[3] -ne 'LINEAR_SOLVER' -or $vTask[4] -ne $expectedPassTask -or $vTask[5] -ne 0 -or $vTask[6] -or $vTask[7] -ne 0 -or $vTask[8] -ne $expectedTrialsTask){throw ('Fault '+$faultTask+' unsafe classification: '+($vTask -join '|'))}
  $recordsTask+=@{fixed_fs=$fixedTask;fault=$faultTask;values=$vTask}
 }finally{if($null -ne $wbTask){$wbTask.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbTask)};$xlTask.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)}
}
}
if((Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash -ne $hashTask){throw 'Source changed'}
@{status='PASS';source_sha256=$hashTask;records=$recordsTask}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $OutputRoot 'failure_checks.json') -Encoding UTF8
Write-Host 'PASS: linear/singular/capacity failures do not establish an SRM upper bound'