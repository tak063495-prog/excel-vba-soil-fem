param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force $OutputRoot|Out-Null
$sourceTask=(Resolve-Path -LiteralPath $SourceWorkbook).Path;$hashTask=(Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash
$copyTask=Join-Path $OutputRoot 'fd_guard_control.xlsm';Copy-Item -LiteralPath $sourceTask -Destination $copyTask -Force
$xlTask=New-Object -ComObject Excel.Application;$xlTask.DisplayAlerts=$false;$xlTask.EnableEvents=$false;$xlTask.AutomationSecurity=1;$wbTask=$null
try{
 $wbTask=$xlTask.Workbooks.Open([IO.Path]::GetFullPath($copyTask),0,$false)
 $cmTask=$wbTask.VBProject.VBComponents.Item('FEMEngine').CodeModule;$sTask=$cmTask.Lines(1,$cmTask.CountOfLines)
 $sTask=$sTask.Replace('Option Explicit',"Option Explicit`r`nPrivate IncoFdMode As Long, IncoFdCalls As Long")
 $needleTask='Private Function P2FiniteDifferenceColumn(ByRef inputState As P2_MaterialPointInput, ByRef baseState As P2_MaterialPointOutput, ByVal stepCount As Long, ByVal col As Long, ByVal h As Double, ByRef columnValue() As Double, ByRef stencilKind As Long) As Boolean'
 if(-not $sTask.Contains($needleTask)){throw 'FD guard hook absent'}
 $hookTask=@'
  If IncoFdMode > 0 Then
    IncoFdCalls = IncoFdCalls + 1
    columnValue(0) = 0#: columnValue(1) = 0#: columnValue(2) = 0#: columnValue(col) = 1000#
    stencilKind = 2
    If IncoFdMode = 1 And IncoFdCalls Mod 2 = 1 Then stencilKind = 1
    P2FiniteDifferenceColumn = True: Exit Function
  End If
'@
 $sTask=$sTask.Replace($needleTask,($needleTask+"`r`n"+($hookTask -replace "`r?`n","`r`n")))
 $cmTask.DeleteLines(1,$cmTask.CountOfLines);$cmTask.AddFromString($sTask)
 $cmTask.AddFromString(@'
Public Function IncoFdGuardRun(ByVal modeId As Long) As Variant
  Dim i As P2_MaterialPointInput, o As P2_MaterialPointOutput, row As Long, col As Long, ok As Boolean, errorValue As Double
  i.Young = 1000#: i.Poisson = 0.3: i.cohesion = 1#: i.tolerance = 1E-10
  i.PreviousStress(0) = 12#: i.PreviousStress(1) = 4#: i.PreviousStress(3) = 5#
  For row = 0 To 3: o.Stress(row) = i.PreviousStress(row): Next row
  For row = 0 To 2: For col = 0 To 2: o.tangent(row,col) = -19#: Next col: Next row
  IncoFdMode = modeId: IncoFdCalls = 0
  ok = P2FillFiniteDifferenceTangent(i,o,4)
  For row = 0 To 2
    For col = 0 To 2
      If modeId = 1 Then
        errorValue = errorValue + Abs(o.tangent(row,col)+19#)
      ElseIf row = col Then
        errorValue = errorValue + Abs(o.tangent(row,col)-1000#)
      Else
        errorValue = errorValue + Abs(o.tangent(row,col))
      End If
    Next col
  Next row
  IncoFdGuardRun = Array(ok,errorValue,IncoFdCalls)
End Function
'@)
 $compileTask=$xlTask.VBE.CommandBars.FindControl(1,578);if($compileTask.Enabled){$compileTask.Execute()}
 $rowsTask=@()
 foreach($modeTask in @(1,2)){
  $vTask=@($xlTask.Run(("'"+$wbTask.Name+"'!IncoFdGuardRun"),$modeTask))
  if($vTask[0] -ne ($modeTask -eq 2) -or $vTask[1] -ne 0 -or $vTask[2] -lt 2){throw ('FD stencil guard failed '+($vTask -join '|'))}
  $rowsTask+=@{mode=$modeTask;values=$vTask}
 }
 if((Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash -ne $hashTask){throw 'Source changed'}
 @{status='PASS';source_sha256=$hashTask;description='Native atomic FD commit and stencil agreement with synthetic columns';records=$rowsTask}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $OutputRoot 'fd_guard_checks.json') -Encoding UTF8
 Write-Host 'PASS: mixed stencils rejected atomically, matching one-sided stencils accepted'
}finally{if($wbTask){$wbTask.Close($false)};$xlTask.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)}
