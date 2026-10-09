param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputRoot|Out-Null
$hashTask=(Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash
$copyTask=Join-Path ([IO.Path]::GetTempPath()) ('lm_'+[guid]::NewGuid().ToString('N')+'.xlsm')
Copy-Item -LiteralPath $SourceWorkbook -Destination $copyTask
$excelTask=New-Object -ComObject Excel.Application
$excelTask.Visible=$false;$excelTask.DisplayAlerts=$false;$excelTask.EnableEvents=$false;$excelTask.AutomationSecurity=1
$bookTask=$null
try{
 $bookTask=$excelTask.Workbooks.Open($copyTask,0,$false)
 $sheetTask=$bookTask.Worksheets.Add();$sheetTask.Name='lm_test'
 $bookTask.VBProject.VBComponents.Item('FEMSolver').CodeModule.AddFromString(@'
Public Function LMHarness(ByVal n As Long, ByVal mu As Double, ByVal bw As Long, ByVal budget As Double) As Variant
 On Error GoTo Failed
 SetP0SilentMode True
 Dim i As Long, j As Long, p As Long, ok As Boolean, same As Boolean, sheet As Worksheet
 Set sheet=ThisWorkbook.Worksheets("lm_test")
 nDof=n:lastDof=n-1:BandWidth=bw:P6CSRNNZ=n*n:P6CSRStorageBytes=0#
 ReDim P6CSRRowPtr(n):ReDim P6CSRColumnIndex(n*n-1):ReDim P6CSRValues(n*n-1):ReDim P6CSROriginalValues(n*n-1)
 ReDim Force(n-1):ReDim Disp(n-1):ReDim NodeCond(n-1)
 For i=0 To n-1
  P6CSRRowPtr(i)=i*n:Force(i)=sheet.Cells(i+1,n+1).Value2:NodeCond(i)=sheet.Cells(i+1,n+2).Value2:Disp(i)=77#+i
  For j=0 To n-1
   p=i*n+j:P6CSRColumnIndex(p)=j:P6CSROriginalValues(p)=sheet.Cells(i+1,j+1).Value2:P6CSRValues(p)=P6CSROriginalValues(p)
  Next j
 Next i
 P6CSRRowPtr(n)=n*n:P6CSRReady=True:P6SolverMemoryLimitBytes=budget:P6IterativeTolerance=1E-8
 P6IterativeLastResidual=.123:P6IterativeLastIterations=47
 ok=P6SolveResidualRecovery(mu):same=True
 For p=0 To n*n-1
  If P6CSRValues(p)<>P6CSROriginalValues(p) Then same=False
 Next p
 LMHarness=Array(ok,Disp,same,P6IterativeLastResidual,P6IterativeLastIterations)
 Exit Function
Failed:
 LMHarness=Array("ERROR",CStr(Err.Number)&" "&Err.Description)
End Function
'@)
 $casesTask=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'regularized_cases.json') -Raw|ConvertFrom-Json
 $resultsTask=@()
 foreach($caseTask in $casesTask){
  $nTask=$caseTask.rhs.Count
  for($iTask=0;$iTask -lt $nTask;$iTask++){
   for($jTask=0;$jTask -lt $nTask;$jTask++){$sheetTask.Cells.Item($iTask+1,$jTask+1).Value2=[double]$caseTask.A[$iTask][$jTask]}
   $sheetTask.Cells.Item($iTask+1,$nTask+1).Value2=[double]$caseTask.rhs[$iTask]
   $sheetTask.Cells.Item($iTask+1,$nTask+2).Value2=[double]([int]($caseTask.fixed -contains $iTask))
  }
  $vTask=@($excelTask.Run("'"+$bookTask.Name+"'!LMHarness",$nTask,[double]$caseTask.damping,[int]$caseTask.band,[double]$caseTask.budget))
  if([string]$vTask[0] -ceq 'ERROR'){throw $vTask[1]}
  if([bool]$vTask[0] -ne $caseTask.ok){throw ('Status '+$caseTask.name+' damping='+$caseTask.damping)}
  if(-not [bool]$vTask[2] -or [double]$vTask[3] -ne .123 -or [int]$vTask[4] -ne 47){throw 'Plastic operator or primary diagnostic changed'}
  $gapTask=0.;$normTask=0.
  for($iTask=0;$iTask -lt $nTask;$iTask++){$gapTask+=([double]$vTask[1][$iTask]-$caseTask.expected[$iTask])*([double]$vTask[1][$iTask]-$caseTask.expected[$iTask]);$normTask+=$caseTask.expected[$iTask]*$caseTask.expected[$iTask]}
  $gapTask=[math]::Sqrt($gapTask/[math]::Max(1e-30,$normTask))
  if($gapTask -gt 2e-7){throw ('Incorrect regularized direction '+$caseTask.name+' gap='+$gapTask)}
  $resultsTask+=@{case=$caseTask.name;damping=$caseTask.damping;ok=$vTask[0];relative_gap=$gapTask}
 }
 [IO.File]::WriteAllText((Join-Path $OutputRoot 'regularized_checks.json'),(@{status='PASS';source_sha256=$hashTask;count=$resultsTask.Count;results=$resultsTask}|ConvertTo-Json -Depth 6))
 Write-Output ('PASS: '+$resultsTask.Count+' independent regularized solves and rejection checks')
}finally{if($bookTask){$bookTask.Close($false)};$excelTask.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excelTask);Remove-Item -LiteralPath $copyTask -Force -ErrorAction SilentlyContinue}
if((Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash -ne $hashTask){throw 'Source changed'}
