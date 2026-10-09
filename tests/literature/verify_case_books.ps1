param(
  [Parameter(Mandatory=$true)][string]$CasesDirectory,
  [Parameter(Mandatory=$true)][string]$ResultsDirectory,
  [Parameter(Mandatory=$true)][string]$BaselineWorkbook,
  [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Path $OutputDirectory -Force|Out-Null
function ReleaseTask($obj){
  if($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)){
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj)
  }
}
function CodesTask($book){
  $codesTask=@{};$projectTask=$book.VBProject;$componentsTask=$projectTask.VBComponents
  try{
    for($iTask=1;$iTask -le $componentsTask.Count;$iTask++){
      $componentTask=$componentsTask.Item($iTask);$moduleTask=$componentTask.CodeModule
      try{
        $bodyTask=''
        if($moduleTask.CountOfLines -gt 0){$bodyTask=$moduleTask.Lines(1,$moduleTask.CountOfLines)}
        $codesTask[$componentTask.Name]=@{type=[int]$componentTask.Type;body=($bodyTask -replace "`r`n","`n" -replace "`r","`n")}
      }finally{ReleaseTask $moduleTask;ReleaseTask $componentTask}
    }
  }finally{ReleaseTask $componentsTask;ReleaseTask $projectTask}
  return $codesTask
}
$xlTask=$null;$booksTask=$null;$bookTask=$null;$checksTask=[Collections.Generic.List[object]]::new()
try{
  $xlTask=New-Object -ComObject Excel.Application
  $xlTask.Visible=$false;$xlTask.DisplayAlerts=$false;$xlTask.EnableEvents=$false;$xlTask.AutomationSecurity=3
  $booksTask=$xlTask.Workbooks
  $baselineHashTask=(Get-FileHash -LiteralPath $BaselineWorkbook -Algorithm SHA256).Hash
  $bookTask=$booksTask.Open((Resolve-Path -LiteralPath $BaselineWorkbook).Path,0,$true)
  $baselineCodesTask=CodesTask $bookTask
  $bookTask.Close($false);ReleaseTask $bookTask;$bookTask=$null
  foreach($fileTask in Get-ChildItem -LiteralPath $ResultsDirectory -Filter '*.json' -File){
    $resultTask=Get-Content -LiteralPath $fileTask.FullName -Raw -Encoding UTF8|ConvertFrom-Json
    if($null -eq $resultTask.metrics){continue}
    $idTask=[string]$resultTask.case;$pathTask=Join-Path $CasesDirectory ($idTask+'.xlsm')
    $hashTask=(Get-FileHash -LiteralPath $pathTask -Algorithm SHA256).Hash
    $bookTask=$booksTask.Open((Resolve-Path -LiteralPath $pathTask).Path,0,$true)
    $codesTask=CodesTask $bookTask
    $bookTask.Close($false);ReleaseTask $bookTask;$bookTask=$null
    $checksTask.Add(@{case=$idTask;check='no_test_harness';ok=(-not $codesTask.ContainsKey('LiteratureHarness'))})
    $sameNamesTask=(($codesTask.Keys|Sort-Object)-join '|') -ceq (($baselineCodesTask.Keys|Sort-Object)-join '|')
    $checksTask.Add(@{case=$idTask;check='component_names';ok=$sameNamesTask;components=$codesTask.Count})
    foreach($nameTask in $baselineCodesTask.Keys|Sort-Object){
      $okTask=$codesTask.ContainsKey($nameTask)
      if($okTask){$okTask=($codesTask[$nameTask].type -eq $baselineCodesTask[$nameTask].type -and $codesTask[$nameTask].body -ceq $baselineCodesTask[$nameTask].body)}
      $checksTask.Add(@{case=$idTask;check=('vba:'+ $nameTask);ok=$okTask})
    }
    $checksTask.Add(@{case=$idTask;check='read_only_hash';ok=($hashTask -ceq (Get-FileHash -LiteralPath $pathTask -Algorithm SHA256).Hash);sha256=$hashTask})
  }
  $checksTask.Add(@{case='BASELINE';check='read_only_hash';ok=($baselineHashTask -ceq (Get-FileHash -LiteralPath $BaselineWorkbook -Algorithm SHA256).Hash);sha256=$baselineHashTask})
}catch{$checksTask.Add(@{case='ALL';check='exception';ok=$false;detail=$_.Exception.Message})}
finally{
  if($null -ne $bookTask){$bookTask.Close($false);ReleaseTask $bookTask}
  ReleaseTask $booksTask
  if($null -ne $xlTask){$xlTask.Quit();ReleaseTask $xlTask}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
$failureTask=@($checksTask|Where-Object {-not $_.ok})
@{ok=($failureTask.Count -eq 0);check_count=$checksTask.Count;failure_count=$failureTask.Count;checks=$checksTask}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $OutputDirectory 'verify_vba.json') -Encoding UTF8
Write-Host ('VBA checks='+$checksTask.Count+' failures='+$failureTask.Count)
if($failureTask.Count -gt 0){$failureTask|ConvertTo-Json -Depth 5;exit 1}
