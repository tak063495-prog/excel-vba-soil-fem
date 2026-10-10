$ErrorActionPreference='Stop'
$rootTask=Split-Path -Parent $PSScriptRoot
Push-Location $rootTask
try {
$copyTask=Join-Path $pwd 'tests/tmp/practical_final.xlsm'
Copy-Item -LiteralPath workbook/2DSoilFEM_20261008_practical.xlsm -Destination $copyTask -Force
$checksTask=[Collections.Generic.List[string]]::new()
function Check($value,$label){if(-not $value){throw $label};$checksTask.Add('PASS '+$label)}
$xlTask=New-Object -ComObject Excel.Application
try {
  $xlTask.EnableEvents=$true;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1
  $wbTask=$xlTask.Workbooks.Open($copyTask,0,$false)
  Check ($wbTask.ActiveSheet.Name -eq '操作パネル') 'Workbook_Open shows operation panel'
  Check ($xlTask.EnableEvents) 'Workbook_Open restores events'
  Check ($wbTask.Worksheets.Item('高速化設定').ProtectContents) 'settings protected after reopen'
  Check (-not $wbTask.Worksheets.Item('高速化設定').Range('C12').Locked) 'setting input editable after reopen'
  Check ($null -eq $wbTask.LinkSources(1)) 'no external workbook links'
  $tmTask=$wbTask.VBProject.VBComponents.Add(1);$tmTask.Name='FinalUiHarness'
  $tmTask.CodeModule.AddFromString(@'
Public Function FinalFormulaErrors() As String
  Dim ws As Worksheet, cells As Range, cell As Range
  ' A nonempty success sentinel avoids marshaling a null BSTR in PowerShell 7.
  FinalFormulaErrors = "PASS"
  For Each ws In ThisWorkbook.Worksheets
    Set cells = Nothing
    On Error Resume Next
    Set cells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not cells Is Nothing Then
      For Each cell In cells
        If IsError(cell.Value2) Then FinalFormulaErrors = FinalFormulaErrors & ws.Name & "!" & cell.Address & ";"
      Next cell
    End If
  Next ws
End Function
Public Function FinalBuildShape() As String
  On Error GoTo Rejected
  FEMUiBuildGeometry
  FinalBuildShape = "PASS"
  Exit Function
Rejected:
  FinalBuildShape = "REJECTED|" & Err.Description
End Function
'@)
  $prefixTask="'"+$wbTask.Name+"'!"
  $xlTask.CalculateFull()
  $formulaErrorsTask=[string]$xlTask.Run($prefixTask+'FinalFormulaErrors')
  Check ($formulaErrorsTask -eq 'PASS') ('all native Excel formula results have no errors: '+$formulaErrorsTask)
  $mappingTask=Get-Content tests/fixtures/ui_mapping.json -Raw -Encoding UTF8 | ConvertFrom-Json
  foreach($keyTask in @('ACCEL_PREDICT_BETA','ACCEL_PREDICT_MAX_RATIO','ACCEL_TANGENT_CHANGE','SRM_FIXED_FS')){
    $mTask=$mappingTask | Where-Object {$_.key -eq $keyTask}
    $cellTask=$wbTask.Worksheets.Item($mTask.sheet).Range($mTask.cell)
    $originalTask=[double]$cellTask.Value2
    $cellTask.Value2=[double]-1
    Check (-not $cellTask.Validation.Value) ('custom rule rejects invalid value '+$keyTask)
    $cellTask.Value2=$originalTask
    Check $cellTask.Validation.Value ('custom rule accepts original value '+$keyTask)
  }
  $xlTask.EnableEvents=$false
  $shapeTask=$wbTask.Worksheets.Item('形状入力');$legacyTask=$wbTask.Worksheets.Item('図形定義')
  $shapeTask.Range('B8:D2000').ClearContents()|Out-Null;$shapeTask.Range('G8:J2000').ClearContents()|Out-Null
  $pointsTask=@(@(50.0,1.0,1.0),@(60.0,1.0,2.0),@(70.0,2.0,2.0),@(80.0,2.0,1.0),@(10.0,0.0,0.0),@(20.0,0.0,3.0),@(30.0,3.0,3.0),@(40.0,3.0,0.0))
  for($rTask=0;$rTask -lt 8;$rTask++){for($cTask=0;$cTask -lt 3;$cTask++){$shapeTask.Cells($rTask+8,$cTask+2).Value2=[double]$pointsTask[$rTask][$cTask]}}
  $shapeTask.Range('H8').Value2='内側';$shapeTask.Range('I8').Value2='50,60,70,80';$shapeTask.Range('J8').Value2=[double]0.5
  $shapeTask.Range('H9').Value2='外側';$shapeTask.Range('I9').Value2='10,20,30,40';$shapeTask.Range('J9').Value2=[double]0.5
  $legacyTask.Range('D24').Value2=[double]2;$legacyTask.Range('G25').Value2=[double]999
  Check ($xlTask.Run($prefixTask+'FinalBuildShape') -eq 'PASS') 'arbitrary IDs and hole before outer boundary accepted'
  Check ($legacyTask.Range('H1').Value2 -eq 4 -and $legacyTask.Range('J1').Value2 -eq 4 -and $legacyTask.Range('L1').Value2 -eq 8) 'outer/inner/total coordinate counts'
  Check ($legacyTask.Range('D3').Value2 -eq 1 -and $legacyTask.Range('D24').Value2 -eq 1) 'outer/inner loop counts'
  Check ($legacyTask.Range('G3').Value2 -eq 1 -and $legacyTask.Range('J3').Value2 -eq 4 -and $legacyTask.Range('G24').Value2 -eq 5 -and $legacyTask.Range('J24').Value2 -eq 8) 'boundary IDs remapped to native schema'
  Check ($legacyTask.Range('B2').Value2 -eq 0 -and $legacyTask.Range('B6').Value2 -eq 1) 'outer coordinates precede inner coordinates'
  Check ($null -eq $legacyTask.Range('G25').Value2) 'previous extra inner boundary cleared'
  $snapshotTask=ConvertTo-Json -InputObject $legacyTask.Range('A1:Q26').Value2 -Compress
  foreach($variantTask in @('duplicate_coordinate','unknown_boundary_id','zero_spacing','missing_coordinate','duplicate_boundary_id')){
    switch($variantTask){
      'duplicate_coordinate' {$shapeTask.Range('B9').Value2=[double]50}
      'unknown_boundary_id' {$shapeTask.Range('I8').Value2='50,60,70,999'}
      'zero_spacing' {$shapeTask.Range('J8').Value2=[double]0}
      'missing_coordinate' {$shapeTask.Range('C8').ClearContents()|Out-Null}
      'duplicate_boundary_id' {$shapeTask.Range('I8').Value2='50,60,70,50'}
    }
    Check ($xlTask.Run($prefixTask+'FinalBuildShape').StartsWith('REJECTED|')) ('reject '+$variantTask)
    Check ((ConvertTo-Json -InputObject $legacyTask.Range('A1:Q26').Value2 -Compress) -ceq $snapshotTask) ('rejection preserves backend '+$variantTask)
    Check (-not $xlTask.EnableEvents) ('rejection preserves event mode '+$variantTask)
    $shapeTask.Range('B9').Value2=[double]60;$shapeTask.Range('I8').Value2='50,60,70,80';$shapeTask.Range('J8').Value2=[double]0.5;$shapeTask.Range('C8').Value2=[double]1
  }
  $exportTask=Join-Path $pwd 'tests/tmp/modules';New-Item -ItemType Directory -Path $exportTask -Force|Out-Null
  foreach($compTask in $wbTask.VBProject.VBComponents){
    if($compTask.Name -eq 'FinalUiHarness'){continue}
    $cmTask=$compTask.CodeModule
    if($cmTask.CountOfLines){[IO.File]::WriteAllText((Join-Path $exportTask ($compTask.Name+'.txt')),$cmTask.Lines(1,$cmTask.CountOfLines),[Text.Encoding]::UTF8)}
  }
  $checksTask | Set-Content tests/results/practical_final_checks_20261008.txt -Encoding UTF8
  Write-Output ('PASS '+$checksTask.Count+' reopen, validation, formula, shape mapping and error preservation checks.')
  $wbTask.Close($false)
} finally {
  if($null -ne $wbTask){try{$wbTask.Close($false)}catch{}}
  try{$xlTask.EnableEvents=$false;$xlTask.Quit()}catch{}
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)|Out-Null
}

} finally { Pop-Location }
