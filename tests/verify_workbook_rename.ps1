param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Reference,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Existing Excel; do not dispatch.'}
$Book=(Resolve-Path $Book).Path;$Reference=(Resolve-Path $Reference).Path
if((Get-FileHash $Book).Hash -ne (Get-FileHash $Reference).Hash){throw 'Rename must preserve workbook bytes.'}
if(Test-Path $Out){throw 'Preserve existing evidence.'}
$Out=[IO.Path]::GetFullPath($Out);New-Item -ItemType Directory -Path $Out | Out-Null
$copy=Join-Path $Out (Split-Path $Book -Leaf);Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$modules=(Get-Content (Join-Path $PSScriptRoot '../workbook/ui_20261022/BUILD_RECORD.json') -Raw -Encoding UTF8 | ConvertFrom-Json).modules
function Norm([string]$s){(($s -replace "`r`n","`n" -replace "`r","`n") -split "`n" | Where-Object {$_ -notmatch '^Attribute VB_'} | ForEach-Object {$_.TrimEnd()}) -join "`n" -replace "`n+$",""}
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.AutomationSecurity=1;$xl.EnableEvents=$true
 $wb=$xl.Workbooks.Open($copy,0,$false)
 $xl.EnableEvents=$false
 foreach($m in $modules){
  $ext=if($m -eq 'ThisWorkbook'){'.cls'}else{'.bas'}
  $source=[IO.File]::ReadAllText((Join-Path $PSScriptRoot "../src/vba/$m$ext"))
  $cm=$wb.VBProject.VBComponents.Item($m).CodeModule
  if((Norm $source) -ne (Norm $cm.Lines(1,$cm.CountOfLines))){throw "Source mismatch: $m"}
 }
 $count=0;$prefix="'"+$wb.Name+"'!"
 foreach($ws in $wb.Worksheets){foreach($shape in $ws.Shapes){
  $action='';try{$action=[string]$shape.OnAction}catch{}
  if($action){if(-not $action.StartsWith($prefix)){throw "Action points outside renamed book: $action"};$count++}
 }}
 if($count -lt 15){throw 'Missing navigation actions.'}
 if($wb.Worksheets.Item('水条件').Range('B3').Value2 -ne 'NONE'){throw 'Unexpected release default.'}
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()}
 if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Compile failed.'}
 $xl.Visible=$true;$xl.WindowState=-4143;$xl.Width=1050;$xl.Height=650
 $null=$xl.Run($prefix+'FEMUiPreviewMesh')
 if($wb.ActiveSheet.Name -ne '図' -or $wb.ActiveSheet.Shapes.Count -eq 0){throw 'Preview not displayed.'}
 $xl.ActiveWindow.Zoom=60
 @{status='PASS';book_sha256=(Get-FileHash $Book).Hash;reference_sha256=(Get-FileHash $Reference).Hash;byte_identical=$true;name=$wb.Name;compile='PASS';source_modules=$modules;actions_rebound=$count;normal_open=$true;preview_shapes=$wb.ActiveSheet.Shapes.Count;excel_version=$xl.Version;excel_build=$xl.Build}|ConvertTo-Json -Depth 6 | Set-Content (Join-Path $Out 'rename_check.json') -Encoding UTF8
 # The caller captures this owned window, then writes close.txt to finish.
 while(-not(Test-Path (Join-Path $Out 'close.txt'))){Start-Sleep -Milliseconds 250}
 $wb.Close($false)
}finally{
 foreach($obj in @($ctl,$cm,$wb)){if($obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj)}catch{}}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
