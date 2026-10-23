param([Parameter(Mandatory=$true)][string]$SrmBook,[Parameter(Mandatory=$true)][string]$SeepageBook,[Parameter(Mandatory=$true)][string]$Out,[switch]$Visible)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
if(Test-Path -LiteralPath $Out){throw 'Preserve previous evidence.'}
$Out=[IO.Path]::GetFullPath($Out);New-Item -ItemType Directory -Path $Out | Out-Null
$srm=Join-Path $Out 'mixed_srm_display.xlsm';$seep=Join-Path $Out 'mixed_seepage_display.xlsm'
Copy-Item -LiteralPath $SrmBook -Destination $srm;Copy-Item -LiteralPath $SeepageBook -Destination $seep
Copy-Item -LiteralPath (Join-Path (Split-Path ([IO.Path]::GetFullPath($SrmBook))) 'soilfem_native_x64.dll') -Destination $Out
$checks=[Collections.Generic.List[object]]::new()
function Assert($value,[string]$label){if(-not $value){throw "FAIL: $label"};$checks.Add(@{label=$label;status='PASS'});"PASS $label"}
$xl=New-Object -ComObject Excel.Application;$wb=$null;$wet=$null
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$true;$xl.AutomationSecurity=1;$xl.Visible=[bool]$Visible
 $wb=$xl.Workbooks.Open($srm,0,$false);$xl.EnableEvents=$false
 $mod=$wb.VBProject.VBComponents.Add(1);$mod.Name='MixedOutputDriver'
 $mod.CodeModule.AddFromString(@'
Public Function M6Draw(ByVal mode As String, ByVal resultId As Long) As String
 On Error GoTo Failed
 SuppressUserMessages = True
 FEMWriteNumericSetting "VIEW_COLOR_RESULT", CDbl(resultId)
 FEMWriteNumericSetting "VIEW_DISP_SCALE", 0#
 ThisWorkbook.Worksheets("出力・表示").Range("J28:J29").Value2 = True
 ThisWorkbook.Worksheets("出力・表示").Range("C27").Value2 = IIf(mode = "A", "A：選択結果へ等値線を追加", "B：増分専用カラー")
 FEMViewResult
 If Len(FEMViewerLastError) > 0 Then Err.Raise vbObjectError + 3988, , FEMViewerLastError
 If Len(FEMIncrementLastError) > 0 Then Err.Raise vbObjectError + 3988, , FEMIncrementLastError
 If Len(FEMViewerLastSvg) < 1000 Then Err.Raise vbObjectError + 3988, , "図が空です。"
 M6Draw = "PASS"
 Exit Function
Failed: M6Draw = "ERROR|" & Err.Description
End Function
'@)
 function Draw([string]$mode,[int]$resultId){$response=[string]$xl.Run("'"+$wb.Name+"'!M6Draw",$mode,$resultId);if($response.StartsWith("ERROR|")){throw $response};$response}
 foreach($resultId in 3,69,70,71,72){Assert ((Draw 'A' $resultId) -eq 'PASS') "saved mixed SRM result $resultId with shear/cross A"}
 Assert ((Draw 'B' 3) -eq 'PASS') 'saved mixed SRM increment color B with independent overlays'
 Assert ($wb.Worksheets.Item('図').Shapes.Count -gt 0) 'mixed SRM native shapes rendered'
 $wb.Save();$wb.Close($false);$xl.EnableEvents=$true;$wb=$xl.Workbooks.Open($srm,0,$false);$xl.EnableEvents=$false
 Assert ((Draw 'A' 3) -eq 'PASS') 'increment overlays render after normal reopen'
 $wet=$xl.Workbooks.Open($seep,0,$false)
 $null=$xl.Run("'"+$wet.Name+"'!FEMWaterResults")
 $water=$wet.Worksheets.Item('水圧結果')
 Assert ($water.Shapes.Item('GW_VIEW_FLUX_1').Type -gt 0) 'saved mixed seepage has Darcy arrow'
 Assert ($water.Shapes.Item('GW_VIEW_HEAD_CONTOUR_1').Type -gt 0) 'saved mixed seepage has computed head contour'
 Assert ([string]$water.Range('A4').Value2 -like '浸透計算済*') 'water caption distinguishes computed hydraulic field'
 $water.Activate();$xl.Goto($water.Range('A1'),$true);$xl.ActiveWindow.Zoom=85
 @{status='PASS';checks=@($checks);srm_sha256=(Get-FileHash $SrmBook).Hash;seepage_sha256=(Get-FileHash $SeepageBook).Hash}|ConvertTo-Json -Depth 7|Set-Content (Join-Path $Out 'checks.json') -Encoding utf8
 if($Visible){
  Set-Content (Join-Path $Out 'ready.txt') ready
  while(-not(Test-Path (Join-Path $Out 'close.txt'))){
   $command=Join-Path $Out 'command.txt'
   if(Test-Path $command){$mode=(Get-Content $command -Raw).Trim();Remove-Item -LiteralPath $command
    if($mode -in 'A','B'){$wb.Activate();Assert ((Draw $mode 3) -eq 'PASS') "visual $mode"}
    if($mode -eq 'water'){$wet.Activate();$water.Activate();$xl.Goto($water.Range('A1'),$true)}
   }
   Start-Sleep -Milliseconds 300
  }
 }
 $wet.Close($false);$wet=$null;$wb.Close($false);$wb=$null
}catch{$_.ScriptStackTrace|Write-Output;throw}finally{
 $xl.EnableEvents=$false
 if($wet){try{$wet.Close($false)}catch{}}
 if($wb){try{$wb.Close($false)}catch{}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
'ALL PASS'
