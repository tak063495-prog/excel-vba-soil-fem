param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel running; no test started.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path $Out){throw 'Preserve previous evidence.'}
New-Item -ItemType Directory $Out | Out-Null
$copy=Join-Path $Out 'water_display.xlsm';Copy-Item $Book $copy
$xl=New-Object -ComObject Excel.Application
try{
 $xl.EnableEvents=$true;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false);$xl.EnableEvents=$false;$xl.Visible=$true;$xl.WindowState=-4137
 $mod=$wb.VBProject.VBComponents.Add(1);$mod.Name='WaterViewUiDriver'
 $mod.CodeModule.AddFromString("Public Function WVLastError() As String`r`nWVLastError = FEMViewerLastError`r`nEnd Function`r`nPublic Function WVResults() As String`r`nOn Error GoTo Failed`r`nSuppressUserMessages = True`r`nFEMWaterResults`r`nWVResults = ""PASS""`r`nExit Function`r`nFailed:`r`nWVResults = ""ERROR|"" & Err.Description`r`nEnd Function")
 $prefix="'"+$wb.Name+"'!";$xl.ActiveWindow.Zoom=85
 $previous='';Set-Content (Join-Path $Out 'ready.txt') $wb.FullName
 while($true){
  $command=if(Test-Path (Join-Path $Out 'command.txt')){(Get-Content (Join-Path $Out 'command.txt') -Raw).Trim()}else{'input'}
  if($command -ne $previous){
   $previous=$command;$status='PASS'
   switch($command){
    'input'{$wb.Worksheets.Item('水条件').Activate();$xl.ActiveWindow.ScrollColumn=1;$xl.ActiveWindow.ScrollRow=1}
    'plans'{$wb.Worksheets.Item('水条件').Activate();$xl.ActiveWindow.ScrollColumn=12;$xl.ActiveWindow.ScrollRow=1}
    'stages'{$wb.Worksheets.Item('ステージ').Activate();$xl.ActiveWindow.ScrollColumn=1;$xl.ActiveWindow.ScrollRow=1}
    'boundary'{$wb.Worksheets.Item('水条件').Activate();$xl.ActiveWindow.ScrollColumn=1;$xl.ActiveWindow.ScrollRow=210}
    'preview'{$xl.Run($prefix+'SetP0SilentMode',$true)|Out-Null;$xl.Run($prefix+'FEMWaterPreview')|Out-Null;$xl.ActiveWindow.ScrollColumn=1;$xl.ActiveWindow.ScrollRow=1}
    'result'{$xl.Run($prefix+'SetP0SilentMode',$true)|Out-Null;$xl.Run($prefix+'FEMWaterResults')|Out-Null;$xl.ActiveWindow.ScrollColumn=1;$xl.ActiveWindow.ScrollRow=1}
    'table'{$wb.Worksheets.Item('水圧結果').Activate();$xl.ActiveWindow.ScrollColumn=1;$xl.ActiveWindow.ScrollRow=32}
    'heads'{$wb.Worksheets.Item('水圧結果').Activate();$xl.ActiveWindow.ScrollColumn=32;$xl.ActiveWindow.ScrollRow=32}
    'mesh'{$xl.Run($prefix+'FEMUiPreviewMesh')|Out-Null;if($xl.Run($prefix+'FEMReadTextSetting','MESH_GEN_LAST_METHOD','') -eq 'ERROR'){throw 'Mesh preview failed'}}
    'stress'{$w=$wb.Worksheets.Item('出力・表示');$w.Unprotect();for($r=8;$r -le $w.UsedRange.Rows.Count;$r++){switch([string]$w.Cells.Item($r,6).Value2){'VIEW_COLOR_RESULT'{$w.Cells.Item($r,3).Value2='Y方向応力（積分点1）'};'VIEW_DISP_SCALE'{$w.Cells.Item($r,3).Value2=1.}}};$xl.Run($prefix+'FEMInvalidateSettingCache')|Out-Null;$xl.Run($prefix+'FEMViewResult')|Out-Null;$err=$xl.Run($prefix+'WVLastError');if($err){throw $err}}
    'reopen'{$wb.Save();$wb.Close($false);$wb=$xl.Workbooks.Open($copy,0,$false);$xl.Run($prefix+'FEMWaterResults')|Out-Null}
    'stale'{$wb.Worksheets.Item('水条件').Range('N12').Value2='CHANGED_PROFILE';$expected=$xl.Run($prefix+'WVResults');if($expected -notmatch '^ERROR\|入力が変更'){throw 'Changed stage-water input was not rejected.'};Set-Content (Join-Path $Out 'stale_error.txt') $expected}
    'quit'{break}
    default{throw 'Unknown UI test command'}
   }
   if($command -eq 'quit'){break}
   $w=$wb.ActiveSheet
   if($command -in @('preview','result','reopen')){$xl.ActiveWindow.Zoom=70}else{$xl.ActiveWindow.Zoom=85}
   $shapes=@()
   foreach($sh in $w.Shapes){
    $text='';try{$text=$sh.TextFrame2.TextRange.Text}catch{}
    $shapes+=@{name=$sh.Name;text=$text;left=$sh.Left;top=$sh.Top;width=$sh.Width;height=$sh.Height}
   }
   @{status=$status;action=$command;sheet=$w.Name;shapes=$shapes;source_book_sha256=(Get-FileHash $Book).Hash}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $Out ($command+'.json')) -Encoding utf8
   Set-Content (Join-Path $Out 'status.txt') $command
  }
  Start-Sleep -Milliseconds 300
 }
}finally{
 if($wb){$wb.Close($false)};$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
