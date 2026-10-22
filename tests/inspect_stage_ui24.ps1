param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
if(Test-Path $Out){throw 'Preserve earlier evidence.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory $Out | Out-Null
$copy=Join-Path $Out 'UI24_screen_check.xlsm';Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$xl=New-Object -ComObject Excel.Application
try{
 $xl.EnableEvents=$true;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false);$xl.Visible=$true;$xl.WindowState=-4137
 $mod=$wb.VBProject.VBComponents.Add(1);$mod.Name='Screen24Driver'
 $mod.CodeModule.AddFromString(@'
Public Function Screen24Error() As String
 Screen24Error = FEM24Error & FEMViewerLastError
End Function
Public Function Screen24Quiet() As String
 SuppressUserMessages = True: FEM24Error = "": FEMViewerLastError = ""
 Screen24Quiet = "PASS"
End Function
Public Function Screen24Overlay(ByVal mode As String, ByVal shear As Boolean, ByVal cross As Boolean) As String
 Dim w As Worksheet, r As Long
 Set w = ThisWorkbook.Worksheets("結果")
 w.CheckBoxes("UI24Overlay0").Value = IIf(shear, 1, -4146)
 w.CheckBoxes("UI24Overlay1").Value = IIf(cross, 1, -4146)
 FEM24Overlay
 For r = 8 To 40
  If Val(CStr(w.Cells(r, 15).Value2)) = 27 Then w.Cells(r, 4).Value2 = mode: Exit For
 Next r
 Screen24Overlay = "PASS"
End Function
'@)
 $prefix="'"+$wb.Name+"'!"
 function Call([string]$name){$xl.Run($prefix+$name)}
 $previous='';Set-Content (Join-Path $Out 'ready.txt') $wb.FullName
 while($true){
  $command=if(Test-Path (Join-Path $Out 'command.txt')){(Get-Content (Join-Path $Out 'command.txt') -Raw).Trim()}else{'home'}
  if($command -ne $previous){
   $previous=$command;$null=Call 'Screen24Quiet';$timer=[Diagnostics.Stopwatch]::StartNew()
   Write-Output ('START '+$command)
   switch($command){
    'home'{$null=Call 'FEM24Home'}
    'model'{$null=Call 'FEM24Model'}
    'materials'{$null=Call 'FEM24Materials'}
    'water'{$null=Call 'FEM24Water'}
    'profiles'{$null=Call 'FEM24Water';$wb.Worksheets.Item('境界・水条件').Range('C4').Value2='水位線'}
    'stages'{$null=Call 'FEM24Stages'}
    'analysis'{$null=Call 'FEM24Analysis'}
    'save'{$null=Call 'FEM24Analysis';$wb.Worksheets.Item('解析・保存').Range('C4').Value2='保存・再開'}
    'results'{$null=Call 'FEM24Results'}
    'options'{$null=Call 'FEM24Results';$wb.Worksheets.Item('結果').Range('C4').Value2='表示設定'}
    'stage_results'{$null=Call 'FEM24Results';$wb.Worksheets.Item('結果').Range('C4').Value2='工程別数値'}
    'preview'{$null=Call 'FEM24Preview'}
    'generate'{$null=Call 'FEM24Generate'}
    'plot'{$null=Call 'FEM24Plot'}
    'overlay_a'{$null=Call 'FEM24Results';$wb.Worksheets.Item('結果').Range('C4').Value2='表示設定';$null=$xl.Run($prefix+'Screen24Overlay','A',$true,$false);$null=Call 'FEM24Plot'}
    'overlay_b'{$null=Call 'FEM24Results';$wb.Worksheets.Item('結果').Range('C4').Value2='表示設定';$null=$xl.Run($prefix+'Screen24Overlay','B',$true,$false);$null=Call 'FEM24Plot'}
    'cross_only'{$null=Call 'FEM24Results';$wb.Worksheets.Item('結果').Range('C4').Value2='表示設定';$null=$xl.Run($prefix+'Screen24Overlay','B',$false,$true);$null=Call 'FEM24Plot'}
    'overlay_both'{$null=Call 'FEM24Results';$wb.Worksheets.Item('結果').Range('C4').Value2='表示設定';$null=$xl.Run($prefix+'Screen24Overlay','A',$true,$true);$null=Call 'FEM24Plot'}

    'water_preview'{$beforeDirty=$wb.Names.Item('_FEMViewerResultsDirty').RefersTo;$beforeDiagnostic=ConvertTo-Json -InputObject $wb.Worksheets.Item('診断').Range('A1:B200').Value2 -Depth 6 -Compress;$null=Call 'FEM24WaterPreview';if($wb.Names.Item('_FEMViewerResultsDirty').RefersTo -ne $beforeDirty){throw 'Water preview invalidated saved results.'};if((ConvertTo-Json -InputObject $wb.Worksheets.Item('診断').Range('A1:B200').Value2 -Depth 6 -Compress) -ne $beforeDiagnostic){$beforeDiagnostic|Set-Content (Join-Path $Out 'diagnostic_before.json');(ConvertTo-Json -InputObject $wb.Worksheets.Item('診断').Range('A1:B200').Value2 -Depth 6 -Compress)|Set-Content (Join-Path $Out 'diagnostic_after.json');throw 'Water preview overwrote saved diagnostics.'}}
    'water_result'{$null=Call 'FEM24WaterFigure'}
    'detail'{$null=Call 'FEM24Stages';$null=Call 'FEM24Detail'}
    'reopen'{$wb.Save();$wb.Close($false);$wb=$xl.Workbooks.Open($copy,0,$false);$null=Call 'Screen24Quiet';$null=Call 'FEM24Plot'}
    'stale'{
      $null=Call 'FEM24Materials';$wb.Worksheets.Item('材料・接合').Range('B8').Value2=[double]$wb.Worksheets.Item('材料・接合').Range('B8').Value2+1.0
      $null=Call 'FEM24Plot';$expected=Call 'Screen24Error'
      if(-not $expected){throw 'Stale plot was not rejected.'}
      if(@($wb.Worksheets.Item('結果').Shapes | Where-Object Name -eq 'UI24Plot').Count -gt 0){throw 'Stale front plot was not removed.'}
      Set-Content (Join-Path $Out 'stale_error.txt') $expected
    }
    'quit'{break}
    default{throw "Unknown command $command"}
   }
   if($command -eq 'quit'){break}
   $timer.Stop();$err=Call 'Screen24Error'
   if($err -and $command -ne 'stale'){throw "$command : $err"}
   if(-not $xl.EnableEvents -or -not $xl.ScreenUpdating){throw 'Excel events/screen state not restored.'}
   $w=$wb.ActiveSheet;$shapes=@()
   if($command -in @('preview','generate','plot','reopen','overlay_a','overlay_b','overlay_both','cross_only')){
    if($w.Name -ne '結果' -or $w.Shapes.Item('UI24Plot').GroupItems.Count -le 1){throw 'Native front plot is missing.'}
   }
   foreach($shape in $w.Shapes){
    $text='';try{$text=$shape.TextFrame2.TextRange.Text}catch{}
    $shapes+=@{name=$shape.Name;text=$text;left=$shape.Left;top=$shape.Top;width=$shape.Width;height=$shape.Height}
   }
   $xl.ActiveWindow.ScrollRow=1;$xl.ActiveWindow.ScrollColumn=1
   if($command -in @('preview','generate','plot','reopen','overlay_a','overlay_b','overlay_both','cross_only')){$xl.ActiveWindow.Zoom=70}else{$xl.ActiveWindow.Zoom=85}
   @{status='PASS';action=$command;sheet=$w.Name;seconds=$timer.Elapsed.TotalSeconds;shapes=$shapes;source_book_sha256=(Get-FileHash $Book).Hash;visible_sheets=@($wb.Worksheets|Where-Object Visible -eq -1|ForEach-Object Name);dirty=$wb.Names.Item('_FEMViewerResultsDirty').RefersTo;node_count=$wb.Worksheets.Item('節点データ').Cells.Item(1048576,1).End(-4162).Row-1;element_count=$wb.Worksheets.Item('要素データ').Cells.Item(1048576,1).End(-4162).Row-1}|ConvertTo-Json -Depth 6|Set-Content (Join-Path $Out ($command+'.json')) -Encoding utf8
   Set-Content (Join-Path $Out 'status.txt') $command
   Write-Output ('DONE '+$command)
  }
  Start-Sleep -Milliseconds 250
 }
}catch{
 Write-Output $_.ScriptStackTrace
 Write-Output $_.Exception.Message
 throw
}finally{
 try{$xl.EnableEvents=$false;if($wb){$wb.Close($false)};$xl.Quit()}catch{Write-Output ('CLEANUP: '+$_.Exception.Message)}
 [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
