param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out,[switch]$Screen)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no dispatch.'}
if(Test-Path $Out){throw 'Preserve previous evidence.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory $Out|Out-Null
$copy=Join-Path $Out 'ui_test.xlsm';Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$sourceHash=(Get-FileHash $Book).Hash
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1;$xl.Visible=[bool]$Screen
 $wb=$xl.Workbooks.Open($copy,0,$false)
 function Run([string]$name){$xl.Run("'"+$wb.Name+"'!"+$name)}
 function Assert($condition,[string]$message){if(-not $condition){throw $message}}
 $mod=$wb.VBProject.VBComponents.Add(1);$mod.Name='FixedStressUiDriver'
 $mod.CodeModule.AddFromString(@'
Public Function FSUStatus() As String
 FSUStatus = FEM24Error & "|" & FEM24StageError
End Function
Public Sub FSUQuiet()
 SuppressUserMessages = True
End Sub
'@)
 # Use an explicit middle-stage fixture. The distributed book may legitimately
 # have only two stages; uniqueness and DOWN must not depend on its defaults.
 $uiStage=$wb.Worksheets.Item('ステージ');$uiStage.Range('A2:F2000').ClearContents()
 $uiKinds=@('GRAVITY','LOAD','SRM')
 for($r=0;$r -lt 3;$r++){
  $uiStage.Cells($r+2,1).Value2=[double]($r+1);$uiStage.Cells($r+2,2).Value2=$uiKinds[$r]
  $uiStage.Cells($r+2,4).Value2=1.;$uiStage.Cells($r+2,5).Value2=1.
 }
 $null=Run 'FSUQuiet';$null=Run 'FEM24Install'
 $water=$wb.Worksheets.Item('水条件');$stage=$wb.Worksheets.Item('ステージ');$front=$wb.Worksheets.Item('施工工程')
 # Use actual front-sheet change events for coupled material and time detail.
 $mat=$wb.Worksheets.Item('材料・接合');$mat.Activate();$xl.EnableEvents=$true
 $mat.Range('C4').Value2='圧密材料特性'
 $old=$water.Range('X12').Value2
 $mat.Range('C8').Value2=2e-6
 Assert ([double]$water.Range('X12').Value2 -eq 2e-6) 'Material front event did not update storage'
 $mat.Range('C8').Value2=$old
 $xl.EnableEvents=$false
 $water.Range('L12:T211').ClearContents();$water.Range('AC12:AL211').ClearContents()
 $water.Range('B3').Value2='HYDROSTATIC'
 $water.Range('L12').Value2=2.;$water.Range('M12').Value2='FIXED_STRESS'
 $water.Range('O12').Value2=1.;$water.Range('P12').Value2=0.25
 $water.Range('AC12').Value2=2.;$water.Range('AD12').Value2='D';$water.Range('AE12').Value2='HOLD';$water.Range('AF12').Value2='INSTANT'
 $water.Range('AK12').Value2='UI_SET'
 $front.Activate();$xl.EnableEvents=$true;$front.Range('C4').Value2='時間の詳細設定'
 $front.Range('B8').Value2='H'
 Assert ([string]$water.Range('AD12').Value2 -eq 'H') 'Time unit front event failed'
 $front.Range('B8').Value2='D'
 $front.Range('C4').Value2='工程の編集'
 $xl.EnableEvents=$false
 $before=[int]$stage.Cells($stage.Rows.Count,1).End(-4162).Row
 $front.Range('A9').Select();$null=Run 'FEM24StageCopy'
 $uiStatus=Run 'FSUStatus'
 Assert ($uiStatus -eq '|') "Copy returned UI error: $uiStatus"
 $newId=[int]$stage.Cells($before+1,1).Value2
 Assert ($newId -gt 3) 'Copy stage ID not unique'
 $found=$false
 for($r=12;$r -le 211;$r++){
  if($water.Cells($r,29).Value2 -eq $newId){
   Assert ([string]$water.Cells($r,30).Value2 -eq 'D') 'Copied time unit lost'
   Assert ([string]$water.Cells($r,37).Value2 -eq 'UI_SET') 'Copied load set lost';$found=$true
  }
 }
 Assert $found 'Copied stage has no time detail'
 $null=Run 'FEM24StageUndo'
 Assert ([int]$stage.Cells($stage.Rows.Count,1).End(-4162).Row -eq $before) 'Undo did not restore stage count'
 for($r=12;$r -le 211;$r++){Assert ($water.Cells($r,29).Value2 -ne $newId) 'Undo retained copied time detail'}
 $front.Range('A9').Select();$null=Run 'FEM24StageDown'
 Assert ([int]$stage.Range('A4').Value2 -eq 2) 'Move changed the stable ID'
 Assert ([string]$water.Range('AD12').Value2 -eq 'D') 'Move changed time detail'
 $null=Run 'FEM24StageUndo'
 Assert ([int]$stage.Range('A3').Value2 -eq 2) 'Undo move failed'
 $front.Range('A9').Select();$null=Run 'FEM24StageDelete'
 Assert ($null -eq $water.Range('AC12').Value2) 'Delete retained time detail'
 $null=Run 'FEM24StageUndo'
 Assert ([double]$water.Range('AC12').Value2 -eq 2) 'Undo delete lost time detail'
 $xl.EnableEvents=$true;$front.Range('C4').Value2='時間・連成設定'
 Assert ([string]$front.Range('J8').Value2 -eq 'D') 'Displayed time unit differs'
 $xl.EnableEvents=$false
 @{status='PASS';checks=@('material-event','time-unit-event','copy-time-and-loadset','undo-copy','move-stable-ID','undo-move','delete-time','undo-delete','display-unit');excel_version=$xl.Version;excel_build=$xl.Build;book_sha256=(Get-FileHash $Book).Hash}|ConvertTo-Json|Set-Content (Join-Path $Out 'ui.json') -Encoding utf8
 if($Screen){Set-Content (Join-Path $Out 'screen_ready.txt') ready;while(-not(Test-Path (Join-Path $Out 'screen_release.txt'))){Start-Sleep -Milliseconds 300}}
 $wb.Close($false)
}catch{
 Write-Output $_.ScriptStackTrace
 throw
}finally{
 if($wb){try{$wb.Close($false)}catch{}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}

if((Get-FileHash $Book).Hash -ne $sourceHash){throw 'Source evidence workbook changed'}
