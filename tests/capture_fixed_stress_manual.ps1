param([string]$Book='workbook/ui_20261024/2DSoilFEM_ui_20261024.xlsm',[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
if(Test-Path $Out){throw 'Preserve previous evidence.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory $Out|Out-Null
$copy=Join-Path $Out 'FS_manual_capture.xlsm';Copy-Item $Book $copy
$xl=New-Object -ComObject Excel.Application
try{
 $xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false);$xl.Visible=$true;$xl.WindowState=-4137
 $prefix="'"+$wb.Name+"'!"
 $mod=$wb.VBProject.VBComponents.Add(1);$mod.Name='Manual24CaptureDriver'
 $mod.CodeModule.AddFromString(@'
Public Function M24Quiet() As String
 SuppressUserMessages = True
 FEM24Error = "": FEM24StageError = "": FEMViewerLastError = ""
 M24Quiet = "PASS"
End Function
Public Function M24Error() As String
 M24Error = FEM24Error & FEM24StageError & FEMViewerLastError
End Function
'@)
 $null=$xl.Run($prefix+'M24Quiet')
 $null=$xl.Run($prefix+'FEM24Home')
 try{$null=$wb.Worksheets.Item('時間履歴')}catch{$timeSheet=$wb.Worksheets.Add();$timeSheet.Name='時間履歴'}
 $tables=@('材料データ','形状入力','節点データ','要素データ','接合','載荷','ステージ','設定','メッシュ設定','解析設定','高速化設定','出力・表示','詳細設定','水条件','結果節点','結果要素','結果増分','結果塑性','ステージ結果','診断','接合結果','水圧結果','時間履歴')
 $baseline=@{}
 foreach($name in $tables){try{$ws=$wb.Worksheets.Item($name);$baseline[$name]=@{address=$ws.UsedRange.Address();values=$ws.UsedRange.Value2}}catch{}}
 function Put($area,$value){$null=$area.GetType().InvokeMember('Value2',[Reflection.BindingFlags]::SetProperty,$null,$area,[object[]](,$value))}
 function ImportCase([string]$path){
  $xl.EnableEvents=$false;$xl.ScreenUpdating=$false
  foreach($name in $baseline.Keys){$ws=$wb.Worksheets.Item($name);$ws.Unprotect();$ws.UsedRange.ClearContents();Put $ws.Range($baseline[$name].address) $baseline[$name].values}
  $source=$xl.Workbooks.Open((Resolve-Path $path).Path,0,$true)
  try{
   foreach($name in $tables){
    $src=$null;$dst=$null
    try{$src=$source.Worksheets.Item($name);$dst=$wb.Worksheets.Item($name)}catch{continue}
    $dst.Unprotect();$dst.UsedRange.ClearContents();Put $dst.Range($src.UsedRange.Address()) $src.UsedRange.Value2
   }
  }finally{$source.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($source)}
  $null=$xl.Run($prefix+'FEMUiBindSettings')
  $wb.Names.Item('_FEMViewerResultsDirty').RefersTo='=FALSE'
  $null=$xl.Run($prefix+'FEMInvalidateInputCache');$null=$xl.Run($prefix+'FEMInvalidateSettingCache')
  $xl.ScreenUpdating=$true
 }
 $previous='';$currentCase=''
 Set-Content (Join-Path $Out 'ready.txt') 'READY'
 while($true){
  $file=Join-Path $Out 'command.json'
  if(Test-Path $file){
   $raw=Get-Content $file -Raw -Encoding utf8
   if($raw -ne $previous){
    $previous=$raw;$q=$raw|ConvertFrom-Json
    if($q.action -eq 'quit'){break}
    Write-Output ('START '+$q.name)
    if($q.case -and $q.case -ne $currentCase){ImportCase $q.case;$currentCase=$q.case}
    $null=$xl.Run($prefix+'M24Quiet')
    $xl.EnableEvents=$false
    if($q.overrides){foreach($edit in $q.overrides){$w=$wb.Worksheets.Item($edit.sheet);$w.Unprotect();if($edit.clear){$w.Range($edit.cell).ClearContents()}else{Put $w.Range($edit.cell) $edit.value};if($edit.sheet -in @('出力・表示','解析設定','高速化設定','メッシュ設定','詳細設定')){$null=$xl.Run($prefix+'FEMUiSettingsChanged',$w,$w.Range($edit.cell))}};$null=$xl.Run($prefix+'FEMUiCalculateSettings')}
    $null=$xl.Run($prefix+([string]$q.macro))
    $w=$wb.Worksheets.Item([string]$q.page)
    $w.Activate()
    if($q.section){Put $w.Range('C4') ([string]$q.section);$null=$xl.Run($prefix+'FEM24Refresh',$w)}
    $errorText=[string]$xl.Run($prefix+'M24Error');if($errorText){throw $errorText}
    $xl.WindowState=-4137;$xl.ActiveWindow.Zoom=90;$xl.ActiveWindow.ScrollRow=1;$xl.ActiveWindow.ScrollColumn=1
    $xl.EnableEvents=$true;$xl.ScreenUpdating=$true
    $record=@{status='PASS';name=$q.name;page=$w.Name;section=$q.section;source_case=$currentCase;source_sha256=if($currentCase){(Get-FileHash $currentCase).Hash}else{''};book_sha256=(Get-FileHash $Book).Hash;cells=@{}}
    if($q.ranges){foreach($a in $q.ranges){$record.cells[$a]=$w.Range($a).Value2}}
    $record|ConvertTo-Json -Depth 9|Set-Content (Join-Path $Out ($q.name+'.json')) -Encoding utf8
    Set-Content (Join-Path $Out 'status.txt') $q.name
    Write-Output ('DONE '+$q.name)
   }
  }
  Start-Sleep -Milliseconds 200
 }
}finally{
 $xl.EnableEvents=$false
 if($wb){$wb.Close($false)}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
