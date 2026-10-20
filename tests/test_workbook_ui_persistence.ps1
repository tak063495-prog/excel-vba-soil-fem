param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Baseline,[string]$Out='tests/results/ui_persistence.json',[switch]$CaptureBaselineOnly)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
Push-Location $repo
try{
 if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
 $Book=(Resolve-Path -LiteralPath $Book).Path;$Baseline=(Resolve-Path -LiteralPath $Baseline).Path
 $sourceHash=(Get-FileHash -LiteralPath $Book).Hash
 $tmp=Join-Path $repo 'tests/tmp/ui_persistence';New-Item -ItemType Directory -Force -Path $tmp|Out-Null
 $copy=Join-Path $tmp 'ui_persistence.xlsm';Copy-Item -LiteralPath $Book -Destination $copy -Force
 $checks=[Collections.Generic.List[string]]::new()
 function Check($ok,[string]$name){if(-not $ok){throw "FAIL $name"};$checks.Add($name)}
 function Snapshot($book){
  $s=[ordered]@{}
  foreach($entry in @(@('材料データ',13),@('節点データ',9),@('要素データ',10),@('ステージ',6),@('載荷',10),@('接合',5),@('形状入力',10),@('図形定義',14))){
   $ws=$book.Worksheets.Item($entry[0]);$last=$ws.Cells($ws.Rows.Count,1).End(-4162).Row
   if($entry[0] -eq '形状入力'){$last=$ws.Cells($ws.Rows.Count,2).End(-4162).Row}
   $s[$entry[0]]=($ws.Range($ws.Cells(2,1),$ws.Cells([Math]::Max(2,$last),[int]$entry[1])).Value2 | ConvertTo-Json -Depth 5 -Compress)
  }
  return ($s|ConvertTo-Json -Depth 6 -Compress)
 }
 function Settings($book){
  $dict=@{}
  foreach($page in 'メッシュ設定','解析設定','高速化設定','出力・表示','詳細設定'){
   $ws=$book.Worksheets.Item($page);$last=$ws.Cells($ws.Rows.Count,6).End(-4162).Row
   for($r=1;$r -le $last;$r++){
    $k=[string]$ws.Cells($r,6).Value2
    if($k){$dict[$k]=@{page=$page;row=$r;value=$ws.Cells($r,3).Value2}}
   }
  }
  $legacy=$book.Worksheets.Item('設定');$last=$legacy.Cells($legacy.Rows.Count,5).End(-4162).Row
  $legacyData=$legacy.Range("C1:E$last").Value2
  for($r=5;$r -le $last;$r++){
   $k=[string]$legacyData.GetValue($r,3)
   if($k -and -not $dict.ContainsKey($k)){$dict[$k]=@{page='設定';row=$r;value=$legacyData.GetValue($r,1)}}
  }
  return $dict
 }
 if($CaptureBaselineOnly){
  $captureApp=$null;$captureBook=$null
  try{
   $captureApp=New-Object -ComObject Excel.Application
   $captureApp.EnableEvents=$false;$captureApp.DisplayAlerts=$false;$captureApp.AutomationSecurity=1
   $captureBook=$captureApp.Workbooks.Open($copy,0,$true)
   $snapshotResult=@{model=(Snapshot $captureBook);settings=(Settings $captureBook)}
   $captureOut=[IO.Path]::GetFullPath($Out)
   New-Item -ItemType Directory -Force -Path (Split-Path $captureOut)|Out-Null
   $snapshotResult|ConvertTo-Json -Depth 9|Set-Content -LiteralPath $captureOut -Encoding utf8
  }finally{
   if($captureBook){$captureBook.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($captureBook)}
   if($captureApp){$captureApp.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($captureApp)}
   [GC]::Collect();[GC]::WaitForPendingFinalizers()
  }
  return
 }
 # Large VBA fixed arrays require one workbook project per Excel process.
 $captureFile=Join-Path $tmp 'baseline_snapshot.json'
 & pwsh -NoProfile -File $PSCommandPath -Book $Baseline -Baseline $Baseline -Out $captureFile -CaptureBaselineOnly
 if($LASTEXITCODE -ne 0){throw 'Baseline snapshot failed.'}
 for($i=0;$i -lt 20;$i++){if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){break};Start-Sleep -Seconds 1}
 if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel still running after baseline snapshot.'}
 $baselineData=Get-Content -LiteralPath $captureFile -Raw -Encoding utf8|ConvertFrom-Json -AsHashtable
 $originalModel=[string]$baselineData.model;$originalSettings=$baselineData.settings
 Copy-Item -LiteralPath $Book -Destination $copy -Force
 $app=$null;$wb=$null;$oldBook=$null
 try{
  $app=New-Object -ComObject Excel.Application
  $app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
  $wb=$app.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
  $h=$wb.VBProject.VBComponents.Add(1);$h.Name='UiPersistenceHarness'
  $h.CodeModule.AddFromString(@'
Public Function UBind() As String
 On Error GoTo Bad
 Call FEMUiBindSettings
 UBind="PASS":Exit Function
Bad:
 UBind="ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function UAnalysis() As String
 On Error GoTo Bad
 Call P0_RunAnalysis
 UAnalysis=ResultStatus & "|" & AnalysisMessage:Exit Function
Bad:
 UAnalysis="ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Sub UArm()
 AnalysisRunning=False:AnalysisOK=True:ResultRevision=123:P1ResultReady=True
End Sub
Public Function UState() As String
 UState=CStr(AnalysisOK) & "|" & CStr(ResultRevision) & "|" & CStr(P1ResultReady)
End Function
Public Function UErrorCount() As Long
 Dim w As Worksheet,r As Long
 Set w=ThisWorkbook.Worksheets("設定")
 For r=5 To w.Cells(w.Rows.Count,5).End(xlUp).Row
  If IsError(w.Cells(r,3).Value2) Then UErrorCount=UErrorCount+1
 Next r
End Function
'@)
  Check ($app.Run($prefix+'FEMUiRefreshInputs') -ceq 'PASS') 'refresh succeeds'
  Check ((Snapshot $wb) -ceq $originalModel) 'all model tables preserved against original workbook'
  $settings=Settings $wb
  foreach($k in $originalSettings.Keys){
   Check ($settings.ContainsKey($k)) "existing canonical key preserved $k"
   Check (($settings[$k].value|ConvertTo-Json -Compress) -ceq ($originalSettings[$k].value|ConvertTo-Json -Compress)) "existing canonical value/type preserved $k"
  }
  $app.Calculation=-4135
  Check ($app.Run($prefix+'P6ReadSetting','NATIVE_STATE_COPY',-999.0) -eq 1) 'prime cached setting'
  $app.Run($prefix+'FEMWriteNumericSetting','native_state_copy',0.0)
  $settings=Settings $wb
  Check ($settings['NATIVE_STATE_COPY'].value -eq 0) 'VBA writer updates canonical input with events disabled'
  Check ($app.Run($prefix+'P6ReadSetting','NATIVE_STATE_COPY',-999.0) -eq 0) 'manual calculation reader sees writer update'
  $app.Run($prefix+'FEMWriteTextSetting','EXPORT_LOAD','AUTO')
  Check ($app.Run($prefix+'P6ReadTextSetting','EXPORT_LOAD','MISSING') -ceq 'AUTO') 'text writer and manual reader agree'
  $app.Run($prefix+'FEMWriteTextSetting','EXPORT_LOAD','')
  $beforeProfile=Settings $wb
  $app.Run($prefix+'UArm')
  $app.Run($prefix+'FEMUiApplyRecommendedRuntime')
  Check ($app.Run($prefix+'UState') -ceq 'False|0|False') 'profile action invalidates prior result state'
  $afterProfile=Settings $wb
  $profileKeys=@('NATIVE_DLL','NATIVE_PRECISION','NATIVE_STATE_COPY','NATIVE_ASSEMBLY','NATIVE_GMRES_SIMD','NATIVE_GMRES_TUNED','NATIVE_GMRES_PREPARED','IO_FAST_PATH','EVENT_LOG_FAST')
  foreach($k in $beforeProfile.Keys){if($k -notin $profileKeys){Check (($beforeProfile[$k].value|ConvertTo-Json -Compress) -ceq ($afterProfile[$k].value|ConvertTo-Json -Compress)) "profile preserves other setting $k"}}
  Check ((Snapshot $wb) -ceq $originalModel) 'profile action preserves all physical model tables'
  foreach($k in $profileKeys){$expected=if($k -in @('NATIVE_DLL','NATIVE_PRECISION')){'AUTO'}else{1};Check ([string]$afterProfile[$k].value -ceq [string]$expected) "profile restores $k"}
  $app.EnableEvents=$true
  $m=$afterProfile['FLOW_POLICY'];$cell=$wb.Worksheets.Item($m.page).Cells($m.row,3)
  $app.Run($prefix+'UArm');$cell.Value2='DAVIS'
  Check ($app.Run($prefix+'UState') -ceq 'False|0|False') 'analysis setting edit invalidates results'
  $cell.Value2=[string]$m.value
  $m=$afterProfile['VIEW_DISP_SCALE'];$cell=$wb.Worksheets.Item($m.page).Cells($m.row,3)
  $app.Run($prefix+'UArm');$cell.Value2=2.0
  Check ($app.Run($prefix+'UState') -ceq 'True|123|True') 'display-only edit retains numerical result state'
  $cell.Value2=[double]$m.value
  $app.EnableEvents=$false
  $m=$afterProfile['NATIVE_STATE_COPY'];$ws=$wb.Worksheets.Item($m.page);$ws.Unprotect()
  $registry=$wb.Worksheets.Item('設定');$last=$registry.Cells($registry.Rows.Count,5).End(-4162).Row
  $formulaBefore=($registry.Range("C5:C$last").Formula|ConvertTo-Json -Compress -Depth 3)
  $ws.Cells($m.row,6).Value2='NATIVE_DLL'
  Check ($app.Run($prefix+'UBind') -like 'ERROR|*') 'duplicate UI mapping rejected'
  Check (($registry.Range("C5:C$last").Formula|ConvertTo-Json -Compress -Depth 3) -ceq $formulaBefore) 'rejected duplicate mapping causes no partial rebind'
  $ws.Cells($m.row,6).Value2='NATIVE_STATE_COPY'
  Check ($app.Run($prefix+'FEMUiRefreshInputs') -ceq 'PASS') 'layout refresh after mapping repair'
  # Preserve a customized canonical location rather than duplicating its key.
  $moving=(Settings $wb)['NATIVE_STATE_COPY'];$from=$wb.Worksheets.Item($moving.page);$from.Unprotect()
  $custom=$wb.Worksheets.Item('詳細設定');$custom.Unprotect()
  $customRow=[Math]::Max(37,$custom.Cells($custom.Rows.Count,6).End(-4162).Row+1)
  $custom.Cells($customRow,2).Value2='既存カスタムDLL設定'
  $custom.Cells($customRow,3).Value2=0.0;$custom.Cells($customRow,6).Value2='NATIVE_STATE_COPY'
  $from.Cells($moving.row,6).ClearContents()
  Check ($app.Run($prefix+'FEMUiRefreshInputs') -ceq 'PASS') 'custom runtime key on another page migrates without error'
  $customSettings=Settings $wb
  Check ($customSettings['NATIVE_STATE_COPY'].page -ceq '詳細設定' -and $customSettings['NATIVE_STATE_COPY'].value -eq 0) 'custom canonical page and value preserved'
  Check ($app.Run($prefix+'P6ReadSetting','NATIVE_STATE_COPY',-999.0) -eq 0) 'reader follows customized canonical page'
  $app.Run($prefix+'FEMUiApplyRecommendedRuntime')
  Check ($app.Run($prefix+'P6ReadSetting','NATIVE_STATE_COPY',-999.0) -eq 1) 'profile writer follows customized canonical page'
  Check ($app.Run($prefix+'FEMUiRefreshInputs') -ceq 'PASS') 'repeated layout refresh idempotent'
  $saved=Settings $wb
  foreach($sheet in $wb.Worksheets){foreach($shape in $sheet.Shapes){$action='';try{$action=[string]$shape.OnAction}catch{};if($action){Check ($action.Replace("'",'').StartsWith($wb.Name+"!")) ('button targets this workbook '+$shape.Name)}}}
  foreach($page in '操作パネル','操作ガイド','メッシュ設定','解析設定','高速化設定','出力・表示','詳細設定'){
   foreach($link in $wb.Worksheets.Item($page).Hyperlinks){Check (-not ([string]$link.SubAddress).Contains('[')) "internal navigation $page"}
  }
  Check ($app.Run($prefix+'UErrorCount') -eq 0) 'all registry formulas evaluate without errors'
  $wb.Save();$wb.Close($false)
  $app.EnableEvents=$true
  $wb=$app.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
  Check ($wb.ActiveSheet.Name -ceq '操作パネル') 'Workbook_Open shows operation panel'
  $reopened=Settings $wb
  foreach($k in $saved.Keys){Check (($saved[$k].value|ConvertTo-Json -Compress) -ceq ($reopened[$k].value|ConvertTo-Json -Compress)) "save/open preserves $k"}
  Check ((Snapshot $wb) -ceq $originalModel) 'save/open preserves model tables'
  $app.EnableEvents=$false
  $m=$reopened['SRM_TOL'];$cell=$wb.Worksheets.Item($m.page).Cells($m.row,3);$cell.Value2=-1.0
  $outputBefore=($wb.Worksheets.Item('結果節点').UsedRange.Value2|ConvertTo-Json -Depth 4 -Compress)
  $run=[string]$app.Run($prefix+'UAnalysis')
  Check ($run.StartsWith('INPUT_ERROR|')) "analysis rejects pasted invalid setting as INPUT_ERROR: $run"
  Check ($cell.Value2 -eq -1) 'invalid input is not silently replaced by default'
  Check (($wb.Worksheets.Item('結果節点').UsedRange.Value2|ConvertTo-Json -Depth 4 -Compress) -ceq $outputBefore) 'invalid setting rejected before erasing numerical output'
  Check ((Snapshot $wb) -ceq $originalModel) 'invalid analysis attempt preserves model tables'
  $wb.Close($false)
 }finally{
  if($oldBook){try{$oldBook.Close($false)}catch{}}
  if($wb){try{$wb.Close($false)}catch{}}
  if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
 }
 Check ((Get-FileHash -LiteralPath $Book).Hash -ceq $sourceHash) 'tested deliverable unchanged'
 $outPath=[IO.Path]::GetFullPath($Out);New-Item -ItemType Directory -Force -Path (Split-Path $outPath)|Out-Null
 @{status='PASS';count=$checks.Count;checks=$checks;book=$Book;baseline=$Baseline;book_sha256=$sourceHash}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $outPath -Encoding utf8
 "PASS persistence/input preservation checks=$($checks.Count)"
}finally{Pop-Location}

