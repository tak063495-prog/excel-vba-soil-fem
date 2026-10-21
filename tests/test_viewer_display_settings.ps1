param([string]$Book='workbook/ui_20261020/2DSoilFEM_ui_20261020.xlsm',[string]$Out='tests/results/viewer_display_settings.json',[switch]$ExpectDirtyBug,[switch]$UseSource,[switch]$PreviewStart)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
Push-Location $repo
$app=$null;$wb=$null
try{
 if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no test instance started.'}
 $Book=(Resolve-Path -LiteralPath $Book).Path;$hash=(Get-FileHash -LiteralPath $Book).Hash
 $temp=Join-Path $repo ('tests/tmp/viewer_'+[Guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $temp|Out-Null
 $copy=Join-Path $temp 'viewer.xlsm';Copy-Item -LiteralPath $Book -Destination $copy
 $app=New-Object -ComObject Excel.Application
 $app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $wb=$app.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
 if($UseSource){foreach($m in 'FEMUi','FEMViewer','FEMNativeFigure','ThisWorkbook'){$ext=if($m -eq 'ThisWorkbook'){'.cls'}else{'.bas'};$component=$null;try{$component=$wb.VBProject.VBComponents.Item($m)}catch{};if(-not $component){$component=$wb.VBProject.VBComponents.Add(1);$component.Name=$m};$cm=$component.CodeModule;if($cm.CountOfLines){$cm.DeleteLines(1,$cm.CountOfLines)};$cm.AddFromString((Get-Content -LiteralPath (Join-Path $repo ("src/vba/"+$m+$ext)) -Raw -Encoding UTF8))}}
 $h=$wb.VBProject.VBComponents.Add(1);$h.Name='ViewerDisplayTests'
 $h.CodeModule.AddFromString(@'
Option Explicit
Public Sub VSeed()
 Dim ws As Worksheet, p As Variant, r As Long,j As Long, xy As Variant
 SuppressUserMessages=True:AnalysisRunning=False
 For Each p In Array("節点データ","要素データ","材料データ","ステージ","載荷","結果節点","結果要素","ステージ結果")
  Set ws=ThisWorkbook.Worksheets(CStr(p)):ws.Unprotect
  r=ws.Cells(ws.Rows.Count,1).End(xlUp).Row
  If r>=2 Then ws.Rows("2:" & CStr(r)).ClearContents
 Next p
 xy=Array(Array(0#,0#),Array(1#,0#),Array(1#,1#),Array(0#,1#),Array(0.5,0#),Array(1#,0.5),Array(0.5,1#),Array(0#,0.5))
 For r=0 To 7
  With ThisWorkbook.Worksheets("節点データ")
   .Cells(r+2,1).Value2=r+1:.Cells(r+2,2).Value2=xy(r)(0):.Cells(r+2,3).Value2=xy(r)(1)
   .Cells(r+2,4).Value2=IIf(xy(r)(1)=0#,1,0):.Cells(r+2,5).Value2=IIf(xy(r)(1)=0#,1,0)
  End With
  With ThisWorkbook.Worksheets("結果節点")
   .Cells(r+2,1).Value2=1:.Cells(r+2,2).Value2=r+1:.Cells(r+2,3).Value2=xy(r)(0):.Cells(r+2,4).Value2=xy(r)(1)
   .Cells(r+2,5).Value2=0.001*xy(r)(1):.Cells(r+2,6).Value2=-0.002*xy(r)(1)
  End With
 Next r
 With ThisWorkbook.Worksheets("要素データ")
  .Cells(2,1).Value2=1:For j=1 To 8:.Cells(2,j+1).Value2=j:Next j:.Cells(2,10).Value2=1
 End With
 With ThisWorkbook.Worksheets("材料データ")
  .Cells(2,1).Value2=1:.Cells(2,2).Value2=14000:.Cells(2,3).Value2=0.3:.Cells(2,4).Value2=1:.Cells(2,5).Value2=20
  .Cells(2,6).Value2=30:.Cells(2,7).Value2=100:.Cells(2,8).Value2=0:.Cells(2,9).Value2=1:.Cells(2,10).Value2="ON":.Cells(2,11).Value2="SOIL"
 End With
 With ThisWorkbook.Worksheets("ステージ")
  .Cells(2,1).Value2=1:.Cells(2,2).Value2="GRAVITY":.Cells(2,5).Value2=1
 End With
 With ThisWorkbook.Worksheets("結果要素")
  .Cells(2,1).Value2=1:.Cells(2,2).Value2=1:.Cells(2,3).Value2=0.5:.Cells(2,4).Value2=0.5
  For j=5 To 72:.Cells(1,j).Value2="Q" & CStr(j-4):.Cells(2,j).Value2=j-4:Next j
 End With
 FEMWriteTextSetting "VIEW_SCALE_MODE","AUTO"
 FEMWriteNumericSetting "VIEW_DISP_SCALE",1#
 FEMWriteNumericSetting "VIEW_COLOR_RESULT",1#
 VArm
End Sub
Public Sub VArm()
 AnalysisRunning=False:AnalysisOK=True:ResultRevision=123:P1ResultReady=True
 ThisWorkbook.Worksheets("診断").Cells(2,2).Value2="PASS"
 Call FEMViewerMarkClean
End Sub
Public Function VDirty() As Boolean
 VDirty=(ThisWorkbook.Names("_FEMViewerResultsDirty").RefersTo="=TRUE")
End Function
Public Function VState() As String
 VState=CStr(AnalysisOK) & "|" & CStr(ResultRevision) & "|" & CStr(P1ResultReady)
End Function
Public Function VDraw(ByVal mode As String) As String
 On Error GoTo Bad
 SuppressUserMessages=True
 Select Case mode
 Case "RESULT":Call FEMViewResult
 Case "INITIAL":Call FEMViewInitial
 Case "ALL":Call FEMViewAll
 End Select
 VDraw=FEMViewerLastError
 Exit Function
Bad:
 VDraw=CStr(Err.Number) & "|" & Err.Description
End Function
Public Function VRun() As String
 On Error GoTo Bad
 Call P0_RunAnalysis
 VRun=ResultStatus & "|" & CStr(AnalysisOK)
 Exit Function
Bad:
 VRun="ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
'@)
 if($PreviewStart){
  $cm=$h.CodeModule;$code=$cm.Lines(1,$cm.CountOfLines)
  $code=$code.Replace(' Call P0_RunAnalysis'," Call FEMUiPreviewMesh`r`n If Len(FEMUiMeshPreviewError) > 0 Then Err.Raise vbObjectError + 7002, , FEMUiMeshPreviewError`r`n Call FEMUiStartFromMesh")
  $cm.DeleteLines(1,$cm.CountOfLines);$cm.AddFromString($code)
 }
 $compile=$app.VBE.CommandBars.FindControl(1,578)
 if($compile.Enabled){$compile.Execute()}
 if($app.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Viewer test VBA compile failed.'}
 $checks=[Collections.Generic.List[string]]::new()
 function Check($ok,[string]$label){if(-not $ok){throw "FAIL $label"};$checks.Add($label);if($checks.Count % 25 -eq 0){Write-Output ("checked "+$checks.Count+" "+$label)}}
 function FindSetting([string]$key){foreach($p in 'メッシュ設定','解析設定','高速化設定','出力・表示','詳細設定'){$ws=$wb.Worksheets.Item($p);for($r=8;$r -le $ws.Cells($ws.Rows.Count,6).End(-4162).Row;$r++){if([string]$ws.Cells($r,6).Value2 -ceq $key){return $ws.Cells($r,3)}}};throw "Missing key $key"}
 $app.Run($prefix+'VSeed');$app.Calculation=-4105;$app.EnableEvents=$true
 foreach($mode in 'ALL','INITIAL','RESULT'){Check ([string]$app.Run($prefix+'VDraw',$mode) -ceq '') "initial $mode draws"}
 $cell=FindSetting 'VIEW_DISP_SCALE';$cell.Value2=2.0
 if($ExpectDirtyBug){
  Check ([bool]$app.Run($prefix+'VDirty')) 'reproduced: display edit incorrectly dirties results'
  Check ([string]$app.Run($prefix+'VDraw','RESULT') -like '*再解析*') 'reproduced: result display rejected after display-only edit'
 }else{
  Check (-not [bool]$app.Run($prefix+'VDirty')) 'display edit preserves clean result marker'
  Check ([string]$app.Run($prefix+'VState') -ceq 'True|123|True') 'display edit preserves analysis state'
  $allKeys=@(@('VIEW_DISP_SCALE',3.0),@('VIEW_COLOR_RESULT',13.0),@('VIEW_SCALE_MODE','SYMMETRIC'),@('VIEW_DETAIL_MODE','DETAIL'),@('VIEW_RESULT_SOURCE','RAW_GAUSS'),@('DEBUG_FLUSH',4.0),@('EXPORT_LOAD','AUTO'))
  foreach($entry in $allKeys){$c=FindSetting $entry[0];if($entry[1] -is [string]){$c.Value2=[string]$entry[1]}else{$c.Value2=[double]$entry[1]};Check (-not [bool]$app.Run($prefix+'VDirty')) ($entry[0]+' preserves clean results');Check ([string]$app.Run($prefix+'VDraw','RESULT') -ceq '') ($entry[0]+' permits result display')}
  for($id=1;$id -le 68;$id++){$c=FindSetting 'VIEW_COLOR_RESULT';$c.Value2=[double]$id;Check ([string]$app.Run($prefix+'VDraw','RESULT') -ceq '') "result quantity $id draws"}
  $app.Calculation=-4135;$c=FindSetting 'VIEW_DISP_SCALE';$c.Value2=7.0
  Check ([string]$app.Run($prefix+'VDraw','RESULT') -ceq '') 'manual calculation result display'
  $registry=$wb.Worksheets.Item('設定');$rr=0;for($r=5;$r -le $registry.Cells($registry.Rows.Count,5).End(-4162).Row;$r++){if([string]$registry.Cells($r,5).Value2 -ceq 'VIEW_DISP_SCALE'){$rr=$r;break}}
  Check ([double]$registry.Cells($rr,3).Value2 -eq 7.0) 'viewer recalculates display setting formulas in manual mode'
  $c.Formula='=3+5'
  Check ([string]$app.Run($prefix+'VDraw','RESULT') -ceq '') 'manual display setting formula redraw'
  Check ([double]$registry.Cells($rr,3).Value2 -eq 8.0) 'viewer evaluates canonical display formulas before linked registry'
  $app.Calculation=-4105
  foreach($key in 'FLOW_POLICY','SRM_TOL','NATIVE_DLL'){$app.Run($prefix+'VArm');$c=FindSetting $key;$v=$c.Value2;if($v -is [string]){$c.Value2=[string]$v}else{$c.Value2=[double]$v};Check ([bool]$app.Run($prefix+'VDirty')) "$key still invalidates results";Check ([string]$app.Run($prefix+'VDraw','RESULT') -like '*再解析*') "$key stale result is blocked"}
  $app.Run($prefix+'VArm');$wb.Worksheets.Item('節点データ').Cells(2,2).Value2=0.1
  Check ([bool]$app.Run($prefix+'VDirty')) 'mesh edit invalidates results';Check ([string]$app.Run($prefix+'VDraw','RESULT') -like '*再解析*') 'changed mesh stale result is blocked'
  $app.EnableEvents=$false;$app.Run($prefix+'VSeed');$app.Run($prefix+'FEMWriteTextSetting','EXPORT_LOAD','');$app.Run($prefix+'FEMWriteTextSetting','NATIVE_DLL','OFF')
  Check ([string]$app.Run($prefix+'VRun') -ceq 'PASS|True') 'actual gravity calculation passes without DLL'
  $app.EnableEvents=$true;$c=FindSetting 'VIEW_DISP_SCALE';$c.Value2=10.0
  Check ([string]$app.Run($prefix+'VDraw','RESULT') -ceq '') 'result display after actual calculation and scale edit'
  Check ($wb.Worksheets.Item('図').Shapes.Item('P1Agg_ViewSVG').AlternativeText -like '*要素*') 'actual embedded result picture exists'
  $wb.Save();$wb.Close($false);$app.EnableEvents=$true;$wb=$app.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
  Check (-not [bool]$app.Run($prefix+'VDirty')) 'saved successful result remains clean after Workbook_Open'
  $c=FindSetting 'VIEW_COLOR_RESULT';$c.Value2=2.0
  Check ([string]$app.Run($prefix+'VDraw','RESULT') -ceq '') 'reopened result quantity edit redraws without recalculation'
  $wb.Worksheets.Item('材料データ').Cells(2,2).Value2=15000.0
  Check ([bool]$app.Run($prefix+'VDirty')) 'material edit invalidates results'
  $wb.Save();$wb.Close($false);$wb=$app.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
  Check ([bool]$app.Run($prefix+'VDirty')) 'opening preserves genuinely dirty saved state'
  Check ([string]$app.Run($prefix+'VDraw','RESULT') -like '*再解析*') 'reopened changed material still blocks stale result'
 }
 Check ((Get-FileHash -LiteralPath $Book).Hash -ceq $hash) 'source workbook unchanged'
 $result=@{status='PASS';book=$Book;source_sha256=$hash;expected_dirty_bug=[bool]$ExpectDirtyBug;use_source=[bool]$UseSource;preview_start=[bool]$PreviewStart;checks=$checks.ToArray();count=$checks.Count}
 $result|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $Out -Encoding UTF8
 Write-Output ("PASS "+$checks.Count+" viewer checks")
}finally{
 if($wb){try{$wb.Close($false)}catch{}}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers();Pop-Location
}
