param([string]$Book='workbook/ui_20261020/2DSoilFEM_ui_20261020.xlsm',[string]$Out='tests/results/shape_mesh_workflow.json',[switch]$CompileOnly)
$ErrorActionPreference='Stop'; $repo=Split-Path -Parent $PSScriptRoot; Push-Location $repo
$app=$null;$wb=$null
try {
  if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no test instance started.'}
  $src=(Resolve-Path -LiteralPath $Book).Path; $hash=(Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash
  $tmp=Join-Path $repo ('tests/tmp/shape_mesh_'+[Guid]::NewGuid().ToString('N'));New-Item -ItemType Directory -Path $tmp|Out-Null
  $copy=Join-Path $tmp 'shape_mesh.xlsm';Copy-Item -LiteralPath $src -Destination $copy
  $app=New-Object -ComObject Excel.Application;$app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
  $wb=$app.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
  $app.Run($prefix+'FEMWriteBuildStamp',$false)
  $h=$wb.VBProject.VBComponents.Add(1);$h.Name='ShapeMeshWorkflowTests'
  $code=@'
Option Explicit
Public Function ShapeFixture() As String
 On Error GoTo Bad
 SuppressUserMessages=True
 Dim ws As Worksheet, pts As Variant, i As Long
 Set ws=ThisWorkbook.Worksheets("形状入力"):ws.Unprotect
 ws.Range("B8:D2000").ClearContents:ws.Range("G8:J2000").ClearContents
 pts=Array(Array(1#,0#,0#),Array(2#,1#,0#),Array(3#,1#,1#),Array(4#,0#,1#))
 For i=0 To 3:ws.Cells(i+8,2).Value2=pts(i)(0):ws.Cells(i+8,3).Value2=pts(i)(1):ws.Cells(i+8,4).Value2=pts(i)(2):Next i
 ws.Cells(8,7).Value2=1#:ws.Cells(8,8).Value2="外側":ws.Cells(8,9).Value2="1,2,3,4":ws.Cells(8,10).Value2=0.5
 ThisWorkbook.Worksheets("図形定義").Range("A1:N2000").ClearContents
 FEMUiGenerateShape
 If Not FEMShapeGenerationOK Or Len(FEMUiMeshPreviewError) > 0 Or ActiveSheet.Name <> "図" Then
   ShapeFixture = "ERROR|" & FEMUiMeshPreviewError
   Exit Function
 End If
 ShapeFixture="PASS":Exit Function
Bad:ShapeFixture="ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function ShapeCounts() As String
 ShapeCounts=CStr(ThisWorkbook.Worksheets("節点データ").Cells(Rows.Count,1).End(xlUp).Row-1) & "|" & CStr(ThisWorkbook.Worksheets("要素データ").Cells(Rows.Count,1).End(xlUp).Row-1)
End Function
Public Function ShapeQuality() As String
 ShapeQuality = FEMUiMeshQualityNote
 If InStr(1, ShapeQuality, P5QualitySummary(), vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 7001, , "Preview must use evaluated Mesh quality, not analysis globals"
End Function
Public Function ShapePreview() As String
 On Error GoTo Bad
 SuppressUserMessages=True
 AnalysisOK=False:ResultRevision=777
 FEMUiRun
 ShapePreview=CStr(AnalysisOK) & "|" & CStr(ResultRevision) & "|" & FEMUiMeshPreviewError:Exit Function
Bad:ShapePreview="ERROR|" & Err.Description
End Function
Public Function StartReview() As String
 On Error GoTo Bad
 SuppressUserMessages=True
 ThisWorkbook.Worksheets("節点データ").Cells(2,2).Value2=0.25
 AnalysisOK=False:ResultRevision=777
 FEMUiStartFromMesh
 StartReview=CStr(AnalysisOK) & "|" & CStr(ResultRevision) & "|" & FEMUiMeshPreviewError:Exit Function
Bad:StartReview="ERROR|" & Err.Description
End Function
Public Function InvalidateShape() As String
 On Error GoTo Bad
 SuppressUserMessages=True
 ThisWorkbook.Worksheets("形状入力").Cells(8,3).Value2=0.25
 FEMUiStartFromMesh:InvalidateShape="PASS":Exit Function
Bad:InvalidateShape="ERROR|" & Err.Description
End Function
Public Function BadShape() As String
 On Error GoTo Bad
 SuppressUserMessages=True
 ThisWorkbook.Worksheets("形状入力").Cells(8,9).Value2="1,2,9,4"
 FEMUiGenerateShape
 BadShape=CStr(FEMShapeGenerationOK) & "|" & FEMUiMeshPreviewError:Exit Function
Bad:BadShape="ERROR|" & Err.Description
End Function
Public Function CandidateFlow() As String
 On Error GoTo Bad
 Dim n As Long, e As Long, beforeN As Variant, beforeE As Variant, beforeNX As Variant
 SuppressUserMessages=True
 n=ThisWorkbook.Worksheets("節点データ").Cells(Rows.Count,1).End(xlUp).Row
 e=ThisWorkbook.Worksheets("要素データ").Cells(Rows.Count,1).End(xlUp).Row
 beforeN=ThisWorkbook.Worksheets("節点データ").Cells(2,2).Value2:beforeE=ThisWorkbook.Worksheets("要素データ").Cells(2,2).Value2
 If Not MeshPrepareCandidate("SMOOTH") Then CandidateFlow="PREPARE_FAIL|" & MeshCandidateError:Exit Function
 MeshCancelCandidateUI
 If ActiveSheet.Name<>"メッシュ設定" Then CandidateFlow="CANCEL_NOT_MESH":Exit Function
 If ThisWorkbook.Worksheets("節点データ").Cells(Rows.Count,1).End(xlUp).Row<>n Or ThisWorkbook.Worksheets("要素データ").Cells(Rows.Count,1).End(xlUp).Row<>e Then CandidateFlow="CANCEL_MUTATED":Exit Function
 If ThisWorkbook.Worksheets("節点データ").Cells(2,2).Value2<>beforeN Or ThisWorkbook.Worksheets("要素データ").Cells(2,2).Value2<>beforeE Then CandidateFlow="CANCEL_MUTATED":Exit Function
 If Not MeshPrepareCandidate("SMOOTH") Then CandidateFlow="REPREPARE_FAIL|" & MeshCandidateError:Exit Function
 ThisWorkbook.Worksheets("メッシュ設定").Activate
 MeshApplyCandidateUI
 If ActiveSheet.Name <> "図" Or Len(FEMUiMeshPreviewError) > 0 Then
   CandidateFlow = "APPLY_OR_PREVIEW_FAIL|"  &  MeshCandidateError()  &  "|"  &  FEMUiMeshPreviewError
   Exit Function
 End If
 CandidateFlow="PASS|" & CStr(ThisWorkbook.Worksheets("図").Shapes.Count):Exit Function
Bad:CandidateFlow="ERROR|" & Err.Description
End Function
'@
  $h.CodeModule.AddFromString($code)
  # Compile the injected harness as well as the production modules before any test runs.
  $compile=$app.VBE.CommandBars.FindControl(1,578)
  if($compile.Enabled){$compile.Execute()}
  if($app.VBE.CommandBars.FindControl(1,578).Enabled){throw 'VBA compilation did not complete.'}
  Write-Output 'PASS Excel VBA compile (including ShapeMeshWorkflowTests)'
  if($CompileOnly){
    $app.Visible=$true;$app.VBE.MainWindow.Visible=$true;$h.CodeModule.CodePane.Show()
    $record=@{status='PASS';source_sha256=$hash;compile='PASS';harness='ShapeMeshWorkflowTests';compile_command_enabled=$false}
    $record|ConvertTo-Json|Set-Content -LiteralPath $Out -Encoding UTF8
    Write-Output 'READY: compiled test project in VBE, waiting for screen inspection'
    $release=Join-Path $tmp 'screen_checked';Write-Output $release
    for($wait=0;$wait -lt 60;$wait++){if(Test-Path -LiteralPath $release){break};Start-Sleep -Seconds 1}
    return
  }
  $checks=[Collections.Generic.List[string]]::new();function Check($ok,[string]$label){if(-not $ok){throw "FAIL $label"};$checks.Add($label)}
  $pages=@('操作パネル','操作ガイド','材料データ','形状入力','メッシュ設定','接合','載荷','ステージ','解析設定','高速化設定','出力・表示','詳細設定','図','ステージ結果','接合結果')
  $order=@($wb.Worksheets|ForEach-Object{$_.Name});Check (($order[0..($pages.Count-1)]-join '|') -eq ($pages -join '|')) 'tab order'
  $mesh=$wb.Worksheets.Item('メッシュ設定');$mesh1=@($mesh.Shapes|Where-Object{$_.Name -eq 'UI_Mesh1'});Check ($mesh1.Count -eq 0) 'simple mesh button absent'
  for($r=16;$r -le 27;$r++){Check ($mesh.Rows.Item($r).Hidden) ('legacy simple row hidden '+$r)}
  foreach($name in $pages){$ws=$wb.Worksheets.Item($name);foreach($s in $ws.Shapes){$a='';try{$a=[string]$s.OnAction}catch{};if($a){Check ($a.Contains($wb.Name)) ($name+' qualified action '+$s.Name)}}}
  Check ([string]$app.Run($prefix+'ShapeFixture') -eq 'PASS') 'shape fixture generated'
  $counts=[string]$app.Run($prefix+'ShapeCounts');$parts=$counts.Split('|');Check ([int]$parts[0] -gt 0 -and [int]$parts[1] -gt 0) 'shape mesh has nodes/elements'
  Check ([string]$app.Run($prefix+'ShapePreview') -eq 'False|777|') 'main run previews mesh without analysis'
  $quality=[string]$app.Run($prefix+'ShapeQuality');Check ($quality -like '*品質目標=*最小角度=*最大アスペクト比=*') 'preview uses authoritative evaluated quality summary'
  $fig=$wb.Worksheets.Item('図');$svg=$fig.Shapes|Where-Object{$_.Name -like '*ViewSVG*'}|Select-Object -First 1;Check ($null -ne $svg) 'mesh SVG exists'
  Check (($mesh.Shapes|Where-Object{$_.Name -eq 'UI_Mesh2'}).Count -eq 1) 'shape mesh start button exists'
  $review=[string]$app.Run($prefix+'StartReview');Check ($review -eq 'False|777|') 'changed mesh requires review without calculation'
  $oldCounts=$counts
  $nx=0;for($r=1;$r -le 30;$r++){if([string]$mesh.Cells($r,6).Value2 -eq 'MESH_NX'){$nx=$r;break}}
  Check ($nx -gt 0) 'retired nx preserved for compatibility'
  $nxOld=$mesh.Cells($nx,3).Formula;$mesh.Cells($nx,3).Value2='invalid'
  Check ([string]$app.Run($prefix+'ShapeFixture') -eq 'PASS') 'invalid legacy nx ignored by shape mesh';Check ([string]$app.Run($prefix+'ShapeCounts') -eq $oldCounts) 'shape fixture remains stable'
  $mesh.Cells($nx,3).Formula=$nxOld
  function MeshSnapshot {return (@($wb.Worksheets.Item('節点データ').Range('A1:N500').Value2,$wb.Worksheets.Item('要素データ').Range('A1:J500').Value2)|ConvertTo-Json -Compress -Depth 5)}
  $modelBefore=MeshSnapshot
  $bad=[string]$app.Run($prefix+'BadShape');Check ($bad -like 'False|*' -and $bad.Length -gt 6) 'invalid shape rejected';Check ((MeshSnapshot) -ceq $modelBefore) 'invalid shape preserves all mesh input values'
  $repair=@($fig.Shapes|Where-Object{$_.Name -eq 'UI_MeshRepair'});Check ($repair.Count -eq 1) 'repair candidate button exists after preview'
  $candidate=[string]$app.Run($prefix+'CandidateFlow');Check ($candidate -like 'PASS|*') 'smooth candidate cancel/reprepare/adopt flow'
  $fig=$wb.Worksheets.Item('図');$start=@($fig.Shapes|Where-Object{$_.Name -eq 'UI_MeshStart'});Check ($start.Count -eq 1 -and $start[0].Visible) 'adopted mesh start button visible';Check (@($fig.Shapes|Where-Object{$_.Name -like '*ViewSVG*'}).Count -gt 0) 'adopted mesh SVG remains visible'
  $overlaps=[Collections.Generic.List[object]]::new();$buttonRects=[Collections.Generic.List[object]]::new()
  function Intersects($a,$b){return ($a.left -lt $b.left+$b.width-0.1 -and $a.left+$a.width -gt $b.left+0.1 -and $a.top -lt $b.top+$b.height-0.1 -and $a.top+$a.height -gt $b.top+0.1)}
  foreach($sheetName in $pages){
    $ws=$wb.Worksheets.Item($sheetName);$cells=[Collections.Generic.List[object]]::new();$seen=@{};$values=$ws.Range('A1:L100').Value2
    for($r=1;$r -le 100;$r++){for($c=1;$c -le 12;$c++){
      if(-not [string]$values.GetValue($r,$c)){continue};$cell=$ws.Cells($r,$c)
      if($cell.EntireColumn.Hidden -or $cell.EntireRow.Hidden){continue}
      $area=$cell.MergeArea;$address=[string]$area.Address();if($seen.ContainsKey($address)){continue};$seen[$address]=$true
      $cells.Add(@{name=$address;left=[double]$area.Left;top=[double]$area.Top;width=[double]$area.Width;height=[double]$area.Height})
    }}
    foreach($s in $ws.Shapes){
      $a='';try{$a=[string]$s.OnAction}catch{};if(-not $a -or -not $s.Visible){continue}
      $rect=@{page=$sheetName;name=$s.Name;left=[double]$s.Left;top=[double]$s.Top;width=[double]$s.Width;height=[double]$s.Height};$buttonRects.Add($rect)
      foreach($cell in $cells){if(Intersects $rect $cell){$overlaps.Add(@{page=$sheetName;button=$s.Name;cell=$cell.name})}}
    }
  }
  Check ($overlaps.Count -eq 0) ('buttons do not cover visible cell text: '+($overlaps|ConvertTo-Json -Compress))
  for($i=0;$i -lt $buttonRects.Count;$i++){for($j=$i+1;$j -lt $buttonRects.Count;$j++){if($buttonRects[$i].page -eq $buttonRects[$j].page){Check (-not (Intersects $buttonRects[$i] $buttonRects[$j])) ('buttons separated: '+$buttonRects[$i].name+'/'+$buttonRects[$j].name)}}}
  Check ((Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash -ceq $hash) 'source workbook unchanged'
  $r=@{status='PASS';source_sha256=$hash;checks=$checks.ToArray();count=$checks.Count}|ConvertTo-Json -Depth 5
  New-Item -ItemType Directory -Path (Split-Path -Parent (Join-Path $repo $Out)) -Force|Out-Null;$r|Set-Content -LiteralPath (Join-Path $repo $Out) -Encoding UTF8
  Write-Output ('PASS '+$checks.Count+' shape mesh workflow checks')
} finally {if($wb){try{$wb.Close($false)}catch{}};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)};[GC]::Collect();[GC]::WaitForPendingFinalizers();Pop-Location}
