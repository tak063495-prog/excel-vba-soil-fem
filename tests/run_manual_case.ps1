# Run from repository root with desktop Excel and trusted VBA-project access.
# Only a fresh, dedicated copy is edited. Existing Excel processes cause an abort.
# The public workbook, its VBA, the DLL and solver convergence settings are retained.
# -Anchor is a distributed, balanced prestress load example, not an anchor element.
param([Parameter(Mandatory=$true)][string]$Fixture,[Parameter(Mandatory=$true)][string]$OutDirectory,[int]$StageLimit=999,[switch]$Srm,[switch]$Anchor,[ValidateSet('','DAVIS','INCONSISTENT')][string]$Flow='',[string]$SourceWorkbook='workbook/ui_20261020/2DSoilFEM_ui_20261021.xlsm')
$ErrorActionPreference='Stop';$app=$null;$wb=$null
$repo=(Get-Location).Path
$source=(Resolve-Path -LiteralPath $SourceWorkbook).Path;$hash=(Get-FileHash -LiteralPath $source).Hash
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running. No test started.'}
$fixtureData=Get-Content -LiteralPath $Fixture -Raw -Encoding UTF8|ConvertFrom-Json
if(Test-Path -LiteralPath $OutDirectory){throw 'Output exists.'};New-Item -ItemType Directory -Path $OutDirectory|Out-Null
$OutDirectory=(Resolve-Path -LiteralPath $OutDirectory).Path;$book=Join-Path $OutDirectory 'case.xlsm'
Copy-Item -LiteralPath $source -Destination $book;Copy-Item -LiteralPath (Join-Path (Split-Path $source) 'soilfem_native_x64.dll') -Destination $OutDirectory
try{
 $app=New-Object -ComObject Excel.Application;$app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $wb=$app.Workbooks.Open($book,0,$false);$prefix="'"+$wb.Name+"'!"
 [void]$app.Run($prefix+'FEMWriteBuildStamp',$false);[void]$app.Run($prefix+'SetP0SilentMode',$true)
 $moduleSnapshot=@{};foreach($m in $wb.VBProject.VBComponents){$moduleSnapshot[$m.Name]=if($m.CodeModule.CountOfLines -gt 0){$m.CodeModule.Lines(1,$m.CodeModule.CountOfLines)}else{''}}
 function PutRows([string]$page,$rows,[int]$cols,[int]$start=2,[int]$column=1){
  if($rows.Count -eq 0){return};$ws=$wb.Worksheets.Item($page);$ws.Unprotect();$v=[object[,]]::new($rows.Count,$cols)
  for($r=0;$r -lt $rows.Count;$r++){if($rows[$r].Count -ne $cols){throw "$page invalid cols $r"};for($c=0;$c -lt $cols;$c++){$v[$r,$c]=$rows[$r][$c]}}
  $ws.Range($ws.Cells($start,$column),$ws.Cells($start+$rows.Count-1,$column+$cols-1)).Value2=$v
 }
 foreach($page in '材料データ','節点データ','要素データ','ステージ','載荷','接合','結果節点','結果要素','ステージ結果','結果増分'){
  try{$ws=$wb.Worksheets.Item($page)}catch{continue};$ws.Unprotect();$last=[Math]::Max(2000,$ws.UsedRange.Rows.Count);$ws.Range($ws.Cells(2,1),$ws.Cells($last,[Math]::Max(20,$ws.UsedRange.Columns.Count))).ClearContents()
 }
 PutRows '材料データ' $fixtureData.materials 13;PutRows 'ステージ' $fixtureData.stages 6
 foreach($entry in @(@('OUTPUT_STAGE_MODE','ALL'),@('OUTPUT_INACTIVE_ELEMENTS','SKIP'),@('EXPORT_LOAD',''),@('EXPORT_MODE','STAGE'),@('VIEW_DETAIL_MODE','SIMPLE'))){[void]$app.Run($prefix+'FEMWriteTextSetting',[string]$entry[0],[string]$entry[1])}
 foreach($entry in @(@('VIEW_DISP_SCALE',1.0),@('VIEW_COLOR_RESULT',2.0))){[void]$app.Run($prefix+'FEMWriteNumericSetting',[string]$entry[0],[double]$entry[1])}
 if($Flow -ne ''){[void]$app.Run($prefix+'FEMWriteTextSetting','FLOW_POLICY',$Flow)}
 # Do not retain the source workbook's unrelated shape when importing a saved mesh.
 $ws=$wb.Worksheets.Item('形状入力');$ws.Unprotect();$ws.Range('B8:D2000').ClearContents();$ws.Range('G8:J2000').ClearContents()
 if($fixtureData.points){
  PutRows '形状入力' $fixtureData.points 3 8 2;PutRows '形状入力' $fixtureData.boundaries 4 8 7
  foreach($entry in @(@('MESH_BC_BOTTOM','PINNED'),@('MESH_BC_LEFT','ROLLER'),@('MESH_BC_RIGHT','ROLLER'),@('MESH_BC_TOP','NONE'),@('MESH_BC_PIN_CORNER','AUTO'))){[void]$app.Run($prefix+'FEMWriteTextSetting',[string]$entry[0],[string]$entry[1])}
  [void]$app.Run($prefix+'FEMUiGenerateShape')
 }else{PutRows '節点データ' $fixtureData.nodes 9;PutRows '要素データ' $fixtureData.elements 10}
 $nodes=$wb.Worksheets.Item('節点データ');$nodeCount=$nodes.Cells($nodes.Rows.Count,1).End(-4162).Row-1
 $elements=$wb.Worksheets.Item('要素データ');$elementCount=$elements.Cells($elements.Rows.Count,1).End(-4162).Row-1
 if($nodeCount -lt 8 -or $elementCount -lt 1){throw 'No mesh generated.'}
 if($Anchor){
  $nodeValues=$nodes.Range('A2:C'+($nodeCount+1)).Value2
  $elValues=$elements.Range('A2:J'+($elementCount+1)).Value2
  $wallNodes=@{};$soilNodes=@{}
  for($r=1;$r -le $elementCount;$r++){for($c=2;$c -le 9;$c++){if($elValues[$r,10] -eq 2){$wallNodes[[int]$elValues[$r,$c]]=$true};if($elValues[$r,10] -eq 1){$soilNodes[[int]$elValues[$r,$c]]=$true}}}
  $headSet=@();$bondSet=@()
  for($r=1;$r -le $nodeCount;$r++){
   $id=[int]$nodeValues[$r,1];$x=[double]$nodeValues[$r,2];$y=[double]$nodeValues[$r,3]
   if($wallNodes.ContainsKey($id) -and [Math]::Abs($x-18.6) -lt 1e-6 -and $y -ge 2.0 -and $y -le 2.8){$headSet+=,@($id,$x,$y,[Math]::Abs($y-2.4))}
   if($soilNodes.ContainsKey($id) -and $x -ge 6 -and $x -le 9 -and $y -ge -1.5 -and $y -le 0){$bondSet+=,@($id,$x,$y,0)}
  }
  $headSet=@($headSet|Sort-Object {$_[3]}|Select-Object -First 3)
  if($headSet.Count -lt 2 -or $bondSet.Count -lt 2){throw 'Anchor patches have insufficient nodes.'}
  $hx=0.0;$hy=0.0;$bx=0.0;$by=0.0
  foreach($n in $headSet){$hx+=$n[1];$hy+=$n[2]};$hx/=$headSet.Count;$hy/=$headSet.Count
  foreach($n in $bondSet){$bx+=$n[1];$by+=$n[2]};$bx/=$bondSet.Count;$by/=$bondSet.Count
  $length=[Math]::Sqrt(($bx-$hx)*($bx-$hx)+($by-$hy)*($by-$hy));$fx=50.0*($bx-$hx)/$length;$fy=50.0*($by-$hy)/$length
  $loads=[Collections.Generic.List[object]]::new()
  foreach($n in $headSet){$loads.Add(@(2.0,[double]$n[0],'X','FORCE',[double]($fx/$headSet.Count),$null,$null,$null,$null,'頭部受圧部へ分配'));$loads.Add(@(2.0,[double]$n[0],'Y','FORCE',[double]($fy/$headSet.Count),$null,$null,$null,$null,'頭部受圧部へ分配'))}
  foreach($n in $bondSet){$loads.Add(@(1.0,[double]$n[0],'X','FORCE',[double](-$fx/$bondSet.Count),$null,$null,$null,$null,'定着側の等価分布力'));$loads.Add(@(1.0,[double]$n[0],'Y','FORCE',[double](-$fy/$bondSet.Count),$null,$null,$null,$null,'定着側の等価分布力'))}
  PutRows '載荷' $loads.ToArray() 10
  $anchorMeta=@{head_nodes=$headSet;bond_nodes=$bondSet;head_centroid=@($hx,$hy);bond_centroid=@($bx,$by);force_x=$fx;force_y=$fy;force_norm=50.0;anchor_per_piece_kN=100;spacing_m=2;angle_degrees=([Math]::Atan2([Math]::Abs($fy),[Math]::Abs($fx))*180.0/[Math]::PI)}
  $last=$fixtureData.stages.Count+2;PutRows 'ステージ' (,@(($last-1),'LOAD',$null,$null,1,'アンカー相当力を同時に載荷')) 6 $last
 }else{if($fixtureData.loads){PutRows '載荷' $fixtureData.loads 10}}
 $ws=$wb.Worksheets.Item('ステージ');$last=$ws.Cells($ws.Rows.Count,1).End(-4162).Row
 if($Srm){PutRows 'ステージ' (,@($last,'SRM',$null,$null,1,'施工後のSRM')) 6 ($last+1);$last++}
 if($StageLimit -ne 999){for($r=2;$r -le $last;$r++){$ws.Cells($r,5).Value2=[double]([int](($r-1) -le $StageLimit))}}
 [void]$app.Run($prefix+'FEMInvalidateInputCache');[void]$app.Run($prefix+'FEMInvalidateSettingCache')
 $compile=$app.VBE.CommandBars.FindControl(1,578);if($compile.Enabled){$compile.Execute()};if($app.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Excel compile failed.'}
 $wb.Save();$timer=[Diagnostics.Stopwatch]::StartNew();[void]$app.Run($prefix+'P0_RunAnalysis');$timer.Stop()
 function MatrixRows($matrix){$rows=[Collections.Generic.List[object]]::new();for($r=$matrix.GetLowerBound(0);$r -le $matrix.GetUpperBound(0);$r++){$row=@();for($c=$matrix.GetLowerBound(1);$c -le $matrix.GetUpperBound(1);$c++){$row+=,$matrix.GetValue($r,$c)};$rows.Add($row)};return ,$rows.ToArray()}
 $tables=@{};foreach($entry in @(@('節点データ',9),@('要素データ',10),@('材料データ',13),@('ステージ',6),@('載荷',10),@('ステージ結果',20),@('結果節点',12),@('結果要素',72),@('結果増分',20),@('診断',5),@('設定',5))){$ws=$wb.Worksheets.Item($entry[0]);$lastRow=$ws.Cells($ws.Rows.Count,1).End(-4162).Row;$tables[$entry[0]]=MatrixRows $ws.Range($ws.Cells(1,1),$ws.Cells([Math]::Max(2,$lastRow),[int]$entry[1])).Value2}
 $status=[string]$wb.Worksheets.Item('診断').Cells(2,2).Value2
 $result=@{anchor=$anchorMeta;status=$status;seconds=$timer.Elapsed.TotalSeconds;nodes=$nodeCount;elements=$elementCount;source_sha256=$hash;fixture=$Fixture;tables=$tables}
 $result|ConvertTo-Json -Depth 13|Set-Content -LiteralPath (Join-Path $OutDirectory 'result.json') -Encoding UTF8
 if($status -eq 'PASS'){
  foreach($mode in 'A','B'){
   $opt=$wb.Worksheets.Item('出力・表示');$opt.Unprotect();$opt.Range('C27').Value2=if($mode -eq 'A'){'A：選択結果へ等値線を追加'}else{'B：増分専用カラー'};$opt.Range('J28').Value2=$true;$opt.Range('J29').Value2=$true
   [void]$app.Run($prefix+'FEMViewResult');$wb.SaveCopyAs((Join-Path $OutDirectory ($mode+'.xlsm')))
  }
  # Leave the usable example on the user's preferred base view A.
  $opt.Unprotect();$opt.Range('C27').Value2='A：選択結果へ等値線を追加'
  [void]$app.Run($prefix+'FEMViewResult')
 }
 foreach($m in $wb.VBProject.VBComponents){$text=if($m.CodeModule.CountOfLines -gt 0){$m.CodeModule.Lines(1,$m.CodeModule.CountOfLines)}else{''};if(($moduleSnapshot.ContainsKey($m.Name) -and $moduleSnapshot[$m.Name] -cne $text) -or (-not $moduleSnapshot.ContainsKey($m.Name) -and $text.Length -gt 0)){throw "Code changed $($m.Name)"}}
 $result['application_code_unchanged']=$true;$result|ConvertTo-Json -Depth 13|Set-Content -LiteralPath (Join-Path $OutDirectory 'result.json') -Encoding UTF8
 $wb.SaveAs((Join-Path $OutDirectory '2DSoilFEM_ui_20261020.xlsm'),52);$wb.Close($false);$wb=$null
 if((Get-FileHash -LiteralPath $source).Hash -cne $hash){throw 'Source workbook changed.'}
 Write-Output "$status nodes=$nodeCount elements=$elementCount seconds=$($timer.Elapsed.TotalSeconds)"
}catch{Write-Output $_.ScriptStackTrace;throw}finally{if($wb){try{$wb.Close($false)}catch{}};if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
