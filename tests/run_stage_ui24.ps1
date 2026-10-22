param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
if(Test-Path $Out){throw 'Preserve earlier evidence.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory $Out | Out-Null
$copy=Join-Path $Out 'stage_ui24_test.xlsm';Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$checks=[Collections.Generic.List[object]]::new()
function Assert($ok,[string]$label){if(-not $ok){throw "FAIL: $label"};$checks.Add(@{label=$label;status='PASS'});"PASS $label"}
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$true;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false)
 $mod=$wb.VBProject.VBComponents.Add(1);$mod.Name='UI24TestDriver'
 $mod.CodeModule.AddFromString(@'
Public Function F24Quiet() As String
 SuppressUserMessages = True
 FEM24Error = "": FEM24StageError = ""
 F24Quiet = "PASS"
End Function
Public Function F24Error() As String
 F24Error = FEM24Error & FEM24StageError
End Function
Public Function F24Toggle() As String
 ThisWorkbook.Worksheets("結果").CheckBoxes("UI24Overlay0").Value = 1
 ThisWorkbook.Worksheets("結果").CheckBoxes("UI24Overlay1").Value = -4146
 FEM24Overlay
 F24Toggle = "PASS"
End Function
Public Function F24ReadOverlay() As String
 F24ReadOverlay = IIf(FEMOverlayModeA(), "A", "B")
End Function
'@)
 function Call([string]$name){$xl.Run("'"+$wb.Name+"'!"+$name)}
 function Quiet{$null=Call 'F24Quiet'}
 function Clean([string]$label){$err=Call 'F24Error';Assert ($err -eq '') "$label error=$err"}
 function Section([string]$page,[string]$section){
  $ws=$wb.Worksheets.Item($page);$ws.Activate();Quiet
  $ws.Range('C4').Value2=$section;Clean "section $page / $section"
 }
 function JsonRange([string]$sheet,[string]$address){ConvertTo-Json -InputObject $wb.Worksheets.Item($sheet).Range($address).Value2 -Depth 8 -Compress}
 Quiet
 $visible=@($wb.Worksheets | Where-Object Visible -eq -1 | ForEach-Object Name)
 Assert (($visible -join '|') -eq 'ホーム|モデル|材料・接合|境界・水条件|施工工程|解析・保存|結果') 'exactly seven normal tabs in workflow order'
 $sections=@{
  'ホーム'=@('作業の流れ');'モデル'=@('形状の座標','形状の領域','メッシュ生成・支持条件')
  '材料・接合'=@('材料特性','接合条件','材料の水理特性')
  '境界・水条件'=@('水条件の基本','水位線','水理境界','載荷条件','節点の支持・変位')
  '施工工程'=@('工程の編集');'解析・保存'=@('解析・SRM','保存・再開','高速化','詳細設定')
  '結果'=@('図と概要','表示設定','工程別数値','節点結果','要素結果','水圧結果','接合結果','診断')
 }
 foreach($page in $sections.Keys){foreach($sec in $sections[$page]){Section $page $sec}}
 $probe=$wb.Worksheets.Add();$probe.Name='ClipboardProbe24'
 Section '材料・接合' '材料特性'
 $probe.Visible=-1
 $wb.Worksheets.Item('材料・接合').Range('A8:B8').Copy()
 $null=Call 'FEM24Water'
 $probe.Activate();$probe.Range('A1').PasteSpecial(-4163)
 Assert ($probe.Range('A1').Value2 -eq 1 -and $probe.Range('B1').Value2 -eq 14000) 'navigation preserves user copy data'
 $xl.CutCopyMode=0;$probe.Delete()
 Section 'モデル' 'メッシュ生成・支持条件'
 $null=Call 'FEM24Detail'
 Assert ($wb.ActiveSheet.Name -eq 'メッシュ設定' -and $wb.Worksheets.Item('メッシュ設定').ProtectContents) 'detail navigation preserves canonical sheet protection'
 $null=Call 'FEM24Home'
 Assert (@($wb.Worksheets | Where-Object Visible -eq -1).Count -eq 7) 'return from detail restores seven tabs'

 Section '材料・接合' '材料特性'
 $wb.Worksheets.Item('材料・接合').Range('E8').Value2=21.25
 Assert ($wb.Worksheets.Item('材料データ').Range('E2').Value2 -eq 21.25) 'material write through'
 $xl.EnableEvents=$false;$wb.Worksheets.Item('材料データ').Range('B2').Value2=15000.0;$xl.EnableEvents=$true
 Quiet;$wb.Worksheets.Item('材料・接合').Range('E8').Value2=22.0
 Assert ((Call 'F24Error') -ne '') 'stale canonical input rejected before overwrite'
 Assert ($wb.Worksheets.Item('材料データ').Range('B2').Value2 -eq 15000 -and $wb.Worksheets.Item('材料データ').Range('E2').Value2 -eq 21.25) 'external edit retained after stale frontend rejection'
 $xl.EnableEvents=$false;$wb.Worksheets.Item('材料データ').Range('B2').Value2=14000.0;$xl.EnableEvents=$true
 $null=Call 'FEM24Materials';Quiet

 $wb.Worksheets.Item('材料・接合').Range('E8').Value2=20.0
 $preciseWeight=[double]20.123456789012344
 $xl.EnableEvents=$false;$wb.Worksheets.Item('材料データ').Range('E2').Value2=$preciseWeight;$wb.Worksheets.Item('材料データ').Range('C2').Formula='=0.1+0.2';$xl.EnableEvents=$true
 Section '材料・接合' '材料特性'
 Quiet;$wb.Worksheets.Item('材料・接合').Range('E8').Formula='=1'
 Assert ($wb.Worksheets.Item('材料データ').Range('E2').Value2 -eq $preciseWeight) 'rollback retains all digits of Double constants'
 Assert ($wb.Worksheets.Item('材料データ').Range('C2').Formula -eq '=0.1+0.2') 'rollback retains adjacent formulas'
 $xl.EnableEvents=$false;$wb.Worksheets.Item('材料データ').Range('E2').Value2=20.0;$wb.Worksheets.Item('材料データ').Range('C2').Value2=0.3;$xl.EnableEvents=$true
 Assert ($wb.Names.Item('_FEMViewerResultsDirty').RefersTo -eq '=TRUE') 'physical input invalidates saved results'
 Section 'モデル' '形状の座標'
 $wb.Worksheets.Item('モデル').Range('B8').Value2=-0.125
 Assert ($wb.Worksheets.Item('形状入力').Range('C8').Value2 -eq -0.125) 'geometry write through'
 $wb.Worksheets.Item('モデル').Range('B8').Value2=0.0
 $xl.EnableEvents=$false;$wb.Worksheets.Item('形状入力').Range('C8').Formula='=0+0';$xl.EnableEvents=$true
 Section 'モデル' '形状の座標'
 Quiet;$wb.Worksheets.Item('モデル').Range('B8').Formula='=1'
 Assert ((Call 'F24Error') -ne '') 'frontend formula input rejected'
 Assert ($wb.Worksheets.Item('形状入力').Range('C8').Formula -eq '=0+0') 'rejected input preserves canonical geometry formula'
 $xl.EnableEvents=$false;$wb.Worksheets.Item('形状入力').Range('C8').Value2=0.0;$xl.EnableEvents=$true
 Section '境界・水条件' '水条件の基本'
 $wb.Worksheets.Item('境界・水条件').Range('B8').Value2='HYDROSTATIC'
 Assert ($wb.Worksheets.Item('水条件').Range('B3').Value2 -eq 'HYDROSTATIC') 'water mode write through'
 Assert ($wb.Worksheets.Item('境界・水条件').Range('B11').Validation.Formula1 -eq 'AUTO,OFF') 'water basic rules are copied per row'
 $wb.Worksheets.Item('境界・水条件').Range('B11').Value2='OFF'
 Assert ($wb.Worksheets.Item('水条件').Range('B6').Value2 -eq 'OFF') 'external pressure option write through'
 $wb.Worksheets.Item('境界・水条件').Range('B11').Value2='AUTO'

 Section '境界・水条件' '水位線'
 $values=[object[,]]::new(2,4)
 $values[0,0]='W0';$values[0,1]=0.0;$values[0,2]=3.0;$values[0,3]='left'
 $values[1,0]='W0';$values[1,1]=20.0;$values[1,2]=2.0;$values[1,3]='right'
 $pasteArea=$wb.Worksheets.Item('境界・水条件').Range('A8:D9')
 $null=$pasteArea.GetType().InvokeMember('Value2',[Reflection.BindingFlags]::SetProperty,$null,$pasteArea,[object[]](,$values))
 Assert ($wb.Worksheets.Item('水条件').Range('C13').Value2 -eq 2.0) 'polyline table paste write through'
 Section '施工工程' '工程の編集'
 $stage=$wb.Worksheets.Item('施工工程');$raw=$wb.Worksheets.Item('ステージ');$water=$wb.Worksheets.Item('水条件')
 $null=Call 'FEM24StageAdd';Clean 'stage add'
 Assert ($raw.Range('A4').Value2 -eq 3 -and $raw.Range('B4').Value2 -eq 'GRAVITY') 'new stable stage ID'
 $stage.Range('B10').Value2='水位変更';Clean 'Japanese operation'
 $stage.Range('H10').Value2='W0';Clean 'water stage profile'
 Assert ($water.Range('L12').Value2 -eq 3 -and $water.Range('M12').Value2 -eq 'ONE_WAY' -and $water.Range('N12').Value2 -eq 'W0') 'new water plan and profile'
 Assert ([string]::IsNullOrEmpty($raw.Range('C4').Value2)) 'water profile has one writable frontend owner'
 $xl.EnableEvents=$false;$raw.Range('D4').Formula='=1+0';$xl.EnableEvents=$true
 $preciseStage=[double]1.0000000000000002
 $xl.EnableEvents=$false;$raw.Range('D2').Value2=$preciseStage;$xl.EnableEvents=$true
 Section '施工工程' '工程の編集'
 $stage.Range('B10').Select();$null=Call 'FEM24StageUp';Clean 'move stage up'
 Assert ($raw.Range('A3').Value2 -eq 3 -and $water.Range('L12').Value2 -eq 3) 'reorder preserves ID and water plan'
 Assert ($raw.Range('D3').Formula -eq '=1+0') 'reorder preserves stage formula'
 $null=Call 'FEM24StageCopy';Clean 'copy water stage'
 Assert ($raw.Range('A5').Value2 -eq 4 -and $water.Range('L13').Value2 -eq 4 -and $water.Range('N13').Value2 -eq 'W0') 'copy gives new ID with copied water plan'
 Assert ($raw.Range('D5').Formula -eq '=1+0') 'copy preserves stage formula'
 $lineBefore=JsonRange '水条件' 'A12:J211'
 $null=Call 'FEM24StageDelete';Clean 'delete water stage'
 Assert ([string]::IsNullOrEmpty($raw.Range('A5').Value2) -and [string]::IsNullOrEmpty($water.Range('L13').Value2)) 'delete removes only matching water plan'
 Assert ((JsonRange '水条件' 'A12:J211') -eq $lineBefore) 'delete does not shift water/material tables'
 $null=Call 'FEM24StageUndo';Clean 'undo delete'
 Assert ($raw.Range('A5').Value2 -eq 4 -and $water.Range('L13').Value2 -eq 4) 'undo restores stages and water plans'
 Assert ($raw.Range('D5').Formula -eq '=1+0') 'undo preserves stage formula'
 Assert ($raw.Range('D2').Value2 -eq $preciseStage) 'stage undo retains all Double digits'
 Quiet;$stage.Range('E9').Value2=0.0;Clean 'disable water stage with retained plan'
 Assert ($water.Range('N12').Value2 -eq 'W0' -and $raw.Range('E3').Value2 -eq 0.0) 'disabled stage retains water plan'
 $stage.Range('E9').Value2=1.0;Clean 'reenable water stage'
 Assert ($water.Range('N12').Value2 -eq 'W0' -and $raw.Range('E3').Value2 -eq 1.0) 'reenabled stage restores same profile'

 $before=JsonRange 'ステージ' 'A2:F2000';$plansBefore=JsonRange '水条件' 'L12:T211'
 Quiet;$stage.Range('H9').Value2='UNREGISTERED'
 Assert ((Call 'F24Error') -ne '') 'unknown profile rejected'
 Assert ((JsonRange 'ステージ' 'A2:F2000') -eq $before -and (JsonRange '水条件' 'L12:T211') -eq $plansBefore) 'unknown profile rolls back both tables'
 Assert ($raw.Range('D3').Formula -eq '=1+0') 'rejected stage edit preserves formula'
 Quiet;$stage.Range('G9').Value2='FIXED_STRESS'
 Assert ((Call 'F24Error') -ne '') 'unsupported coupling rejected'
 Quiet;$stage.Range('E8').Value2='abc'
 Assert ((Call 'F24Error') -ne '') 'nonnumeric enabled flag rejected'
 Assert ((JsonRange 'ステージ' 'A2:F2000') -eq $before) 'invalid enabled flag rolled back'
 Assert $xl.EnableEvents 'events restored after failed edits'
 Quiet
 $null=Call 'FEMViewerMarkClean'
 Section '結果' '表示設定'
 $row=0
 for($r=8;$r -lt 28;$r++){if($wb.Worksheets.Item('結果').Cells.Item($r,15).Value2 -eq 17){$row=$r;break}}
 Assert ($row -gt 0) 'display scale row mapped'
 $wb.Worksheets.Item('結果').Cells.Item($row,4).Value2=2.0
 Assert ($wb.Worksheets.Item('出力・表示').Range('C17').Value2 -eq 2.0) 'display setting write through'
 Assert ($wb.Names.Item('_FEMViewerResultsDirty').RefersTo -eq '=FALSE') 'display change preserves mechanical result validity'
 $xl.EnableEvents=$false;$wb.Worksheets.Item('出力・表示').Range('C17').Formula='=1+1';$xl.EnableEvents=$true
 Section '結果' '表示設定'
 Quiet;$wb.Worksheets.Item('結果').Cells.Item($row,4).Formula='=3'
 Assert ((Call 'F24Error') -ne '') 'frontend setting formula rejected'
 Assert ($wb.Worksheets.Item('出力・表示').Range('C17').Formula -eq '=1+1') 'rejected setting preserves canonical formula'
 Quiet
 $preciseScale=[double]1.1234567890123457
 $xl.EnableEvents=$false;$wb.Worksheets.Item('出力・表示').Range('C30').Value2=$preciseScale;$xl.EnableEvents=$true
 Section '結果' '表示設定'
 $crossRow=0;for($r=8;$r -lt 28;$r++){if($wb.Worksheets.Item('結果').Cells.Item($r,15).Value2 -eq 30){$crossRow=$r;break}}
 Quiet;$wb.Worksheets.Item('結果').Cells.Item($crossRow,4).Formula='=3'
 Assert ($wb.Worksheets.Item('出力・表示').Range('C30').Value2 -eq $preciseScale) 'setting rollback retains all Double digits'
 Quiet
 $modeRow=0
 for($r=8;$r -lt 28;$r++){if($wb.Worksheets.Item('結果').Cells.Item($r,15).Value2 -eq 27){$modeRow=$r;break}}
 Assert ($modeRow -gt 0 -and $modeRow -le 20) 'overlay controls immediately follow display settings'
 $wb.Worksheets.Item('結果').Cells.Item($modeRow,4).Value2='B'
 Assert ($wb.Worksheets.Item('出力・表示').Range('C27').Value2 -eq 'B：増分専用カラー') 'overlay A/B write through'
 Assert ((Call 'F24ReadOverlay') -eq 'B') 'actual viewer accepts frontend overlay mode'
 $null=Call 'F24Toggle'
 Assert ($wb.Worksheets.Item('出力・表示').Range('J28').Value2 -eq $true -and $wb.Worksheets.Item('出力・表示').Range('J29').Value2 -eq $false) 'independent overlay checkboxes write through'
 Assert ($wb.Names.Item('_FEMViewerResultsDirty').RefersTo -eq '=FALSE') 'overlay changes preserve result validity'

 Section '解析・保存' '解析・SRM'
 $wb.Worksheets.Item('解析・保存').Range('D9').Value2='DAVIS'
 Assert ($wb.Worksheets.Item('解析設定').Range('C10').Value2 -eq 'DAVIS') 'solver policy write through'
 Assert ($wb.Names.Item('_FEMViewerResultsDirty').RefersTo -eq '=TRUE') 'solver change invalidates results'
 $null=Call 'FEM24Home';$wb.Save();$wb.Close($false)
 $wb=$xl.Workbooks.Open($copy,0,$false)
 Assert ($wb.Worksheets.Item('ステージ').Range('A3').Value2 -eq 3) 'stable stage ID survives normal reopen'
 Assert ($wb.Worksheets.Item('水条件').Range('N12').Value2 -eq 'W0') 'water plan survives normal reopen'
 Assert (@($wb.Worksheets | Where-Object Visible -eq -1).Count -eq 7) 'normal reopen restores seven tabs'
 $wb.Close($false)
 $record=@{status='PASS';excel_version=$xl.Version;excel_build=$xl.Build;checks=@($checks);book_sha256=(Get-FileHash $Book).Hash}
}finally{
 $xl.EnableEvents=$false
 if($wb){try{$wb.Close($false)}catch{}}
 foreach($o in @($probe,$water,$raw,$stage,$ws,$mod,$wb)){if($o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}catch{}}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
$record | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $Out 'checks.json') -Encoding utf8
'ALL PASS'
