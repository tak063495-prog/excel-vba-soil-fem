param([Parameter(Mandatory=$true)][string]$Book,[string]$Out='tests/tmp/manual_review/ui_checks.json')
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running. No test started.'}
$repo=Split-Path -Parent $PSScriptRoot;$Book=(Resolve-Path -LiteralPath $Book).Path
$copy=Join-Path (Split-Path $Book) ('ui_checks_'+[guid]::NewGuid().ToString('N')+'.xlsm');Copy-Item -LiteralPath $Book -Destination $copy
$checks=[Collections.Generic.List[string]]::new();$xl=$null;$wb=$null
function Check($ok,$label){if(-not $ok){throw $label};$checks.Add($label)}
function Snapshot($wb){$v=@();foreach($s in $wb.Worksheets){$v+=@{name=$s.Name;formula=$s.UsedRange.Formula}};return ($v|ConvertTo-Json -Depth 8 -Compress)}
try{
 $xl=New-Object -ComObject Excel.Application;$xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
 # Only six intentional source modules. Preserve all other embedded code.
 foreach($m in 'FEMUiLayout','FEMUiInputs','FEMIo','FEMViewIncrement','FEMIncrementSvg','FEMStrainMath'){
  $cm=$wb.VBProject.VBComponents.Item($m).CodeModule;$cm.DeleteLines(1,$cm.CountOfLines);$cm.AddFromString([IO.File]::ReadAllText((Join-Path $repo "src/vba/$m.bas")))
 }
 $refresh=[string]$xl.Run($prefix+'FEMUiRefreshInputs');Check ($refresh -ceq 'PASS') ('refresh updated source: '+$refresh)
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()};Check (-not $xl.VBE.CommandBars.FindControl(1,578).Enabled) 'Excel VBE compilation'
 $load=$wb.Worksheets.Item('載荷');$v=$load.Range('B2').Validation
 Check ($v.Type -eq 7 -and $v.ShowError -and $v.AlertStyle -eq 1) 'custom STOP validation installed'
 foreach($value in @('TOP','BOTTOM','LEFT','RIGHT','BOX',39,46,44,56,54,66,64,76,74)){
  $load.Range('B2').Formula=[string]$value;Check ([bool]$v.Value) "validation accepts $value"
 }
 foreach($value in @(-1,0,1.5,999999)){$load.Range('B2').Formula=[string]$value;Check (-not [bool]$v.Value) "validation rejects $value"}
 Check ([string]$load.Range('M12').Value2 -match 'DISP=0') 'zero DISP legend refreshed'
 # Row insertion + formula references must fail before any sheet data mutation.
 $rt=$wb.Worksheets.Item('高速化設定');$rt.Range('C11').Value2=[double]0;$rt.Range('C12').Formula='=C11'
 Check ([string]$xl.Run($prefix+'FEMUiRefreshInputs') -ceq 'PASS') 'canonical formulas preserved'
 $rt.Rows.Item(9).Insert()|Out-Null;$before=Snapshot $wb
 $failure=[string]$xl.Run($prefix+'FEMUiRefreshInputs')
 Check ($failure.StartsWith('ERROR|') -and $failure -match '数式') 'row insertion/formula migration refused'
 Check ((Snapshot $wb) -ceq $before) 'refused migration changes no worksheet data/formulas'
 $rt.Rows.Item(9).Delete()|Out-Null
 Check ([string]$xl.Run($prefix+'FEMUiRefreshInputs') -ceq 'PASS') 'canonical position restores formula references'
 # Observer threshold/snapshot/sign probe in test-only module.
 $h=$wb.VBProject.VBComponents.Add(1);$h.Name='ManualReviewHarness';$h.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $repo 'tests/native/ManualReviewHarness.bas')))
 $probe=[string]$xl.Run($prefix+'ManualReviewObserverProbe');Check ($probe -like 'PASS:*') ('observer synthetic zero/nonzero/history/snapshot/sign/color probe: '+$probe)
 $wb.Close($false);$wb=$null
 # Prepare distributable candidate with tested sources, without test harness/data.
 $wb=$xl.Workbooks.Open($Book,0,$false);$prefix="'"+$wb.Name+"'!"
 foreach($m in 'FEMUiLayout','FEMUiInputs','FEMIo','FEMViewIncrement','FEMIncrementSvg','FEMStrainMath'){
  $cm=$wb.VBProject.VBComponents.Item($m).CodeModule;$cm.DeleteLines(1,$cm.CountOfLines);$cm.AddFromString([IO.File]::ReadAllText((Join-Path $repo "src/vba/$m.bas")))
 }
 Check ([string]$xl.Run($prefix+'FEMUiRefreshInputs') -ceq 'PASS') 'candidate refresh'
 $st=$wb.Worksheets.Item('ステージ');$st.Range('F3').Value2='最後の収束Fsの状態を残す'
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()};Check (-not $xl.VBE.CommandBars.FindControl(1,578).Enabled) 'candidate compilation'
 $wb.Save();$wb.Close($false);$wb=$null
 $wb=$xl.Workbooks.Open($Book,0,$true);$prefix="'"+$wb.Name+"'!";Check ([string]$xl.Run($prefix+'FEMUiRefreshInputs') -ceq 'PASS') 'candidate reopen refresh'
 Check ($wb.Worksheets.Item('載荷').Range('B2').Validation.Type -eq 7) 'candidate validation persisted'
 Check ($wb.Worksheets.Item('ステージ').Range('F3').Value2 -match '収束Fs') 'candidate stage wording persisted'
 foreach($m in 'FEMUiLayout','FEMUiInputs','FEMIo','FEMViewIncrement','FEMIncrementSvg','FEMStrainMath'){
  $actual=$wb.VBProject.VBComponents.Item($m).CodeModule;$expected=[IO.File]::ReadAllText((Join-Path $repo "src/vba/$m.bas"))
  Check (($actual.Lines(1,$actual.CountOfLines).Trim() -replace "`r",'') -eq ($expected.Trim() -replace "`r",'')) ("embedded source matches $m")
 }
 $checks.ToArray()|ConvertTo-Json|Set-Content -LiteralPath $Out -Encoding UTF8
 Write-Output "PASS $($checks.Count) Excel checks; programmatic validation, not physical keyboard-dialog user testing"
}finally{if($wb){$wb.Close($false)};if($xl){$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)};[GC]::Collect();[GC]::WaitForPendingFinalizers()}
