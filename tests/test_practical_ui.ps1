$ErrorActionPreference='Stop'
$rootTask=Split-Path -Parent $PSScriptRoot
Push-Location $rootTask
try {
$mappingTask=Get-Content -LiteralPath tests/fixtures/ui_mapping.json -Raw -Encoding UTF8 | ConvertFrom-Json
$inventoryTask=Get-Content -LiteralPath tests/fixtures/ui_inventory.json -Raw -Encoding UTF8 | ConvertFrom-Json
$copyTask=Join-Path $pwd 'tests/tmp/practical_ui.xlsm'
Copy-Item -LiteralPath workbook/2DSoilFEM_20261008_practical.xlsm -Destination $copyTask -Force
$xlTask=New-Object -ComObject Excel.Application
$checksTask=[Collections.Generic.List[string]]::new()
function UiCheck($condition,$label){if(-not $condition){throw $label};$checksTask.Add('PASS '+$label); if($checksTask.Count % 50 -eq 0){Write-Output ('Checked '+$checksTask.Count+' '+$label)}}
try {
  $xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1
  $wbTask=$xlTask.Workbooks.Open($copyTask,0,$false)
  $tmTask=$wbTask.VBProject.VBComponents.Add(1);$tmTask.Name='UiTestHarness'
  $tmTask.CodeModule.AddFromString(@'
Public Function UiTestRead(ByVal key As String, ByVal numericValue As Boolean) As Variant
  On Error GoTo Failed
  If numericValue Then UiTestRead = FEMReadSetting(key, -999999#) Else UiTestRead = FEMReadTextSetting(key, "")
  Exit Function
Failed:
  UiTestRead = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Sub UiTestArmState()
  AnalysisRunning = False: AnalysisOK = True: ResultRevision = 123: P1ResultReady = True
End Sub
Public Function UiTestState() As String
  UiTestState = CStr(AnalysisOK) & "|" & CStr(ResultRevision) & "|" & CStr(P1ResultReady)
End Function
Public Function UiTestBindingReject() As String
  On Error GoTo Rejected
  FEMUiBindSettings
  UiTestBindingReject = "NOT_REJECTED"
  Exit Function
Rejected:
  UiTestBindingReject = "REJECTED|" & CStr(Err.Number)
End Function
Public Function UiTestCellError(ByVal sheetName As String, ByVal address As String) As Boolean
  UiTestCellError = IsError(ThisWorkbook.Worksheets(sheetName).Range(address).Value2)
End Function
Public Function UiTestLayout() As String
  On Error GoTo Failed
  FEMUiAfterLayout
  UiTestLayout = "PASS"
  Exit Function
Failed:
  UiTestLayout = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function UiTestRebuild() As String
  On Error GoTo Failed
  FEMWriteBuildStamp False
  UiTestRebuild = "PASS"
  Exit Function
Failed:
  UiTestRebuild = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
'@)
  $prefixTask="'"+$wbTask.Name+"'!"
  $layoutResultTask=$xlTask.Run($prefixTask+'UiTestLayout'); if($layoutResultTask -ne 'PASS'){throw $layoutResultTask}
  $xlTask.CalculateFull()
  foreach($mTask in $mappingTask){
    $originalTask=$inventoryTask.Settings | Where-Object {$_.Key -ceq $mTask.key}
    $actualTask=$wbTask.Worksheets.Item($mTask.sheet).Range($mTask.cell).Value2
    $numericTask=$originalTask.Value -is [ValueType]
    $sameInputTask=$(if($numericTask){[double]$actualTask -eq [double]$originalTask.Value}else{[string]$actualTask -ceq [string]$originalTask.Value})
    UiCheck $sameInputTask ('input '+$mTask.key+' actual='+[string]$actualTask+' expected='+[string]$originalTask.Value)
    $readTask=$xlTask.Run($prefixTask+'UiTestRead',$mTask.key,$numericTask)
    $sameReadTask=$(if($numericTask){[double]$readTask -eq [double]$originalTask.Value}else{[string]$readTask -ceq [string]$originalTask.Value})
    UiCheck $sameReadTask ('reader '+$mTask.key+' actual='+[string]$readTask+' expected='+[string]$originalTask.Value)
    $formulaTask=[string]$wbTask.Worksheets.Item('設定').Cells($mTask.legacyRow,3).Formula
    UiCheck ($formulaTask.Replace("'",'').Contains($mTask.sheet+'!'+$mTask.cell)) ('binding '+$mTask.key)
    if($null -ne $originalTask.Validation){
      $typesTask=@{whole=1;decimal=2;list=3;date=4;time=5;textLength=6;custom=7}
      $operatorsTask=@{between=1;notBetween=2;equal=3;notEqual=4;greaterThan=5;lessThan=6;greaterThanOrEqual=7;lessThanOrEqual=8}
      $vTask=$wbTask.Worksheets.Item($mTask.sheet).Range($mTask.cell).Validation
      UiCheck ($vTask.Type -eq $typesTask[$originalTask.Validation.type]) ('validation type '+$mTask.key)
      if($originalTask.Validation.type -eq 'list'){
        $expectedRuleTask=[string]$originalTask.Validation.formula1
        if($expectedRuleTask.StartsWith('"')){$expectedRuleTask=$expectedRuleTask.Trim('"')}else{$expectedRuleTask='='+$expectedRuleTask.TrimStart('=')}
        UiCheck ($vTask.Formula1 -ceq $expectedRuleTask) ('choices '+$mTask.key+' actual='+$vTask.Formula1)
      } elseif($originalTask.Validation.type -eq 'custom'){
        $expectedRuleTask='='+[regex]::Replace(([string]$originalTask.Validation.formula1).TrimStart('='),'\bC'+$mTask.legacyRow+'\b',$mTask.cell)
        UiCheck ($vTask.Formula1 -ceq $expectedRuleTask) ('validation custom reference '+$mTask.key)
        UiCheck $vTask.Value ('validation accepts preserved value '+$mTask.key)
      } else {
        UiCheck ($vTask.Operator -eq $operatorsTask[$originalTask.Validation.operator]) ('validation operator '+$mTask.key)
        UiCheck ([double]$vTask.Formula1 -eq [double]$originalTask.Validation.formula1) ('validation limit '+$mTask.key)
      }
    }
  }
  $xlTask.EnableEvents=$true
  $accTask=$mappingTask | Where-Object {$_.key -eq 'ACCEL_STEP_RECOVERY'}
  $oldTask=$wbTask.Worksheets.Item($accTask.sheet).Range($accTask.cell).Value2
  $xlTask.Run($prefixTask+'UiTestArmState')
  $wbTask.Worksheets.Item($accTask.sheet).Range($accTask.cell).Value2=[double]0
  UiCheck ($xlTask.Run($prefixTask+'UiTestRead','ACCEL_STEP_RECOVERY',$true) -eq 0) 'edit invalidates cached physical setting'
  UiCheck ($xlTask.Run($prefixTask+'UiTestState') -eq 'False|0|False') 'physical edit invalidates analysis'
  $wbTask.Worksheets.Item($accTask.sheet).Range($accTask.cell).Value2=$oldTask
  foreach($keyTask in @('VIEW_DISP_SCALE','DEBUG_FLUSH','EXPORT_LOAD','RCM_POLICY','ADAPT_MARK_FRACTION')){
    $mTask=$mappingTask | Where-Object {$_.key -eq $keyTask}
    $cellTask=$wbTask.Worksheets.Item($mTask.sheet).Range($mTask.cell);$valueTask=$cellTask.Value2
    $xlTask.Run($prefixTask+'UiTestArmState')
    if($valueTask -is [ValueType]){$cellTask.Value2=[double]([double]$valueTask+0.01)}else{$cellTask.Value2=[string]'OFF'}
    UiCheck ($xlTask.Run($prefixTask+'UiTestState') -eq 'True|123|True') ('nonphysical edit retains analysis '+$keyTask)
    if($null -eq $valueTask){$cellTask.ClearContents() | Out-Null}elseif($valueTask -is [ValueType]){$cellTask.Value2=[double]$valueTask}else{$cellTask.Value2=[string]$valueTask}
  }
  $xlTask.Calculation=-4135
  $cellTask=$wbTask.Worksheets.Item($accTask.sheet).Range($accTask.cell)
  $cellTask.Value2=[double]0
  UiCheck ($xlTask.Run($prefixTask+'UiTestRead','ACCEL_STEP_RECOVERY',$true) -eq 0) 'manual calculation reads changed setting'
  $xlTask.Run($prefixTask+'FEMWriteNumericSetting','accel_step_recovery',[double]1)
  UiCheck ($cellTask.Value2 -eq 1) 'lowercase VBA writer updates canonical cell'
  UiCheck ($wbTask.Worksheets.Item('設定').Cells($accTask.legacyRow,3).HasFormula) 'writer preserves registry formula'
  $xlTask.Run($prefixTask+'FEMWriteTextSetting','EXPORT_LOAD','AUTO')
  UiCheck ($xlTask.Run($prefixTask+'UiTestRead','EXPORT_LOAD',$false) -eq 'AUTO') 'text writer canonical input'
  $xlTask.Run($prefixTask+'FEMWriteTextSetting','EXPORT_LOAD','')
  $layoutResultTask=$xlTask.Run($prefixTask+'UiTestRebuild'); if($layoutResultTask -ne 'PASS'){throw $layoutResultTask}
  UiCheck ($xlTask.Run($prefixTask+'UiTestRead','EXPORT_LOAD',$false) -eq '') 'blank resume remains blank after layout'
  UiCheck ($cellTask.Value2 -eq 1 -and $wbTask.Worksheets.Item('設定').Cells($accTask.legacyRow,3).HasFormula) 'layout rebuild retains values and binding'
  $xlTask.Calculation=-4105
  $loadTask=$wbTask.Worksheets.Item('載荷');$loadTask.Range('H2').Value2=[double]2.5;$loadTask.Range('I2').Value2=[double]3.5
  $layoutResultTask=$xlTask.Run($prefixTask+'UiTestRebuild'); if($layoutResultTask -ne 'PASS'){throw $layoutResultTask}
  UiCheck ($loadTask.Range('H2').Value2 -eq 2.5 -and $loadTask.Range('I2').Value2 -eq 3.5) 'layout retains BOX X2/Y2 inputs'
  UiCheck ($wbTask.Worksheets.Item('ステージ').Range('B101').Validation.Formula1.Contains('BIRTH')) 'stage dropdown beyond row 100'
  UiCheck ($wbTask.Worksheets.Item('接合').Range('D101').Validation.Formula1 -eq '1,0') 'joint activation dropdown'
  $wsTask=$wbTask.Worksheets.Item($accTask.sheet);$wsTask.Unprotect()
  $metaTask=$wsTask.Cells($accTask.row,6);$keySavedTask=$metaTask.Value2
  $beforeFormulaTask=$wbTask.Worksheets.Item('設定').Cells($accTask.legacyRow,3).Formula
  $metaTask.Value2='ACCEL_ADAPTIVE'
  UiCheck ($xlTask.Run($prefixTask+'UiTestBindingReject').StartsWith('REJECTED|')) 'duplicate mapping is rejected'
  UiCheck ($wbTask.Worksheets.Item('設定').Cells($accTask.legacyRow,3).Formula -eq $beforeFormulaTask) 'invalid mapping does not partially rebind'
  $metaTask.Value2=[string]$keySavedTask
  $layoutResultTask=$xlTask.Run($prefixTask+'UiTestLayout'); if($layoutResultTask -ne 'PASS'){throw $layoutResultTask}
  foreach($wsTask in $wbTask.Worksheets){
    foreach($shapeTask in $wsTask.Shapes){
      try{$actionTask=$shapeTask.OnAction}catch{continue}
      if($actionTask){UiCheck ($actionTask.Replace("'",'').StartsWith($wbTask.Name+'!')) ('button workbook '+$shapeTask.Name+' '+$actionTask)}
    }
  }
  $homeTask=$wbTask.Worksheets.Item('操作パネル')
  UiCheck ($homeTask.Hyperlinks.Count -eq 11) 'eleven workflow navigation links'
  foreach($linkTask in $homeTask.Hyperlinks){UiCheck (-not $linkTask.SubAddress.Contains('[')) 'internal hyperlink destination'}
  $xlTask.Run($prefixTask+'FEMUiShowNodes');UiCheck ($wbTask.ActiveSheet.Name -eq '節点データ') 'raw input button opens hidden sheet'
  $xlTask.Run($prefixTask+'FEMUiShowHome');UiCheck ($wbTask.ActiveSheet.Name -eq '操作パネル') 'return to operation panel'
  foreach($mTask in $mappingTask){UiCheck (-not $xlTask.Run($prefixTask+'UiTestCellError','設定','C'+$mTask.legacyRow)) ('formula result '+$mTask.key)}
  $checksTask | Set-Content -LiteralPath tests/results/practical_ui_checks_20261008.txt -Encoding UTF8
  Write-Output ('PASS '+$checksTask.Count+' UI binding, edit, writer, dropdown, navigation and preservation checks.')
  $wbTask.Close($false)
} finally {
  if($null -ne $wbTask){try{$wbTask.Close($false)}catch{}}
  try{$xlTask.EnableEvents=$false;$xlTask.Quit()}catch{}
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)|Out-Null
}

} finally { Pop-Location }
