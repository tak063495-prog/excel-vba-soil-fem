param([Parameter(Mandatory=$true)][string]$Candidate,[Parameter(Mandatory=$true)][string]$Destination,[Parameter(Mandatory=$true)][string]$Report)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$Candidate=(Resolve-Path -LiteralPath $Candidate).Path
$Destination=[IO.Path]::GetFullPath($Destination)
if($Candidate -eq $Destination){throw 'Candidate must be preserved.'}
if(Test-Path -LiteralPath $Destination){throw 'Destination already exists.'}
$names=@('FEMMixedElement','MeshAngle','CONTROL','FEM','FEMCore','FEMSolver','FEMIo','FEMEngine','FEMNative','FEMNativeMat','FEMNativeKern','FEMUiInputs','FEMUiValidation','FEMUiLayout','FEMUi','FEMViewer','FEMNativeFigure','ThisWorkbook','FEMUiMeshWorkflow','FEMUiGeometry','FEMPractical','MeshAdapt','Mesh','FEMTerrainBlocks','FEMResultOptions','FEMStrainMath','FEMViewIncrement','FEMIncrementSvg','FEMWater','FEMSeepage','FEMWaterUi','FEMWaterGeometry','FEMWaterView','FEMSeepageFlux','FEMJointWater','FEMStageUi','FEMStageEditor')
function Norm([string]$s){(($s -replace "`r`n","`n" -replace "`r","`n") -split "`n" | Where-Object {$_ -notmatch '^Attribute VB_'} | ForEach-Object {$_.TrimEnd()}) -join "`n" -replace "`n+$",""}
function Verify($book){
 foreach($name in $names){
  $extension=if($name -eq 'ThisWorkbook'){'.cls'}else{'.bas'}
  $source=[IO.File]::ReadAllText((Join-Path "$PSScriptRoot/../src/vba" ($name+$extension)))
  $module=$book.VBProject.VBComponents.Item($name).CodeModule
  if((Norm $source) -ne (Norm $module.Lines(1,$module.CountOfLines))){throw "Embedded source mismatch: $name"}
 }
}
$snapshots=@{}
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1
 Copy-Item -LiteralPath $Candidate -Destination $Destination
 $wb=$xl.Workbooks.Open($Destination,0,$false)
 foreach($sheet in '節点データ','要素データ','材料データ','ステージ','載荷','形状入力'){
  $range=$wb.Worksheets.Item($sheet).UsedRange
  if($sheet -eq '要素データ'){$range=$wb.Worksheets.Item($sheet).Range('A1:J'+$wb.Worksheets.Item($sheet).Cells($wb.Worksheets.Item($sheet).Rows.Count,1).End(-4162).Row)}
  $snapshots[$sheet]=ConvertTo-Json -InputObject $range.Value2 -Depth 5 -Compress
 }
 $elementTypes=ConvertTo-Json -InputObject $wb.Worksheets.Item("要素データ").Range("K2:K"+[Math]::Max(2,$wb.Worksheets.Item("要素データ").Cells($wb.Worksheets.Item("要素データ").Rows.Count,1).End(-4162).Row)).Value2 -Depth 5 -Compress
 $null=$xl.Run("'"+$wb.Name+"'!FEMWaterInstallUi")
 $refresh=[string]$xl.Run("'"+$wb.Name+"'!FEMUiRefreshInputs")
 if($refresh -ne "PASS"){throw "Input migration failed: $refresh"}
 if($wb.Worksheets.Item("メッシュ設定").Range("F30").Value2 -ne "MESH_T6_FALLBACK" -or $wb.Worksheets.Item("メッシュ設定").Range("C30").Value2 -ne 0){throw "Release mesh fallback option/default missing."}
 $null=$xl.Run("'"+$wb.Name+"'!FEMViewerMarkDirty")
 $water=$wb.Worksheets.Item('水条件')
 if($water.Range('B3').Value2 -ne 'NONE'){throw 'Release default must be NONE.'}
 $null=$xl.Run("'"+$wb.Name+"'!FEM24Install")
 $null=$xl.Run("'"+$wb.Name+"'!FEM24Home")
 Verify $wb
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()}
 if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Compile failed.'}
 $wb.Save();$wb.Close($false)
 $wb=$xl.Workbooks.Open($Destination,0,$true)
 $water=$wb.Worksheets.Item('水条件')
 Verify $wb
 foreach($sheet in $snapshots.Keys){
  $range=$wb.Worksheets.Item($sheet).UsedRange
  if($sheet -eq '要素データ'){$range=$wb.Worksheets.Item($sheet).Range('A1:J'+$wb.Worksheets.Item($sheet).Cells($wb.Worksheets.Item($sheet).Rows.Count,1).End(-4162).Row)}
  $actual=ConvertTo-Json -InputObject $range.Value2 -Depth 5 -Compress
  if($actual -ne $snapshots[$sheet]){throw "Canonical input changed: $sheet"}
 }
 $actualTypes=ConvertTo-Json -InputObject $wb.Worksheets.Item("要素データ").Range("K2:K"+[Math]::Max(2,$wb.Worksheets.Item("要素データ").Cells($wb.Worksheets.Item("要素データ").Rows.Count,1).End(-4162).Row)).Value2 -Depth 5 -Compress
 if($actualTypes -ne $elementTypes){throw "Element types changed during packaging."}
 if($wb.Names.Item('_FEMViewerResultsDirty').RefersTo -ne '=TRUE'){throw 'Old results must be marked dirty.'}
 $record=@{status='PASS';compile='PASS';source_verification='before-save and reopen';modules=$names;canonical_inputs_unchanged=@($snapshots.Keys);element_types_unchanged=$true;mesh_t6_fallback_default=0;defaults=@($water.Range('B3:B6').Value2);excel_version=$xl.Version;excel_build=$xl.Build;candidate_sha256=(Get-FileHash $Candidate).Hash}
 $wb.Close($false)
}catch{
 $_.ScriptStackTrace | Write-Output
 throw
}finally{
 foreach($reference in @($water,$ctl,$module,$wb)){if($reference -and [Runtime.InteropServices.Marshal]::IsComObject($reference)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference)}catch{}}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
$record.ui_stamp='20261024_STAGE_01'
$record.package_sha256=(Get-FileHash $Destination).Hash
$record | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Report -Encoding UTF8
'PASS: compiled, saved, reopened; source and canonical inputs verified.'
