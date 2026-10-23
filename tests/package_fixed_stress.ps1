param([Parameter(Mandatory=$true)][string]$Candidate,[Parameter(Mandatory=$true)][string]$Destination,[Parameter(Mandatory=$true)][string]$Report)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no dispatch.'}
$Candidate=(Resolve-Path $Candidate).Path;$Destination=[IO.Path]::GetFullPath($Destination)
if(Test-Path $Destination){throw 'Preserve previous package.'}
$sources=@(Get-ChildItem "$PSScriptRoot/../src/vba" -File | Where-Object {$_.Extension -in '.bas','.cls'})
function Norm([string]$s){(($s -replace "`r`n","`n" -replace "`r","`n") -split "`n" | Where-Object {$_ -notmatch '^Attribute VB_'} | ForEach-Object {$_.TrimEnd()}) -join "`n" -replace "`n+$",""}
function Verify($book){
 foreach($file in $sources){
  $mod=$book.VBProject.VBComponents.Item($file.BaseName).CodeModule
  if((Norm ([IO.File]::ReadAllText($file.FullName))) -ne (Norm $mod.Lines(1,$mod.CountOfLines))){throw "Embedded source mismatch: $($file.Name)"}
 }
}
$areas=@{
 '節点データ'='A2:I2001';'要素データ'='A2:K2001';'材料データ'='A2:M2001';'ステージ'='A2:F2000';'載荷'='A2:K2000';'形状入力'='A1:Q2000';
 '水条件'='A12:T211';'接合'='A2:G2000'
}
$snap=@{};$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1
 Copy-Item $Candidate $Destination
 $wb=$xl.Workbooks.Open($Destination,0,$false)
 foreach($sheet in $areas.Keys){$snap[$sheet]=ConvertTo-Json -InputObject $wb.Worksheets.Item($sheet).Range($areas[$sheet]).Formula -Depth 5 -Compress}
 $settings=ConvertTo-Json -InputObject $wb.Worksheets.Item('設定').UsedRange.Formula -Depth 5 -Compress
 $null=$xl.Run("'"+$wb.Name+"'!FEMWaterInstallUi")
 $null=$xl.Run("'"+$wb.Name+"'!FEM24Install")
 $null=$xl.Run("'"+$wb.Name+"'!FEM24Home")
 $null=$xl.Run("'"+$wb.Name+"'!FEMViewerMarkDirty")
 Verify $wb
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()}
 if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Compile failed'}
 $wb.Save();$wb.Close($false);$wb=$xl.Workbooks.Open($Destination,0,$true)
 Verify $wb
 foreach($sheet in $areas.Keys){
  $actual=ConvertTo-Json -InputObject $wb.Worksheets.Item($sheet).Range($areas[$sheet]).Formula -Depth 5 -Compress
  if($actual -ne $snap[$sheet]){throw "Canonical input changed: $sheet"}
 }
 if((ConvertTo-Json -InputObject $wb.Worksheets.Item('設定').UsedRange.Formula -Depth 5 -Compress) -ne $settings){throw 'Calculation settings changed'}
 if($wb.Worksheets.Item('水条件').Range('B3').Value2 -ne 'NONE'){throw 'Default water changed'}
 if($wb.Worksheets.Item('水条件').Range('AK11').Value2 -ne '載荷セットID'){throw 'Time header not installed'}
 if($wb.Names.Item('_FEMViewerResultsDirty').RefersTo -ne '=TRUE'){throw 'Old results not invalidated'}
 $record=@{status='PASS';compile='PASS';source_verification='before-save and reopen';module_count=$sources.Count;modules=@($sources.BaseName);canonical_inputs_unchanged=@($areas.Keys);calculation_settings_unchanged=$true;excel_version=$xl.Version;excel_build=$xl.Build;candidate_sha256=(Get-FileHash $Candidate).Hash}
 $wb.Close($false)
}catch{
 Write-Output $_.ScriptStackTrace
 throw
}finally{
 if($wb){try{$wb.Close($false)}catch{}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
$record.package_sha256=(Get-FileHash $Destination).Hash
$record|ConvertTo-Json -Depth 8|Set-Content $Report -Encoding utf8
