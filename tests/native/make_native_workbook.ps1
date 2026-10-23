<#
  ネイティブ版ブックの再構築。元ブックは上書きしない（-Dst へコピーしてから更新）。
  必須モジュール（$Required）は常に取り込む。-Modules は「追加分」として扱い、必須分を落とせない。
  Excel起動前にソース欠落を検出し、保存前と保存後（開き直し）に埋込コードとソースを照合する。
#>
param([string]$Src,[string]$Dst,[string]$ModDir,[string]$Modules='')
$ErrorActionPreference='Stop'
$Required=@('FEMFixedStressGauge','FEMFixedStressLedger','FEMFixedStressResults','FEMFixedStressInput','FEMWaterUi','FEMPlasticLog','FEMFixedStressSrm','FEMFixedStressMatrix','FEMFixedStress','FEMStageEditor','FEMWaterView','FEM','FEMWater','FEMSeepage','FEMSeepageFlux','FEMJointWater','FEMStageUi','MeshAngle','CONTROL','FEMMixedElement','FEMCore','FEMSolver','FEMIo','FEMEngine','FEMNative','FEMNativeMat','FEMNativeKern','FEMUiInputs','FEMUiValidation','FEMUiLayout','FEMUi','FEMViewer','FEMNativeFigure','ThisWorkbook','FEMUiMeshWorkflow','FEMUiGeometry','FEMPractical','MeshAdapt','Mesh','FEMTerrainBlocks','FEMResultOptions','FEMStrainMath','FEMViewIncrement','FEMIncrementSvg')
$list=[Collections.Generic.List[string]]::new(); foreach($m in $Required){$list.Add($m)}
foreach($m in ($Modules -split ',')){ $m=$m.Trim(); if($m -and -not $list.Contains($m)){$list.Add($m)} }
if(-not (Test-Path -LiteralPath $Src -PathType Leaf)){throw "Src not found: $Src"}
if(Test-Path -LiteralPath $Dst){throw "Dst already exists (refusing to overwrite): $Dst"}
if((Resolve-Path -LiteralPath $Src).Path -eq [IO.Path]::GetFullPath($Dst)){throw "Dst must differ from Src"}
function Norm([string]$s){ (($s -replace "`r`n","`n" -replace "`r","`n") -split "`n" | ?{ $_ -notmatch '^Attribute VB_' } | %{ $_.TrimEnd() }) -join "`n" -replace "`n+$","" }
$srcText=@{}; $srcHash=@{}
foreach($m in $list){
 $extension=if($m -eq 'ThisWorkbook'){'.cls'}else{'.bas'}
 $p=Join-Path $ModDir ($m+$extension)
 if(-not (Test-Path -LiteralPath $p -PathType Leaf)){throw "Module source missing: $p"}
 $srcText[$m]=Get-Content -LiteralPath $p -Raw -Encoding UTF8
 $srcHash[$m]=(Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash
}
function Verify($wb,$tag){
 foreach($m in $list){
  $c=$null; try{$c=$wb.VBProject.VBComponents.Item($m)}catch{}
  if($null -eq $c){throw "[$tag] module not embedded: $m"}
  $cm=$c.CodeModule; $code=$(if($cm.CountOfLines -gt 0){$cm.Lines(1,$cm.CountOfLines)}else{''})
  if((Norm $code) -ne (Norm $srcText[$m])){
   [IO.File]::WriteAllText($Dst+'.'+$m+'.actual.txt',$code,[Text.UTF8Encoding]::new($false))
   throw "[$tag] embedded code differs from source: $m"
  }
  "verify[$tag] $m OK sha256=$($srcHash[$m]) lines=$($cm.CountOfLines)"
 }
}
$before=@(Get-Process EXCEL -ErrorAction SilentlyContinue|%{$_.Id})
if($before.Count -gt 0){throw "Another Excel is running: $before"}
"modules=$($list -join ',')"
Copy-Item -LiteralPath $Src -Destination $Dst
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class NativeBuildIdentity {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
}
'@
$xl=New-Object -ComObject Excel.Application
[uint32]$ownedBuildPid=0
[void][NativeBuildIdentity]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd,[ref]$ownedBuildPid)
$xpid=[int]$ownedBuildPid
Set-Content "$Dst.running" "pid=$xpid"
$ok=$false
try{
 $xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($Dst,0,$false)
 foreach($m in $list){
  $c=$null; try{$c=$wb.VBProject.VBComponents.Item($m)}catch{}
  if($null -eq $c){$c=$wb.VBProject.VBComponents.Add(1);$c.Name=$m}
  $cm=$c.CodeModule
  if($cm.CountOfLines -gt 0){$cm.DeleteLines(1,$cm.CountOfLines)}
  $cm.AddFromString($srcText[$m])
 }
 Verify $wb 'pre-save'
 $ctl=$xl.VBE.CommandBars.FindControl(1,578)
 if($ctl.Enabled){$ctl.Execute()}
 $compEnabled=$xl.VBE.CommandBars.FindControl(1,578).Enabled
 "compile_enabled_after=$compEnabled"
 if($compEnabled){
   $pane=$xl.VBE.ActiveCodePane
   [int]$sl=0;[int]$sc=0;[int]$el=0;[int]$ec=0
   if($null -ne $pane){
    $pane.GetSelection([ref]$sl,[ref]$sc,[ref]$el,[ref]$ec)
    "compile_location=$($pane.CodeModule.Parent.Name):$sl column=$sc"
    if($sl -gt 0){$pane.CodeModule.Lines($sl,[Math]::Min(5,$pane.CodeModule.CountOfLines-$sl+1))}
   }
   throw "compile failed (Compile command still enabled)"
  }
 $wb.Save(); $wb.Close($false)
 $wb=$xl.Workbooks.Open($Dst,0,$true)
 Verify $wb 'reopen'
 $wb.Close($false)
 $ok=$true
} catch { "ERROR: $($_.Exception.Message)"; throw }
finally {
 foreach($comRef in @($cm,$c,$ctl,$wb)){
  if($null -ne $comRef -and [Runtime.InteropServices.Marshal]::IsComObject($comRef)){
   try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comRef)}catch{}
  }
 }
 $cm=$null;$c=$null;$ctl=$null;$wb=$null
 $xl.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);$xl=$null
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 # Only Quit the COM instance we created. Never terminate an Excel PID.
 Remove-Item "$Dst.running" -ErrorAction SilentlyContinue
 if(-not $ok){ Remove-Item -LiteralPath $Dst -ErrorAction SilentlyContinue; "FAILED: output removed" } }
"build=$([regex]::Match($srcText['FEMCore'],'FEM_BUILD_STAMP As String = "([^"]+)"').Groups[1].Value) (embedded FEMCore verified)"
"OK $Dst"
