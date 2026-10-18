param(
 [string]$SourceWorkbook=(Join-Path $PSScriptRoot '../../workbook/2DSoilFEM_issue23_20261017_1031.xlsm'),
 [string]$OutputWorkbook=(Join-Path $PSScriptRoot '../../workbook/2DSoilFEM_standard48_native_20261018.xlsm'),
 [string]$ModDir=(Join-Path $PSScriptRoot '../../src/vba'),
 [Parameter(Mandatory=$true)][string]$EvidenceDirectory
)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be idle before building.'}
$principal48=[Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
if($principal48.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'Normal user token required.'}
if((Test-Path -LiteralPath $OutputWorkbook) -or (Test-Path -LiteralPath $EvidenceDirectory)){throw 'Preserve existing build and evidence.'}
$source48=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$output48=[IO.Path]::GetFullPath($OutputWorkbook)
$sourceHash48=(Get-FileHash -LiteralPath $source48 -Algorithm SHA256).Hash
$module48=@{}
foreach($file48 in Get-ChildItem -LiteralPath $ModDir -File | Where-Object {$_.Extension -in '.bas','.cls'}){
 $module48[$file48.BaseName]=[IO.File]::ReadAllText($file48.FullName,[Text.Encoding]::UTF8)
}
if($module48.Count -lt 25 -or -not $module48.ContainsKey('FEMNativeKern')){throw 'Incomplete source set.'}
New-Item -ItemType Directory -Path $EvidenceDirectory | Out-Null
Copy-Item -LiteralPath $source48 -Destination $output48
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class Standard48BuildIdentity {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint processId);
}
'@
function Norm48([string]$s){
 # VBE canonicalizes identifier case and trailing whitespace during AddFromString.
 ((($s -replace "`r`n","`n" -replace "`r","`n") -split "`n" | ForEach-Object {$_.TrimEnd()}) -join "`n").TrimEnd("`n").ToUpperInvariant()
}
function Verify48($book48){
 foreach($name48 in $module48.Keys){
  $cm48=$book48.VBProject.VBComponents.Item($name48).CodeModule
  $code48=if($cm48.CountOfLines){$cm48.Lines(1,$cm48.CountOfLines)}else{''}
  if((Norm48 $code48) -cne (Norm48 $module48[$name48])){throw ('Embedded source mismatch: '+$name48)}
 }
}
$app48=$null;$book48=$null;$ok48=$false;$failure48=$null;$identity48=$null
try{
 $app48=New-Object -ComObject Excel.Application
 [uint32]$pid48=0;[void][Standard48BuildIdentity]::GetWindowThreadProcessId([IntPtr]$app48.Hwnd,[ref]$pid48)
 $p48=Get-Process -Id $pid48 -ErrorAction Stop
 $identity48=@{pid=$pid48;birth_utc_ticks=$p48.StartTime.ToUniversalTime().Ticks.ToString();excel_version=$p48.MainModule.FileVersionInfo.FileVersion;binary_sha256=(Get-FileHash -LiteralPath $p48.Path -Algorithm SHA256).Hash}
 $identity48 | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $EvidenceDirectory 'owned_excel.json') -Encoding utf8
 if(@(Get-Process EXCEL).Count -ne 1){throw 'Foreign Excel appeared; no workbook modified.'}
 $app48.Visible=$false;$app48.EnableEvents=$false;$app48.DisplayAlerts=$false;$app48.AutomationSecurity=1
 $book48=$app48.Workbooks.Open($output48,0,$false)
 foreach($name48 in $module48.Keys){
  $component48=$null;try{$component48=$book48.VBProject.VBComponents.Item($name48)}catch{}
  if($null -eq $component48){
   if($name48 -eq 'ThisWorkbook'){throw 'Missing document module.'}
   $component48=$book48.VBProject.VBComponents.Add(1);$component48.Name=$name48
  }
  $cm48=$component48.CodeModule
  if($cm48.CountOfLines){$cm48.DeleteLines(1,$cm48.CountOfLines)}
  if($module48[$name48].Length){$cm48.AddFromString($module48[$name48])}
 }
 Verify48 $book48
 $compile48=$app48.VBE.CommandBars.FindControl(1,578)
 if($compile48.Enabled){$compile48.Execute()}
 if($app48.VBE.CommandBars.FindControl(1,578).Enabled){throw 'VBA compile failed.'}
 $macro48="'"+$book48.Name+"'!"
 $app48.Run($macro48+'FEMWriteTextSetting','FLOW_POLICY','INCONSISTENT')
 $app48.Run($macro48+'FEMWriteTextSetting','NATIVE_DLL','AUTO')
 $app48.Run($macro48+'FEMInvalidateSettingCache')
 $book48.Save();$book48.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book48);$book48=$null
 $book48=$app48.Workbooks.Open($output48,0,$true)
 Verify48 $book48
 $macro48="'"+$book48.Name+"'!"
 $flow48=[string]$app48.Run($macro48+'P6ReadTextSetting','FLOW_POLICY','MISSING')
 $native48=[string]$app48.Run($macro48+'P6ReadTextSetting','NATIVE_DLL','MISSING')
 if($flow48 -cne 'INCONSISTENT' -or $native48 -cne 'AUTO'){throw 'Persisted default KEY mismatch.'}
 $ok48=$true
}catch{$failure48=$_.Exception.ToString();throw}finally{
 if($null -ne $book48){$book48.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book48);$book48=$null}
 if($null -ne $app48){$app48.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app48);$app48=$null}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
 for($wait48=0;$wait48 -lt 20;$wait48++){
  if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){break};Start-Sleep -Milliseconds 500
 }
 $live48=$false
 if($null -ne $identity48){$p48=Get-Process -Id $identity48.pid -ErrorAction SilentlyContinue;if($null -ne $p48){$live48=($p48.StartTime.ToUniversalTime().Ticks.ToString() -ceq $identity48.birth_utc_ticks)}}
 $record48=@{status=$(if($ok48 -and -not $live48){'PASS_BUILD_REOPEN_SOURCE_KEYS'}else{'FAILED_PRESERVED'});failure=$failure48;source_sha256=$sourceHash48;module_count=$module48.Count;compile_pass=$ok48;reopened_sources_exact=$ok48;flow_policy=$flow48;native_dll=$native48;owned_excel_ended=(-not $live48);output_sha256=(Get-FileHash -LiteralPath $output48 -Algorithm SHA256).Hash}
 $record48 | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $EvidenceDirectory 'result.json') -Encoding utf8
}
if($live48){throw 'Owned Excel still present; no forced termination.'}
if((Get-FileHash -LiteralPath $source48 -Algorithm SHA256).Hash -cne $sourceHash48){throw 'Original workbook changed.'}
$record48 | ConvertTo-Json -Depth 6
