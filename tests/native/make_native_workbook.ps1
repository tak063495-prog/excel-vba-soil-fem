param([string]$Src,[string]$Dst,[string]$ModDir)
$ErrorActionPreference='Stop'
$before=@(Get-Process EXCEL -ErrorAction SilentlyContinue|%{$_.Id})
if($before.Count -gt 0){throw "Another Excel is running: $before"}
Copy-Item -LiteralPath $Src -Destination $Dst
$xl=New-Object -ComObject Excel.Application
$xpid=@(Get-Process EXCEL|%{$_.Id}|?{$before -notcontains $_})[0]
Set-Content "$Dst.running" "pid=$xpid"
try{
 $xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($Dst,0,$false)
 foreach($m in 'FEMSolver','FEMIo','FEMCore'){
  $cm=$wb.VBProject.VBComponents.Item($m).CodeModule
  $cm.DeleteLines(1,$cm.CountOfLines)
  $cm.AddFromString((Get-Content (Join-Path $ModDir "$m.bas") -Raw -Encoding UTF8))
 }
 $c=$null; try{$c=$wb.VBProject.VBComponents.Item('FEMNative')}catch{}
 if($null -eq $c){$c=$wb.VBProject.VBComponents.Add(1);$c.Name='FEMNative'}
 $cm=$c.CodeModule; if($cm.CountOfLines -gt 0){$cm.DeleteLines(1,$cm.CountOfLines)}
 $cm.AddFromString((Get-Content (Join-Path $ModDir "FEMNative.bas") -Raw -Encoding UTF8))
 $ctl=$xl.VBE.CommandBars.FindControl(1,578)
 if($ctl.Enabled){$ctl.Execute()}
 "compile_enabled_after=$($xl.VBE.CommandBars.FindControl(1,578).Enabled)"
 $wb.Save(); $wb.Close($false)
} finally { $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null; Start-Sleep 3
 if(Get-Process -Id $xpid -ErrorAction SilentlyContinue){Stop-Process -Id $xpid -Force}
 Remove-Item "$Dst.running" -ErrorAction SilentlyContinue }
