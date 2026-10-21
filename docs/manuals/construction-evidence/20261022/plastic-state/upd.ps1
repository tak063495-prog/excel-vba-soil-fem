param($d)
function L($m){ "$(Get-Date -f HH:mm:ss) $m"|Out-File -Append "$d\upd.log" -Encoding utf8 }
function Norm([string]$s){ (($s -replace "`r`n","`n" -replace "`r","`n") -split "`n" | ?{ $_ -notmatch '^Attribute VB_' } | %{ $_.TrimEnd() }) -join "`n" -replace "`n+$","" }
$mods='FEMPlasticLog','FEMViewIncrement','FEMViewer','FEMResultOptions'
$wbdir="$env:USERPROFILE\manual_rewrite_20261021\workbook\ui_20261020"
$ex="$env:USERPROFILE\manual_rewrite_20261021\docs\manuals\examples"
$jobs=@(@("$wbdir\2DSoilFEM_ui_20261021.xlsm","$d\2DSoilFEM_ui_20261022.xlsm"))
if(Get-Process EXCEL -EA 0){ L "EXCEL running, abort"; exit }
$x=New-Object -ComObject Excel.Application
$me=(Get-Process EXCEL|Sort StartTime|Select -Last 1).Id; L "PID $me"
$x.DisplayAlerts=$false; $x.EnableEvents=$false; $x.AutomationSecurity=1
foreach($j in $jobs){ $src=$j[0];$dst=$j[1]
 try{
  if(Test-Path $dst){ throw "exists $dst" }
  Copy-Item $src $dst; $b=$x.Workbooks.Open($dst)
  foreach($m in $mods){ $cm=$b.VBProject.VBComponents.Item($m).CodeModule; $code=$cm.Lines(1,$cm.CountOfLines)
   $o=Get-Content "$d\orig\$m.bas" -Raw -Encoding UTF8
   if((Norm $code) -ne (Norm $o)){ throw "embedded $m differs from src orig" }
   $cm.DeleteLines(1,$cm.CountOfLines); $cm.AddFromString((Get-Content "$d\new\$m.bas" -Raw -Encoding UTF8)); L "replaced $m in $dst" }
  $ctl=$x.VBE.CommandBars.FindControl(1,578); if($ctl.Enabled){$ctl.Execute()}
  if($x.VBE.CommandBars.FindControl(1,578).Enabled){ throw "compile failed" }
  try{ $x.Run("'$($b.Name)'!FEMResultOptionsInstall")|Out-Null; L "install ok" }catch{ L "install ERR $_" }
  $b.Save(); $b.Close($false); L "saved $dst"
 }catch{ L "ERR $dst $_"; try{$b.Close($false)}catch{} }
}
$x.Quit(); Start-Sleep 5; Stop-Process -Id $me -EA 0; L "done"
