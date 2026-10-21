param($d)
. "$d\common.ps1"
function L($m){ "$(Get-Date -f HH:mm:ss) $m"|Out-File -Append "$d\run11.log" -Encoding utf8 }
Start-Sleep 10
Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","$d\watch.ps1",$d -WindowStyle Hidden
if(Get-Process EXCEL -EA 0){ L "EXCEL already running, abort"; "x"|Out-File "$d\stop"; exit }
New-Item -ItemType Directory "$d\img" -Force | Out-Null
$x=New-Object -ComObject Excel.Application
$me=(Get-Process EXCEL|Sort StartTime|Select -Last 1).Id; "$me"|Out-File "$d\mypid.txt"; L "PID $me"
$x.Visible=$true; $x.WindowState=-4137; $x.DisplayAlerts=$false
$script:b=$null
Add-Type @"
using System;using System.Runtime.InteropServices;
public class PW{[DllImport("user32.dll")]public static extern bool PrintWindow(IntPtr h,IntPtr dc,uint f);}
"@
function Snap($name,$addrs){ Start-Sleep -Milliseconds 1500; $h=[IntPtr]$x.Hwnd; (New-Object -ComObject WScript.Shell).AppActivate([int]$me)|Out-Null; [U]::SetForegroundWindow($h)|Out-Null; Start-Sleep -Milliseconds 2000; $r=New-Object U+R; [U]::GetWindowRect($h,[ref]$r)|Out-Null; $w=$r.Rt-$r.L;$ht=$r.B-$r.T; $bm=New-Object Drawing.Bitmap $w,$ht; $g=[Drawing.Graphics]::FromImage($bm); $hdc=$g.GetHdc(); [PW]::PrintWindow($h,$hdc,2)|Out-Null; $g.ReleaseHdc($hdc); $bm.Save("$d\img\$name.png"); $g.Dispose();$bm.Dispose()
  $rs=@(); $pn=$x.ActiveWindow.ActivePane; $ws=$x.ActiveSheet
  foreach($a in $addrs){ try{ $rg=$ws.Range($a); $rs+=@{a=$a;l=$pn.PointsToScreenPixelsX($rg.Left)-$r.L;t=$pn.PointsToScreenPixelsY($rg.Top)-$r.T;r=$pn.PointsToScreenPixelsX($rg.Left+$rg.Width)-$r.L;b=$pn.PointsToScreenPixelsY($rg.Top+$rg.Height)-$r.T} }catch{} }
  (@{name=$name;rects=$rs}|ConvertTo-Json -Compress -Depth 4)|Out-File -Append "$d\rects.jsonl" -Encoding utf8; L "snap $name" }
function Go($n,$row=1){ [void]$script:b.Activate(); $ws=$script:b.Worksheets.Item($n); $ws.Visible=-1; [void]$ws.Activate(); $x.ActiveWindow.Zoom=100; $x.ActiveWindow.ScrollRow=$row; $x.ActiveWindow.ScrollColumn=1; $ws.Range("A1").Select()|Out-Null; return $ws }
function M($m,$lab){ Start-Sleep 3; $lab|Out-File "$d\label.txt" -Encoding utf8; L "run $m"; $x.EnableEvents=$true; try{ $x.Run($m)|Out-Null; L "ret $m" }catch{ L "ERR $m $($_.Exception.Message)" }; $x.EnableEvents=$false; Start-Sleep 8 }
function MS($k,$v){ try{ if($v -is [string]){ $x.Run('FEMWriteTextSetting',$k,$v)|Out-Null } else { $x.Run('FEMWriteNumericSetting',$k,[double]$v)|Out-Null } }catch{ L "ERRSET $k $_" } }
function OpenF($from,$tag){ if($script:b){ $script:b.Close($false); Start-Sleep 2 }; $f="$d\2DSoilFEM_$($tag)_20261022_$(Get-Date -f HHmm).xlsm"; Copy-Item $from $f; Copy-Item (Join-Path (Split-Path $from) 'soilfem_native_x64.dll') $d -Force -EA 0; $script:b=$x.Workbooks.Open($f); $x.EnableEvents=$false; L "open $f from $from"; return $f }
function DumpSheet($sheet,$file,$cols,$maxr=3000){ try{ $j=$script:b.Worksheets.Item($sheet); $lr=$j.Cells($j.Rows.Count,1).End(-4162).Row; $t=@(); for($r=1;$r -le [Math]::Min($lr,$maxr);$r++){ $row=@(); for($c=1;$c -le $cols;$c++){ $row+=[string]$j.Cells($r,$c).Value2 }; $t+=($row -join "`t") }; ($t -join "`n")|Out-File "$d\$file" -Encoding utf8; L "dump $sheet $file" }catch{ L "nodump $sheet $_" } }
function View($tag,$sel,$scale){ $ws=Go '出力・表示'; MS 'VIEW_COLOR_RESULT' $sel; MS 'VIEW_DISP_SCALE' $scale; $ws=Go '出力・表示'; Snap "$($tag)_out" @('C17','C18'); M 'FEMViewResult' "$tag-view"; Snap "$($tag)_fig" @() }
$steps=[ordered]@{}
$ex="$env:USERPROFILE\manual_rewrite_20261021\docs\manuals\examples"
function RunA($lab){ $x.Run('SetP0SilentMode',$true)|Out-Null; $x.Run('FEMWriteTextSetting','OUTPUT_STAGE_MODE','ALL')|Out-Null; $x.Run('FEMInvalidateInputCache')|Out-Null; $x.Run('FEMInvalidateSettingCache')|Out-Null; "1"|Out-File "$d\analyzing"; $t=[Diagnostics.Stopwatch]::StartNew(); M 'P0_RunAnalysis' $lab; $t.Stop(); Remove-Item "$d\analyzing" -EA 0; L "elapsed $lab $($t.Elapsed.TotalSeconds)" }
function JointDumps($tag){ $ws=Go 'ステージ結果'; Snap "$($tag)_result" @('A1:J3'); DumpSheet 'ステージ結果' "$($tag)_stage.tsv" 20; DumpSheet '診断' "$($tag)_diag.tsv" 5 400
  DumpSheet '結果要素' "$($tag)_elems.tsv" 72 2000; DumpSheet '要素データ' "$($tag)_eldata.tsv" 10 2000
  $ws=Go '操作パネル'; M 'FEMPracticalShowJointResults' "$tag-jres"; Snap "$($tag)_jres" @('G1:H25','I1:I25','L1:L25'); DumpSheet '接合結果' "$($tag)_joint.tsv" 14 200 }

function OpenD($f){ if($script:b){ $script:b.Close($false); Start-Sleep 2 }; $script:b=$x.Workbooks.Open($f); $x.EnableEvents=$false; L "open $f" }
function Lab($t){ $ws=Go '出力・表示'; $ws.Range('C18').Value2=$t; Start-Sleep 1 }
$steps['old']={ OpenD "$d\ex02_old_20261022.xlsm"; RunA 'old'; DumpSheet 'ステージ結果' 'old_stage.tsv' 20; DumpSheet '結果要素' 'old_elems.tsv' 72 3000; DumpSheet '結果節点' 'old_nodes.tsv' 12 3000; $script:b.Save() }
$steps['new']={ OpenD "$d\ex02_new_20261022.xlsm"; RunA 'new'; DumpSheet 'ステージ結果' 'new_stage.tsv' 20; DumpSheet '結果要素' 'new_elems.tsv' 72 3000; DumpSheet '結果節点' 'new_nodes.tsv' 12 3000; DumpSheet '結果塑性' 'new_plastic.tsv' 13 3000; $script:b.Save()
  Lab '塑性状態（弾性／現在降伏／過去に降伏・要素の4点）'; MS 'VIEW_DISP_SCALE' 1; $ws=Go '出力・表示'; Snap 'st_out' @('C18'); M 'FEMViewResult' 'st-view'; Snap 'st_fig' @()
  Lab '累積塑性せん断ひずみ（4点最大）'; $ws=Go '出力・表示'; Snap 'cum_out' @('C18'); M 'FEMViewResult' 'cum-view'; Snap 'cum_fig' @(); $script:b.Save() }
foreach($k in $steps.Keys){ L "step $k"; try{ & $steps[$k] }catch{ L "ERRSTEP $k $_" } }
$script:b.Close($false); $x.Quit(); Start-Sleep 5; Stop-Process -Id $me -EA 0; "x"|Out-File "$d\stop"; L "all done"
