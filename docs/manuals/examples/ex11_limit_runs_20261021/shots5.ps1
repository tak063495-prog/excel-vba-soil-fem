param($d)
. "$d\common.ps1"
function L($m){ "$(Get-Date -f HH:mm:ss) $m"|Out-File -Append "$d\shots5.log" -Encoding utf8 }
while(-not ((Test-Path "$d\stop") -and ((gc "$d\shots3.log" -EA 0) -match "all done"))){ Start-Sleep 15 }
Start-Sleep 10; Remove-Item "$d\stop"; Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","$d\watch.ps1",$d -WindowStyle Hidden
Start-Sleep 10
if(Get-Process EXCEL -EA 0){ L "EXCEL already running, abort"; "x"|Out-File "$d\stop"; exit }
New-Item -ItemType Directory "$d\img" -Force | Out-Null
$x=New-Object -ComObject Excel.Application
$me=(Get-Process EXCEL|Sort StartTime|Select -Last 1).Id; "$me"|Out-File "$d\mypid.txt"; L "PID $me"
$x.Visible=$true; $x.WindowState=-4137; $x.DisplayAlerts=$false
$script:b=$null
function Snap($name,$addrs){ Start-Sleep -Milliseconds 1500; $h=[IntPtr]$x.Hwnd; (New-Object -ComObject WScript.Shell).AppActivate([int]$me)|Out-Null; [U]::SetForegroundWindow($h)|Out-Null; Start-Sleep -Milliseconds 2000; $r=New-Object U+R; [U]::GetWindowRect($h,[ref]$r)|Out-Null; $w=$r.Rt-$r.L;$ht=$r.B-$r.T; $bm=New-Object Drawing.Bitmap $w,$ht; $g=[Drawing.Graphics]::FromImage($bm); $g.CopyFromScreen($r.L,$r.T,0,0,$bm.Size); $bm.Save("$d\img\$name.png"); $g.Dispose();$bm.Dispose()
  $rs=@(); $pn=$x.ActiveWindow.ActivePane; $ws=$x.ActiveSheet
  foreach($a in $addrs){ try{ $rg=$ws.Range($a); $rs+=@{a=$a;l=$pn.PointsToScreenPixelsX($rg.Left)-$r.L;t=$pn.PointsToScreenPixelsY($rg.Top)-$r.T;r=$pn.PointsToScreenPixelsX($rg.Left+$rg.Width)-$r.L;b=$pn.PointsToScreenPixelsY($rg.Top+$rg.Height)-$r.T} }catch{} }
  (@{name=$name;rects=$rs}|ConvertTo-Json -Compress -Depth 4)|Out-File -Append "$d\rects.jsonl" -Encoding utf8; L "snap $name" }
function Go($n,$row=1){ [void]$script:b.Activate(); $ws=$script:b.Worksheets.Item($n); $ws.Visible=-1; [void]$ws.Activate(); $x.ActiveWindow.Zoom=100; $x.ActiveWindow.ScrollRow=$row; $x.ActiveWindow.ScrollColumn=1; $ws.Range("A1").Select()|Out-Null; return $ws }
function M($m,$lab){ Start-Sleep 3; $lab|Out-File "$d\label.txt" -Encoding utf8; L "run $m"; $x.EnableEvents=$true; try{ $x.Run($m)|Out-Null; L "ret $m" }catch{ L "ERR $m $($_.Exception.Message)" }; $x.EnableEvents=$false; Start-Sleep 8 }
function MS($k,$v){ try{ if($v -is [string]){ $x.Run('FEMWriteTextSetting',$k,$v)|Out-Null } else { $x.Run('FEMWriteNumericSetting',$k,[double]$v)|Out-Null } }catch{ L "ERRSET $k $_" } }
function OpenF($from,$tag){ if($script:b){ $script:b.Close($false); Start-Sleep 2 }; $f="$d\2DSoilFEM_$($tag)_20261021_$(Get-Date -f HHmm).xlsm"; Copy-Item $from $f; Copy-Item (Join-Path (Split-Path $from) 'soilfem_native_x64.dll') $d -Force -EA 0; $script:b=$x.Workbooks.Open($f); $x.EnableEvents=$false; L "open $f from $from"; return $f }
function DumpSheet($sheet,$file,$cols,$maxr=3000){ try{ $j=$script:b.Worksheets.Item($sheet); $lr=$j.Cells($j.Rows.Count,1).End(-4162).Row; $t=@(); for($r=1;$r -le [Math]::Min($lr,$maxr);$r++){ $row=@(); for($c=1;$c -le $cols;$c++){ $row+=[string]$j.Cells($r,$c).Value2 }; $t+=($row -join "`t") }; ($t -join "`n")|Out-File "$d\$file" -Encoding utf8; L "dump $sheet $file" }catch{ L "nodump $sheet $_" } }
function View($tag,$sel,$scale){ $ws=Go '出力・表示'; MS 'VIEW_COLOR_RESULT' $sel; MS 'VIEW_DISP_SCALE' $scale; $ws=Go '出力・表示'; Snap "$($tag)_out" @('C17','C18'); M 'FEMViewResult' "$tag-view"; Snap "$($tag)_fig" @() }
$steps=[ordered]@{}
$ex="$env:USERPROFILE\manual_rewrite_20261021\docs\manuals\examples"
function RunA($lab){ $x.Run('SetP0SilentMode',$true)|Out-Null; $x.Run('FEMWriteTextSetting','OUTPUT_STAGE_MODE','ALL')|Out-Null; $x.Run('FEMInvalidateInputCache')|Out-Null; $x.Run('FEMInvalidateSettingCache')|Out-Null; "1"|Out-File "$d\analyzing"; $t=[Diagnostics.Stopwatch]::StartNew(); M 'P0_RunAnalysis' $lab; $t.Stop(); Remove-Item "$d\analyzing" -EA 0; L "elapsed $lab $($t.Elapsed.TotalSeconds)" }
function JointDumps($tag){ $ws=Go 'ステージ結果'; Snap "$($tag)_result" @('A1:J3'); DumpSheet 'ステージ結果' "$($tag)_stage.tsv" 20; DumpSheet '診断' "$($tag)_diag.tsv" 5 400
  DumpSheet '結果要素' "$($tag)_elems.tsv" 72 2000; DumpSheet '要素データ' "$($tag)_eldata.tsv" 10 2000
  $ws=Go '操作パネル'; M 'FEMPracticalShowJointResults' "$tag-jres"; Snap "$($tag)_jres" @('G1:H25','I1:I25','L1:L25'); DumpSheet '接合結果' "$($tag)_joint.tsv" 14 200 }
$steps['gamma']={ $f=OpenF "$ex\ex08_road_result_20261021.xlsm" 'road_gamma20v2'
  $ws=Go '材料データ'; $ws.Range('E9').Value2=[double]20; Snap 'rg2_mat' @('A9:K9','E9')
  RunA 'rg2-run'; $ws=Go 'ステージ結果'; Snap 'rg2_result' @('A1:J12','H12'); DumpSheet 'ステージ結果' 'rg2_stage.tsv' 10; DumpSheet '結果節点' 'rg2_nodes.tsv' 10 20000; DumpSheet '診断' 'rg2_diag.tsv' 5 400
  $script:b.Save(); L "saved $f" }
$steps['taller']={ $f=OpenF "$ex\ex08_road_input_20261021.xlsm" 'road_tallerv2'
  $ws=Go '形状入力'; $ws.Range('D24').Value2=[double]5; $ws.Range('D25').Value2=[double]5; Snap 'rt2_shape' @('B24:D25','D24:D25')
  $ws=Go 'メッシュ設定'; M 'FEMUiGenerateShape' 'rt2-mesh'; Snap 'rt2_mesh' @()
  RunA 'rt2-run'; $ws=Go 'ステージ結果'; Snap 'rt2_result' @('A1:J12','H12'); DumpSheet 'ステージ結果' 'rt2_stage.tsv' 10; DumpSheet '結果節点' 'rt2_nodes.tsv' 10 20000; DumpSheet '診断' 'rt2_diag.tsv' 5 400
  $n=$script:b.Worksheets.Item('節点データ'); L ("nodes "+($n.Cells($n.Rows.Count,1).End(-4162).Row-1)); $e=$script:b.Worksheets.Item('要素データ'); L ("elems "+($e.Cells($e.Rows.Count,1).End(-4162).Row-1))
  View 'rt2_sy' 2 1; $script:b.Save(); L "saved $f" }
foreach($p in 4.68,4.72){ $steps["s$p"]=[scriptblock]::Create(@"
  `$tag='js_p$([int]($p*100))'; `$f=OpenF "`$ex\ex11b_joint_strongblock75_20261021_2009.xlsm" `$tag
  `$ws=Go '載荷'; `$ws.Range('E2').Value2=[double]$p; Snap "`$(`$tag)_load" @('A2:I2','E2'); RunA "`$tag-run"; JointDumps `$tag; `$script:b.Save(); L "saved `$f"
"@) }
foreach($p in 1.5,1.6){ $steps["w$p"]=[scriptblock]::Create(@"
  `$tag='jw_p$([int]($p*10))'; `$f=OpenF "`$ex\ex11_joint_20261021.xlsm" `$tag
  `$ws=Go '載荷'; `$ws.Range('A2:J2').Value2=[object[]]@(2,'BOX','X','FORCE',[double]$p,4,6,8,6,'ブロック上面に水平力')
  `$ws=Go 'ステージ'; `$ws.Range('A3:F3').Value2=[object[]]@(2,'LOAD',2,`$null,1,'水平力')
  Snap "`$(`$tag)_stage" @('A2:F3'); RunA "`$tag-run"; JointDumps `$tag; `$script:b.Save(); L "saved `$f"
"@) }
foreach($k in $steps.Keys){ try{ & $steps[$k] }catch{ L "EXERR $k $($_.Exception.Message) @ $($_.InvocationInfo.ScriptLineNumber)" } }
L "all done"
try{ $script:b.Close($false) }catch{}; $x.Quit(); Start-Sleep 5; if(Get-Process -Id $me -EA 0){ Stop-Process -Id $me -Force; L "killed own $me" }
"x"|Out-File "$d\stop"
