<#
  Issue #2: INCONSISTENT + 連続体DEATH/BIRTHで残る拘束済み孤立DOFをCSR組立が拒否しないこと
  Issue #3: 非対称帯域LU・GMRES・正則化残差回復でのユーザー中断（Err18）が既存中断処理へ伝播すること
  -SourceWorkbook のコピーへ -ModDir のモジュールを取込み（元ブックは変更しない）、DLLを隣に置いて OFF/AUTO で実行する。
  -ExpectFixed 指定時は受入条件を満たさなければ失敗終了する（未指定時は記録のみ。修正前ソースでの再現確認用）。
#>
param(
  [string]$SourceWorkbook=(Join-Path $PSScriptRoot '../workbook/2DSoilFEM_nativekern_20261017_0846.xlsm'),
  [string]$ModDir=(Join-Path $PSScriptRoot '../src/vba'),
  [string]$DllDir=(Join-Path $PSScriptRoot '../workbook'),
  [string]$Modes='OFF,AUTO',
  [string]$Label='after',
  [ValidateSet('ALL','2','3')][string]$Only='ALL',
  [switch]$ExpectFixed
)
$ErrorActionPreference='Stop'
$rootTask=Split-Path -Parent $PSScriptRoot
Push-Location $rootTask
$xlTask=$null;$wbTask=$null;$xpid=$null
try {
New-Item -ItemType Directory -Path tests/tmp,tests/results -Force|Out-Null
$srcTask=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$srcHash=(Get-FileHash -LiteralPath $srcTask -Algorithm SHA256).Hash
$workDir=Join-Path $pwd ('tests/tmp/issue23_'+$Label)
New-Item -ItemType Directory -Path $workDir -Force|Out-Null
$copyTask=Join-Path $workDir ('issue23_'+$Label+'.xlsm')
Copy-Item -LiteralPath $srcTask -Destination $copyTask -Force
foreach($d in 'soilfem_native_x64.dll','soilfem_mat_x64.dll'){ $p=Join-Path $DllDir $d; if(Test-Path -LiteralPath $p){Copy-Item -LiteralPath $p -Destination $workDir -Force} }
$failures=[Collections.Generic.List[string]]::new()
function Check($cond,$msg){ if(-not $cond){ $failures.Add($msg); Write-Host "CHECK_FAIL $msg" } }
function PutRows($wb,$name,$rows,$cols){
  $ws=$wb.Worksheets.Item($name);$ws.Range($ws.Cells(2,1),$ws.Cells(2000,$cols)).ClearContents()
  $data=[object[,]]::new($rows.Count,$cols)
  for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $cols;$c++){$data[$r,$c]=$rows[$r][$c]}}
  $ws.Range($ws.Cells(2,1),$ws.Cells($rows.Count+1,$cols)).Value2=$data
}
function Setting($wb,$key,$value){
  $method=$(if($value -is [string]){'FEMWriteTextSetting'}else{'FEMWriteNumericSetting'})
  $xlTask.Run("'"+$wb.Name+"'!"+$method,$key,$value)
}
function ReadSrc($m){ [IO.File]::ReadAllText((Join-Path $ModDir ($m+'.bas')),[Text.Encoding]::UTF8) }
function Hook([string]$text,[string]$old,[string]$new,[string]$name){
  $o=$old -replace "`r?`n","`r`n"; $n=$new -replace "`r?`n","`r`n"
  $idx=$text.IndexOf($o); if($idx -lt 0){throw "hook not found: $name"}
  if($text.IndexOf($o,$idx+1) -ge 0){throw "hook not unique: $name"}
  return $text.Substring(0,$idx)+$n+$text.Substring($idx+$o.Length)
}
$harness=@'
Public IssueInject As Long, IssueInjectAt As Long, IssueHitLU As Long, IssueHitGM As Long, IssueHitRR As Long
Private Function IssueState() As String
  IssueState = ResultStatus & "|" & CStr(AnalysisOK) & "|" & CStr(P3UserCancel) & "|" & CStr(AnalysisErrorNumber) & "|" & CStr(IssueHitLU) & "|" & CStr(IssueHitGM) & "|" & CStr(IssueHitRR) & "|" & P6CSRLastAssemblyPath & "|" & CStr(P6RCMMapReady) & "|" & FEMNativeStatusText() & "|" & AnalysisMessage
End Function
Public Function IssueRun(ByVal modeText As String, ByVal inject As Long, ByVal injectAt As Long) As String
  On Error GoTo Failed
  FEMNativeSetMode modeText
  FEMInvalidateSettingCache
  IssueInject = inject: IssueInjectAt = injectAt: IssueHitLU = 0: IssueHitGM = 0: IssueHitRR = 0
  P0_RunAnalysis
  IssueRun = IssueState()
  Exit Function
Failed:
  IssueRun = IssueState() & "|RAISED=" & CStr(Err.Number)
End Function
Public Function IssueOriginalDisp() As Variant
  Dim result() As Double, n As Long, k As Long
  ReDim result(2 * NumberOfFreeNode - 1)
  For n = 0 To NumberOfFreeNode - 1
    k = P6GetInternalFreeNode(n)
    result(2 * n) = TDisp(2 * k): result(2 * n + 1) = TDisp(2 * k + 1)
  Next n
  IssueOriginalDisp = result
End Function
Public Function IssueOriginalNodeCond() As Variant
  Dim result() As Double, n As Long, k As Long
  ReDim result(2 * NumberOfFreeNode - 1)
  For n = 0 To NumberOfFreeNode - 1
    k = P6GetInternalFreeNode(n)
    result(2 * n) = NodeCond(2 * k): result(2 * n + 1) = NodeCond(2 * k + 1)
  Next n
  IssueOriginalNodeCond = result
End Function
Public Function IssueFos() As String
  Dim ws As Worksheet, r As Long
  On Error GoTo Failed
  Set ws = ThisWorkbook.Worksheets("ステージ結果")
  For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    IssueFos = IssueFos & CStr(ws.Cells(r, 2).Value2) & ":" & CStr(ws.Cells(r, 3).Value2) & ":" & CStr(ws.Cells(r, 5).Value2) & ":" & CStr(ws.Cells(r, 10).Value2) & ";"
  Next r
  Exit Function
Failed:
  IssueFos = "ERR"
End Function
'@
$diagSelfTest=@'

Public Function IssueDiagSelfTest(ByVal badValue As Double) As String
  ' Test-only: pre-boundary diagonal acceptance on the current (post-DEATH) CSR.
  Dim v() As Double, r As Long, orphan As Long, pos As Long, s As String, keep As Long
  If Not P6CSRReady Then IssueDiagSelfTest = "NOCSR": Exit Function
  v = P6CSROriginalValues
  orphan = -1
  For r = 0 To lastDof
    If NodeCond(r) <> 0 Then
      pos = P6CSRDiagonalPosition(r)
      If pos >= 0 Then
        If Abs(v(pos)) <= 1E-30 Then orphan = r: Exit For
      End If
    End If
  Next r
  If orphan < 0 Then IssueDiagSelfTest = "NOORPHAN": Exit Function
  s = "orphan=" & CStr(orphan)
  s = s & "|allow=" & CStr(P6CheckCSRDiagonalValues(v, True))
  s = s & "|strict=" & CStr(P6CheckCSRDiagonalValues(v, False))
  v(pos) = badValue
  s = s & "|nonfinite=" & CStr(P6CheckCSRDiagonalValues(v, True))
  v(pos) = 0#
  keep = NodeCond(orphan): NodeCond(orphan) = 0
  s = s & "|freezero=" & CStr(P6CheckCSRDiagonalValues(v, True))
  NodeCond(orphan) = keep
  IssueDiagSelfTest = s
End Function
'@

# run.ps1内の直前の試験が解放前のCOM参照でExcelを残すことがあるため、GCで参照を解放し最大60秒だけ終了を待つ（他のExcelは停止しない）
$waitUntil=(Get-Date).AddSeconds(60)
do { [GC]::Collect(); [GC]::WaitForPendingFinalizers(); $before=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object{$_.Id}); if($before.Count -eq 0){break}; Start-Sleep 2 } while((Get-Date) -lt $waitUntil)
if($before.Count -gt 0){throw "Another Excel is running: $before"}
$xlTask=New-Object -ComObject Excel.Application
$xpid=@(Get-Process EXCEL|ForEach-Object{$_.Id}|Where-Object{$before -notcontains $_})[0]
Set-Content -LiteralPath "$copyTask.running" -Value "pid=$xpid"
$xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1
$wbTask=$xlTask.Workbooks.Open($copyTask,0,$false)
$solverText=ReadSrc 'FEMSolver'
$fixedSource=$solverText.Contains('allowConstrainedZero')
$solverText=Hook $solverText "  startedAt = Timer`n  If reused Then`n" "  startedAt = Timer`n  IssueHitLU = IssueHitLU + 1`n  If IssueInject = 1 And IssueHitLU = IssueInjectAt Then Err.Raise 18`n  If IssueInject >= 2 Then P6BandRejectReason = `"INJECT`": GoTo Rejected`n  If reused Then`n" 'nonsym LU'
$solverText=Hook $solverText "  On Error GoTo GMRESFailed`n" "  On Error GoTo GMRESFailed`n  IssueHitGM = IssueHitGM + 1`n  If IssueInject = 2 And IssueHitGM = IssueInjectAt Then Err.Raise 18`n  If IssueInject = 3 Or IssueInject = 5 Then Exit Function`n" 'gmres'
$solverText=Hook $solverText "  startedAt = Timer`n  P6FactorizationCount = P6FactorizationCount + 1`n  kRr =" "  startedAt = Timer`n  IssueHitRR = IssueHitRR + 1`n  If IssueInject = 3 And IssueHitRR = IssueInjectAt Then Err.Raise 18`n  P6FactorizationCount = P6FactorizationCount + 1`n  kRr =" 'residual recovery'
if($fixedSource){ $solverText=$solverText + ($diagSelfTest -replace "`r?`n","`r`n") }
foreach($m in 'CONTROL','FEMCore','FEMSolver','FEMIo','FEMEngine','FEMNative','FEMNativeMat','FEMNativeKern'){
  $c=$null; try{$c=$wbTask.VBProject.VBComponents.Item($m)}catch{}
  if($null -eq $c){$c=$wbTask.VBProject.VBComponents.Add(1);$c.Name=$m}
  $cm=$c.CodeModule; if($cm.CountOfLines -gt 0){$cm.DeleteLines(1,$cm.CountOfLines)}
  $cm.AddFromString($(if($m -eq 'FEMSolver'){$solverText}else{ReadSrc $m}))
}
$h=$wbTask.VBProject.VBComponents.Add(1);$h.Name='IssueHarness';$h.CodeModule.AddFromString($harness)
$ctl=$xlTask.VBE.CommandBars.FindControl(1,578); if($ctl.Enabled){$ctl.Execute()}
if($xlTask.VBE.CommandBars.FindControl(1,578).Enabled){throw 'VBA compile failed'}
Write-Output "source=$srcTask sha256=$srcHash fixedSource=$fixedSource"

# ---- 2要素Q8モデル（幅1×高さ2、下端固定・側面ローラー）。要素1=下（材料1）、要素2=上（材料2）
$nodesAll=@(
 @(1,0.0,0.0,1,1),@(2,0.5,0.0,1,1),@(3,1.0,0.0,1,1),@(4,0.0,0.5,1,0),@(5,1.0,0.5,1,0),
 @(6,0.0,1.0,1,0),@(7,0.5,1.0,0,0),@(8,1.0,1.0,1,0),@(9,0.0,1.5,1,0),@(10,1.0,1.5,1,0),
 @(11,0.0,2.0,1,0),@(12,0.5,2.0,0,0),@(13,1.0,2.0,1,0))
function NodeRows($list){ ,@($list|ForEach-Object{ ,@([double]$_[0],[double]$_[1],[double]$_[2],[double]$_[3],[double]$_[4],$null,$null,$null,$null) }) }
$elemLower=@(1.0,1.0,3.0,8.0,6.0,2.0,5.0,7.0,4.0,1.0)
$elemUpper=@(2.0,6.0,8.0,13.0,11.0,7.0,10.0,12.0,9.0,2.0)
$orphanNodes=@(9,10,11,12,13)
function SetupModel($wb,$variant){
  foreach($k in 'ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V2B_COST','ACCEL_V3_COST_SEARCH','ACCEL_V4_ANDERSON','ACCEL_V5_GMRES_LU','ACCEL_STEP_RECOVERY','ACCEL_ADAPTIVE'){Setting $wb $k 0}
  Setting $wb 'FLOW_POLICY' 'INCONSISTENT'; Setting $wb 'CONSOL_ENABLE' 0
  Setting $wb 'EXPORT_MODE' 'OFF'; Setting $wb 'EXPORT_LOAD' ''; Setting $wb 'DEBUG_MODE' 'OFF'
  Setting $wb 'OUTPUT_STAGE_MODE' 'ALL'; Setting $wb 'OUTPUT_INACTIVE_ELEMENTS' 'SKIP'
  Setting $wb 'RCM_POLICY' $(if($variant.Rcm){'ON'}else{'OFF'})
  Setting $wb 'SRM_FIXED_FS' $(if($variant.Srm){1.0}else{0})
  $wb.Worksheets.Item('接合').Range('A2:E2000').ClearContents()
  $wb.Worksheets.Item('載荷').Range('A2:J2000').ClearContents()
  $mats=@(,@(1.0,14000.0,0.3,1.0,20.0,30.0,100.0,0.0,1.0,'ON','SOIL',$null,$null))
  $mats+=,@(2.0,14000.0,0.3,1.0,20.0,30.0,100.0,0.0,1.0,$(if($variant.UpperOff){'OFF'}else{'ON'}),'SOIL',$null,$null)
  PutRows $wb '材料データ' $mats 13
  if($variant.LowerOnly){
    PutRows $wb '節点データ' (NodeRows ($nodesAll|Where-Object{$_[0] -le 8})) 9
    PutRows $wb '要素データ' (,$elemLower) 10
  } else {
    PutRows $wb '節点データ' (NodeRows $nodesAll) 9
    PutRows $wb '要素データ' @($elemLower,$elemUpper) 10
  }
  PutRows $wb 'ステージ' $variant.Stages 6
}
$stGravity=,@(1.0,'GRAVITY',$null,$null,1.0,'自重')
$stDeath=@(@(1.0,'GRAVITY',$null,$null,1.0,'自重'),@(2.0,'DEATH',2.0,$null,1.0,'上側撤去'),@(3.0,'GRAVITY',$null,$null,1.0,'撤去後'))
$stBirth=@(@(1.0,'GRAVITY',$null,$null,1.0,'自重'),@(2.0,'BIRTH',2.0,$null,1.0,'上側追加'),@(3.0,'GRAVITY',$null,$null,1.0,'追加後'))
$stSrm=@(@(1.0,'GRAVITY',$null,$null,1.0,'自重'),@(2.0,'SRM',$null,$null,1.0,'Fs固定検証'))
$results=[Collections.Generic.List[object]]::new()
function RunCase($name,$variant,$mode,$inject=0,$at=0){
  SetupModel $wbTask $variant
  $r=$xlTask.Run("'"+$wbTask.Name+"'!IssueRun",$mode,[int]$inject,[int]$at)
  $d=@($xlTask.Run("'"+$wbTask.Name+"'!IssueOriginalDisp"))
  $nc=@($xlTask.Run("'"+$wbTask.Name+"'!IssueOriginalNodeCond"))
  $fos=$xlTask.Run("'"+$wbTask.Name+"'!IssueFos")
  $diag=''
  if($variant.DiagSelf -and $fixedSource){ $diag=$xlTask.Run("'"+$wbTask.Name+"'!IssueDiagSelfTest",[double]::PositiveInfinity) }
  $f=$r.Split('|')
  $o=[PSCustomObject]@{Case=$name;Mode=$mode;Inject=$inject;At=$at;Status=$f[0];OK=$f[1];Cancel=$f[2];ErrNo=$f[3];HitLU=[int]$f[4];HitGM=[int]$f[5];HitRR=[int]$f[6];Path=$f[7];Rcm=$f[8];Native=$f[9];Message=($f[10..($f.Count-1)] -join '|');Fos=$fos;Diag=$diag;Disp=$d;NodeCond=$nc}
  $results.Add($o)
  Write-Host ("{0} {1} inject={2}@{3} -> {4}|{5} cancel={6} err={7} LU={8} GM={9} RR={10} path={11} diag={12} fos={13} msg={14}" -f $name,$mode,$inject,$at,$o.Status,$o.OK,$o.Cancel,$o.ErrNo,$o.HitLU,$o.HitGM,$o.HitRR,$o.Path,$diag,$fos,$o.Message)
  return $o
}
function Idx($node,$comp){ 2*($node-1)+$comp }
$modeList=$Modes.Split(',')

# ===== Issue #2 =====
if($Only -ne '3'){
$v2=@(
 @{Name='death_rcm_off';Stages=$stDeath;Rcm=$false;DiagSelf=$true},
 @{Name='death_rcm_on';Stages=$stDeath;Rcm=$true;DiagSelf=$true},
 @{Name='birth_rcm_off';Stages=$stBirth;Rcm=$false;UpperOff=$true},
 @{Name='birth_rcm_on';Stages=$stBirth;Rcm=$true;UpperOff=$true})
foreach($mode in $modeList){
  $g=RunCase 'gravity_both' @{Stages=$stGravity} $mode
  $lo=RunCase 'lower_only' @{Stages=$stGravity;LowerOnly=$true} $mode
  $res=@{}
  foreach($v in $v2){ $res[$v.Name]=RunCase $v.Name $v $mode }
  foreach($k in $res.Keys){ $o=$res[$k]; Check ($o.Status -eq 'PASS' -and $o.OK -eq 'True') "#2 $k $mode not PASS: $($o.Status) $($o.Message)" }
  foreach($k in 'death_rcm_off','death_rcm_on'){
    $o=$res[$k]
    foreach($n in $orphanNodes){ foreach($c in 0,1){
      $i=Idx $n $c
      Check ([double]$o.NodeCond[$i] -ne 0) "#2 $k $mode orphan node $n dof $c not constrained"
      Check ([math]::Abs([double]$o.Disp[$i]-[double]$g.Disp[$i]) -le 1e-12*(1+[math]::Abs([double]$g.Disp[$i]))) "#2 $k $mode orphan node $n dof $c held disp $($o.Disp[$i]) != stage1 $($g.Disp[$i])"
    } }
    $maxd=0.0;$maxr=0.0
    for($i=0;$i -lt 16;$i++){ $maxd=[math]::Max($maxd,[math]::Abs([double]$o.Disp[$i]-[double]$lo.Disp[$i])); $maxr=[math]::Max($maxr,[math]::Abs([double]$lo.Disp[$i])) }
    Write-Output ("#2 {0} {1} lower-element vs lower-only reference: max|diff|={2:E3} max|ref|={3:E3}" -f $k,$mode,$maxd,$maxr)
    Check ($maxr -gt 0 -and $maxd -le 1e-9*$maxr) "#2 $k $mode remaining element response differs from lower-only reference ($maxd)"
    if($fixedSource){ Check ($o.Diag -match '\|allow=True\|strict=False\|nonfinite=False\|freezero=False$') "#2 $k $mode diag self-test: $($o.Diag)" }
  }
  foreach($k in 'birth_rcm_off','birth_rcm_on'){
    $o=$res[$k]
    for($i=0;$i -lt 26;$i++){ $n=[math]::Floor($i/2)+1; $c=$i%2; $sup=$nodesAll[$n-1][3+$c]; Check ([double]$o.NodeCond[$i] -eq [double]$sup) "#2 $k $mode node $n dof $c NodeCond $($o.NodeCond[$i]) != support $sup after BIRTH" }
    Check ([math]::Abs([double]$o.Disp[(Idx 12 1)]) -gt 0) "#2 $k $mode born top node did not move"
  }
  foreach($p in @(@('death_rcm_off','death_rcm_on'),@('birth_rcm_off','birth_rcm_on'))){
    $a=$res[$p[0]];$b=$res[$p[1]];$mx=0.0
    for($i=0;$i -lt $a.Disp.Count;$i++){ $mx=[math]::Max($mx,[math]::Abs([double]$a.Disp[$i]-[double]$b.Disp[$i])) }
    Check ($mx -le 1e-10) "#2 $($p[0]) vs $($p[1]) $mode RCM differs by $mx"
  }
  # 未拘束機構は引き続き拒否
  $nodesFree=@($nodesAll|ForEach-Object{ ,@($_[0],$_[1],$_[2],0,0) })
  SetupModel $wbTask @{Stages=$stGravity}
  PutRows $wbTask '節点データ' (NodeRows $nodesFree) 9
  $r=$xlTask.Run("'"+$wbTask.Name+"'!IssueRun",$mode,0,0)
  Write-Output "unsupported_mechanism $mode -> $r"
  $results.Add([PSCustomObject]@{Case='unsupported_mechanism';Mode=$mode;Raw=$r})
  Check (-not $r.StartsWith('PASS|')) "#2 unsupported mechanism $mode was not rejected"
}

}
# ===== Issue #3 =====
if($Only -ne '2'){
$vS=@{Stages=$stSrm;Srm=$true}
foreach($mode in $modeList){
  $base=RunCase 'srm_base' $vS $mode 0 0
  Check ($base.Status -eq 'PASS' -and $base.HitLU -gt 1 -and $base.HitGM -eq 0) "#3 baseline $mode unexpected: $($base.Status) LU=$($base.HitLU) GM=$($base.HitGM)"
  $ref4=RunCase 'lu_reject_gmres' $vS $mode 4 0
  Check ($ref4.Status -eq 'PASS' -and $ref4.Cancel -eq 'False' -and $ref4.HitGM -eq $ref4.HitLU) "#3 normal LU failure -> GMRES $mode changed: $($ref4.Status) cancel=$($ref4.Cancel) LU=$($ref4.HitLU) GM=$($ref4.HitGM)"
  $ref5=RunCase 'lu_gmres_reject_rr' $vS $mode 5 0
  Check ($ref5.Cancel -eq 'False' -and $ref5.ErrNo -ne '18' -and $ref5.HitRR -ge 1) "#3 normal LU/GMRES failure -> residual recovery $mode changed: $($ref5.Status) cancel=$($ref5.Cancel) RR=$($ref5.HitRR)"
  $cases=@(
    @{Name='cancel_lu';Inject=1;At=$base.HitLU},
    @{Name='cancel_gmres';Inject=2;At=$ref4.HitGM},
    @{Name='cancel_rr';Inject=3;At=[math]::Max(1,$ref5.HitRR)})
  foreach($c in $cases){
    $o=RunCase $c.Name $vS $mode $c.Inject $c.At
    Check ($o.Status -ne 'PASS' -and $o.OK -eq 'False') "#3 $($c.Name) $mode published PASS"
    Check ($o.Cancel -eq 'True' -and $o.ErrNo -eq '18') "#3 $($c.Name) $mode cancel not propagated: cancel=$($o.Cancel) err=$($o.ErrNo)"
    Check ($o.Fos -notmatch 'FIXED_FS_VERIFIED' -and $o.Fos -notmatch 'FOS_PASS=[0-9]') "#3 $($c.Name) $mode published Fs: $($o.Fos)"
    switch($c.Inject){
      1 { Check ($o.HitLU -eq $c.At -and $o.HitGM -eq 0) "#3 cancel_lu $mode continued: LU=$($o.HitLU) GM=$($o.HitGM)" }
      2 { Check ($o.HitGM -eq $c.At -and $o.HitLU -eq $c.At -and $o.HitRR -eq 0) "#3 cancel_gmres $mode continued: LU=$($o.HitLU) GM=$($o.HitGM) RR=$($o.HitRR)" }
      3 { Check ($o.HitRR -eq $c.At) "#3 cancel_rr $mode continued damping/solves: RR=$($o.HitRR)" }
    }
  }
}
}
# OFF/AUTO bit identity for every case
if($modeList.Count -ge 2){
  foreach($g in ($results|Where-Object{$_.Disp}|Group-Object Case,Inject)){
    $a=$g.Group|Where-Object{$_.Mode -eq $modeList[0]}|Select-Object -First 1;$b=$g.Group|Where-Object{$_.Mode -eq $modeList[1]}|Select-Object -First 1
    if($a -and $b){
      $same=($a.Disp.Count -eq $b.Disp.Count); if($same){for($i=0;$i -lt $a.Disp.Count;$i++){ if([double]$a.Disp[$i] -ne [double]$b.Disp[$i]){$same=$false;break} }}
      Check ($same -and $a.Status -eq $b.Status -and $a.HitLU -eq $b.HitLU -and $a.HitGM -eq $b.HitGM -and $a.HitRR -eq $b.HitRR) "$($g.Name) $($modeList[0]) vs $($modeList[1]) not bit-identical"
      Write-Output ("bitident {0}: {1}" -f $g.Name,$same)
    }
  }
}
$wbTask.Close($false); $wbTask=$null
if((Get-FileHash -LiteralPath $srcTask -Algorithm SHA256).Hash -ne $srcHash){throw 'Source workbook changed'}
$results|ConvertTo-Json -Depth 5|Set-Content -LiteralPath ('tests/results/issue2_3_'+$Label+'.json') -Encoding UTF8
Write-Output ("Issue #2/#3 regression: {0} check failures (ExpectFixed={1})" -f $failures.Count,[bool]$ExpectFixed)
if($ExpectFixed -and $failures.Count -gt 0){ throw ("Issue #2/#3 regression failed: "+($failures -join '; ')) }
} finally {
  if($null -ne $wbTask){try{$wbTask.Close($false)}catch{}}
  if($null -ne $xlTask){ try{$xlTask.Quit()}catch{}; [Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)|Out-Null; Start-Sleep 3
    if($xpid -and (Get-Process -Id $xpid -ErrorAction SilentlyContinue)){Stop-Process -Id $xpid -Force} }
  if($copyTask){Remove-Item -LiteralPath "$copyTask.running" -ErrorAction SilentlyContinue}
  Pop-Location
}
