param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop';New-Item -ItemType Directory -Force -Path $OutputRoot|Out-Null
$hash=(Get-FileHash -LiteralPath $SourceWorkbook).Hash
$tmp=Join-Path ([IO.Path]::GetTempPath()) ('hg_force_'+[guid]::NewGuid().ToString('N')+'.xlsm');Copy-Item -LiteralPath $SourceWorkbook -Destination $tmp
$xl=New-Object -ComObject Excel.Application;$xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1;$wb=$null
try{
 $wb=$xl.Workbooks.Open($tmp,0,$false);$cm=$wb.VBProject.VBComponents.Item('FEM').CodeModule
 $cm.AddFromString(@'
Public Function ForcePolicyTest(ByVal flow As Long, ByVal srm As Boolean) As Variant
  Dim B() As Double, dj() As Double, sigma() As Double, j As Long
  SetP0SilentMode True
  ReDim Elem(0): ReDim Material(0): ReDim TDisp(1): ReDim iNForce(1)
  ReDim B(2,15,3): ReDim dj(3): ReDim sigma(16,3)
  Material(0).thickness=1#: Elem(0).MatNo=0
  For j=0 To 15: Elem(0).ElNode(j)=-1: Next j
  Elem(0).ElNode(0)=0: Elem(0).ElNode(1)=1
  Elem(0).HourglassModeCount=1: Elem(0).HourglassScale=10#
  Elem(0).HourglassMode(0,0)=1#: Elem(0).HourglassMode(0,1)=-1#
  TDisp(0)=1#: TDisp(1)=0.25
  B(0,0,0)=2#: B(0,1,0)=-2#: dj(0)=1#: sigma(0,0)=3#
  P3SrmEnabled=srm: P3PolicyCacheReady=True: P3FlowPolicyMode=flow
  CalcForce 0,B,dj,sigma
  ForcePolicyTest=Array(iNForce(0),iNForce(1))
End Function
'@)
 $records=@()
 foreach($case in @(@{flow=0;srm=$true;expected=6},@{flow=0;srm=$false;expected=13.5},@{flow=1;srm=$true;expected=13.5},@{flow=1;srm=$false;expected=13.5})){
  $r=@($xl.Run("'"+$wb.Name+"'!ForcePolicyTest",$case.flow,$case.srm))
  if([math]::Abs($r[0]-$case.expected)-gt 1e-12 -or [math]::Abs($r[1]+$case.expected)-gt 1e-12){throw ('Force policy failed: '+($r -join '|'))}
  $records+=@{flow=$case.flow;srm=$case.srm;expected=$case.expected;actual=$r}
 }
 @{status='PASS';source_sha256=$hash;records=$records}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $OutputRoot 'check_hourglass_force_policy.json') -Encoding UTF8
 Write-Output 'PASS: INCO SRM retains material force and excludes stabilization force; ordinary/Davis controls retain prior force'
}finally{if($wb){$wb.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)};$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue}
if((Get-FileHash -LiteralPath $SourceWorkbook).Hash-ne $hash){throw 'Source changed'}
