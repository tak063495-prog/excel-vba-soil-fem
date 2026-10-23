param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out,[string]$Cases='')
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no dispatch.'}
if(Test-Path $Out){throw 'Preserve previous evidence.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory -Path $Out|Out-Null
if(-not $Cases){$Cases="$PSScriptRoot/fixtures/issue59/q8_jacobian_cases.json"}
$geometryCases=(Get-Content $Cases -Raw|ConvertFrom-Json).cases
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1
 $copy=Join-Path $Out 'geometry_probe.xlsm';Copy-Item $Book $copy
 $wb=$xl.Workbooks.Open($copy,0,$false)
 $wb.VBProject.VBComponents.Item('FEM').CodeModule.AddFromString(@'
Public Function Issue59Probe(ByVal raw As Variant, ByVal tol As Double) As Variant
 Dim x(0 To 7) As Double, y(0 To 7) As Double, i As Long, why As String, ok As Boolean
 Dim ed As Element_Data, runtimeOk As Boolean
 For i = 0 To 7
  x(i) = CDbl(raw(i, 0)): y(i) = CDbl(raw(i, 1))
  ed.x(i) = x(i): ed.y(i) = y(i): ed.node(i) = i
 Next i
 ok = FEMQ8GeometryCertified(x, y, tol, why)
 AnalysisMessage = "": AnalysisErrorNumber = 0
 runtimeOk = CheckElementJacobian(ed, 41, False)
 Issue59Probe = Array(ok, why, runtimeOk, AnalysisMessage, AnalysisErrorNumber)
End Function
'@)
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()}
 if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Probe compile failed'}
 $results=@()
 foreach($case in $geometryCases){
  $a=[object[,]]::new(8,2);$span=0.0
  for($i=0;$i -lt 8;$i++){for($j=0;$j -lt 2;$j++){$a[$i,$j]=[double]$case.coordinates[$i][$j];$span=[Math]::Max($span,[Math]::Abs($a[$i,$j]-[double]$case.coordinates[0][$j]))}}
  $tol=$span*$span*1e-12
  $r=$xl.Run("'"+$wb.Name+"'!Issue59Probe",$a,$tol)
  $accept=$case.name -in @('square','rotated_scaled','curved_valid','square_small','scaled_translated','curved_subdivision')
  if([bool]$r[0] -ne $accept){throw "Wrong certificate for $($case.name): $r"}
  if($accept -and $case.name -ne 'square_small' -and -not [bool]$r[2]){throw "Valid runtime geometry refused: $($case.name): $r"}
  if(-not $accept -and ([bool]$r[2] -or -not ([string]$r[3]).Contains('要素=42'))){throw "Runtime guard failed: $($case.name): $r"}
  $results+=@{case=$case.name;status='PASS';certificate=[bool]$r[0];reason=[string]$r[1];runtime=[bool]$r[2];runtime_message=[string]$r[3];tolerance=$tol}
 }
 @{status='PASS';book_sha256=(Get-FileHash $Book).Hash;cases=$results}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Out 'geometry.json') -Encoding utf8
 $wb.Close($false)
}finally{
 if($wb){try{$wb.Close($false)}catch{}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
