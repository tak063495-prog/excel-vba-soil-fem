param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out,[string]$Cases='tests/fixtures/material_issue54_55.json')
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running'}
if(Test-Path $Out){throw 'Preserve evidence'}
$Book=(Resolve-Path $Book).Path;$hash=(Get-FileHash $Book).Hash;$Cases=(Resolve-Path $Cases).Path;$Out=[IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory $Out|Out-Null;Copy-Item $Book (Join-Path $Out 'material_probe.xlsm')
$xl=New-Object -ComObject Excel.Application
try{
 $xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open((Join-Path $Out 'material_probe.xlsm'),0,$false)
 $engine=$wb.VBProject.VBComponents.Item('FEMEngine').CodeModule
 # Test-only entry points expose the private cone in the same module;
 # production material update and its tolerances remain unchanged.
 $engine.AddFromString(@'
Public Function IMCone(ByVal psi As Double, ByVal ex As Double, ByVal ey As Double, ByVal gamma As Double, ByVal ez As Double) As Variant
 Dim ok As Boolean, total As Double, s As Double
 s = Sin(psi * 3.14159265358979 / 180#)
 ok = P2PlasticStrainInVertexCone(s, ex, ey, gamma, ez, total)
 IMCone = Array(ok, total, P2StrictVertexConeInterior(s, ex, ey, gamma, ez, 0.000000001, 1#, 1000#))
End Function
Public Function IMPoint(ByVal psi As Double, ByVal cohesion As Double, ByVal ex As Double, ByVal ey As Double, ByVal gamma As Double, ByVal sx As Double, ByVal sy As Double, ByVal tau As Double, ByVal sz As Double, ByVal steps As Long) As Variant
 Dim ip As P2_MaterialPointInput, op As P2_MaterialPointOutput, ok As Boolean, i As Long, j As Long, mx As Double
 ip.young = 1000#: ip.Poisson = 0.3: ip.frictionAngle = 30#: ip.dilationAngle = psi: ip.cohesion = cohesion
 ip.StrainIncrement(0) = ex: ip.StrainIncrement(1) = ey: ip.StrainIncrement(2) = gamma
 ip.PreviousStress(0) = sx: ip.PreviousStress(1) = sy: ip.PreviousStress(2) = tau: ip.PreviousStress(3) = sz
 ip.tolerance = 0.000000001: ip.maxIterations = 50: ip.maxSubsteps = steps: ip.PreferredSubsteps = steps: ip.EnableSubstepping = (steps > 1)
 ok = P2MaterialPointUpdate(ip, op)
 For i = 0 To 2: For j = 0 To 2: If Abs(op.tangent(i,j)) > mx Then mx = Abs(op.tangent(i,j))
 Next j, i
 IMPoint = Array(ok, op.converged, op.Elastic, op.failureCode, op.Stress(0), op.Stress(1), op.Stress(2), op.Stress(3), op.PlasticStrain(0), op.PlasticStrain(1), op.PlasticStrain(2), op.PlasticStrain(3), op.PlasticMultiplier, op.AlgorithmicTangentReady, mx)
End Function
'@)
 $records=[Collections.Generic.List[object]]::new();$prefix="'"+$wb.Name+"'!"
 foreach($c in (Get-Content $Cases -Raw -Encoding utf8|ConvertFrom-Json).cases){
  if($c.mode -eq 'cone'){$v=$xl.Run($prefix+'IMCone',[double]$c.psi,[double]$c.plastic[0],[double]$c.plastic[1],[double]$c.plastic[2],[double]$c.plastic[3])}
  else{$steps=1;if($c.substeps){$steps=[int]$c.substeps};$v=$xl.Run($prefix+'IMPoint',[double]$c.psi,[double]$c.c,[double]$c.strain[0],[double]$c.strain[1],[double]$c.strain[2],[double]$c.previous[0],[double]$c.previous[1],[double]$c.previous[2],[double]$c.previous[3],$steps)}
  $records.Add(@{id=$c.id;input=$c;actual=@($v)})
 }
 @{status='EXECUTED';book_sha256=$hash;fixture_sha256=(Get-FileHash $Cases).Hash;excel_version=$xl.Version;excel_build=$xl.Build;records=@($records)}|ConvertTo-Json -Depth 12|Set-Content (Join-Path $Out 'material_points.json') -Encoding utf8
 $wb.Close($false)
}finally{if($wb){try{$wb.Close($false)}catch{}};$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);[GC]::Collect();[GC]::WaitForPendingFinalizers()}
if((Get-FileHash $Book).Hash -ne $hash){throw 'Source changed'}
