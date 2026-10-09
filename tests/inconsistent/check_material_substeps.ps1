param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot,[string]$PythonPath="python")
$ErrorActionPreference='Stop'; New-Item -ItemType Directory -Force $OutputRoot|Out-Null
$tmp=Join-Path ([IO.Path]::GetTempPath()) ('substep_'+[guid]::NewGuid().ToString('N')+'.xlsm'); Copy-Item $SourceWorkbook $tmp -Force
$before=(Get-FileHash $SourceWorkbook -Algorithm SHA256).Hash; $xl=New-Object -ComObject Excel.Application; $xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1;$wb=$null
try {
 $wb=$xl.Workbooks.Open($tmp,0,$false); $cm=$wb.VBProject.VBComponents.Add(1);$cm.Name='SpectralHarness';$cm.CodeModule.AddFromString(@'
Public Function SpectralProbe(ByVal phi As Double, ByVal psi As Double, ByVal sx As Double, ByVal sy As Double, ByVal tau As Double, ByVal sz As Double, ByVal ex As Double, ByVal ey As Double, ByVal gxy As Double, ByVal steps As Long) As Variant
 Dim i As P2_MaterialPointInput, o As P2_MaterialPointOutput, ok As Boolean
 i.Young=1000#:i.Poisson=.3:i.cohesion=1#:i.frictionAngle=phi:i.dilationAngle=psi:i.tolerance=1E-10:i.maxIterations=100:i.maxSubsteps=steps:i.PreferredSubsteps=steps:i.EnableSubstepping=True
 i.PreviousStress(0)=sx:i.PreviousStress(1)=sy:i.PreviousStress(2)=tau:i.PreviousStress(3)=sz:i.StrainIncrement(0)=ex:i.StrainIncrement(1)=ey:i.StrainIncrement(2)=gxy
 ok=P2MaterialPointUpdate(i,o):SpectralProbe=Array(ok,o.converged,o.UsedSubsteps,o.Stress(0),o.Stress(1),o.Stress(2),o.Stress(3),o.PlasticStrain(0),o.PlasticStrain(1),o.PlasticStrain(2),o.PlasticStrain(3),o.Tangent(0,0),o.Tangent(0,1),o.Tangent(0,2),o.Tangent(1,0),o.Tangent(1,1),o.Tangent(1,2),o.Tangent(2,0),o.Tangent(2,1),o.Tangent(2,2))
End Function
'@); $xl.Run("'"+$wb.Name+"'!SetP0SilentMode",$true)|Out-Null
 $casePath=Join-Path $OutputRoot 'substep_cases.json'
 & $PythonPath (Join-Path $PSScriptRoot 'substep_reference.py') --generate $casePath
 if($LASTEXITCODE -ne 0){throw 'Oracle fixture generation failed'}
 $cases=Get-Content $casePath -Raw -Encoding UTF8|ConvertFrom-Json
 $raw=@();foreach($c in $cases){$v=@($xl.Run("'"+$wb.Name+"'!SpectralProbe",$c.phi,$c.psi,$c.prev[0],$c.prev[1],$c.prev[2],$c.prev[3],$c.strain[0],$c.strain[1],$c.strain[2],$c.steps));$raw+=@{case=$c;v=$v}}
 [IO.File]::WriteAllText((Join-Path $OutputRoot 'substep_points.json'),($raw|ConvertTo-Json -Depth 8))
} finally {if($wb){$wb.Close($false);[Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)|Out-Null};$xl.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)|Out-Null;Remove-Item $tmp -Force -ErrorAction SilentlyContinue}
if((Get-FileHash $SourceWorkbook -Algorithm SHA256).Hash -ne $before){throw 'SourceWorkbook changed'}

& $PythonPath (Join-Path $PSScriptRoot 'substep_reference.py') --analyze (Join-Path $OutputRoot 'substep_points.json') --output (Join-Path $OutputRoot 'substep_checks.json')
if($LASTEXITCODE -ne 0){throw 'Composite substep oracle comparison failed'}
$reportTask=Get-Content (Join-Path $OutputRoot 'substep_checks.json') -Raw -Encoding UTF8|ConvertFrom-Json
$reportTask|Add-Member -NotePropertyName source_sha256 -NotePropertyValue $before
$reportTask|ConvertTo-Json -Depth 10|Set-Content (Join-Path $OutputRoot 'substep_checks.json') -Encoding UTF8
