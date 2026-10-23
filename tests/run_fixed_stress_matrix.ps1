param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no test started.'}
$Book=(Resolve-Path -LiteralPath $Book).Path
if(Test-Path -LiteralPath $Out){throw 'Preserve previous evidence.'}
$Out=[IO.Path]::GetFullPath($Out);New-Item -ItemType Directory -Path $Out | Out-Null
$xl=New-Object -ComObject Excel.Application
try {
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($Book,0,$true)
 $mod=$wb.VBProject.VBComponents.Add(1);$mod.Name='FSMatrixDriver'
 $mod.CodeModule.AddFromString([IO.File]::ReadAllText("$PSScriptRoot/fixed_stress_matrix_driver.bas"))
 $output=[ordered]@{}
 foreach($triangle in $false,$true){
   $key=if($triangle){'t6'}else{'q8'}
   $output[$key]=@($xl.Run("'"+$wb.Name+"'!FSMatrixDump",$triangle))
 }
 foreach($fixed in $false,$true){
   $key=if($fixed){'dirichlet'}else{'closed'}
   $output[$key]=@($xl.Run("'"+$wb.Name+"'!FSFlowDump",$fixed))
 }
 $output | ConvertTo-Json -Depth 8 | Set-Content -Encoding utf8 (Join-Path $Out 'actual.json')
 $wb.Close($false)
} finally {
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
'PASS'
