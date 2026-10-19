param([Parameter(Mandatory=$true)][string]$Compiler,[string]$SourceFile=(Join-Path $PSScriptRoot '../../native/soilfem_mat.c'),[Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
$Compiler=(Resolve-Path -LiteralPath $Compiler).Path
$SourceFile=(Resolve-Path -LiteralPath $SourceFile).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve existing build.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$output=Join-Path $OutputDirectory 'soilfem_mat_x64.dll'
$options=@('-O2','-std=c11','-Wall','-Wextra','-Werror','-Wno-unused-parameter','-Wno-misleading-indentation',
 '-mfpmath=387','-fexcess-precision=standard','-ffp-contract=off','-fno-fast-math',
 '-fno-builtin','-fno-stack-protector','-mno-stack-arg-probe','-nostdlib','-nostartfiles','-shared',
 '-Wl,--entry=DllMain','-Wl,--kill-at','-s','-o',$output,$SourceFile,'-lkernel32','-lgcc')
$originalPath=$env:PATH
try{
 $env:PATH=(Split-Path $Compiler)+';'+$originalPath
 & $Compiler @options
 if($LASTEXITCODE -ne 0){throw 'Material DLL compile failed.'}
}finally{$env:PATH=$originalPath}
@{compiler=(& $Compiler --version | Select-Object -First 1);source=$SourceFile;source_sha256=(Get-FileHash -LiteralPath $SourceFile).Hash;options=$options;dll_sha256=(Get-FileHash -LiteralPath $output).Hash}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $OutputDirectory 'build.json') -Encoding UTF8
Write-Output "PASS material DLL x87 / no CRT / warnings as errors: $output"
