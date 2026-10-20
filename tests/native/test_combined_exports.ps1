param([Parameter(Mandatory=$true)][string]$Dll)
$ErrorActionPreference='Stop'
$path=(Resolve-Path -LiteralPath $Dll).Path
$tools=@('dumpbin','llvm-objdump','x86_64-w64-mingw32-objdump')
$tool=$null
foreach($candidate in $tools){$cmd=Get-Command $candidate -ErrorAction SilentlyContinue;if($cmd){$tool=$cmd.Source;break}}
if(-not $tool){
 $bundled=Join-Path $PSScriptRoot '../tmp/toolchain14/extracted/w64devkit/bin/x86_64-w64-mingw32-objdump.exe'
 if(Test-Path -LiteralPath $bundled){$tool=(Resolve-Path -LiteralPath $bundled).Path}
}
if(-not $tool){throw 'No export inspection tool found (dumpbin, llvm-objdump, or x86_64-w64-mingw32-objdump).'}
$raw=& $tool $(if($tool -like '*dumpbin*'){'/exports'}else{'-p'}) $path 2>&1 | Out-String
$required=@('sfn_abi','sfn_set_precision','sfn_get_precision','sfn_band_factor','sfn_band_solve','sfn_nsband_factor','sfn_nsband_forward','sfn_nsband_back','sfn_csr_scatter_list','sfn_gmres','sfn_gmres_tuned','sfn_gmres_prepared','sfn_restore_committed','sfm_init','sfm_eval','sfm_tangent','sfm_band_scatter','sfm_csr_scatter','sfm_csr_resnorm','sfm_ilu0','sfm_gmres','sfm_regularized','sfm_residual','sfm_rcm','sfm_hg_apply')
$missing=@($required|Where-Object{$raw -notmatch "\b$([regex]::Escape($_))\b"})
if($missing.Count){throw ('Missing exports: '+($missing -join ', '))}
Write-Output ('PASS combined export/ABI symbol inventory: '+$path)
