param(
 [string]$SourceWorkbook=(Join-Path $PSScriptRoot '../../workbook/inco_ilumat_20261020/2DSoilFEM_inco_ilumat_20261020.xlsm'),
 [Parameter(Mandatory=$true)][string]$OutputDirectory,
 [string]$MaterialDll=(Join-Path $PSScriptRoot '../../native/soilfem_mat_x64.dll'),
 [string]$NativeDll=(Join-Path $PSScriptRoot '../../workbook/inco_ilumat_20261020/soilfem_native_x64.dll'),
 [Parameter(Mandatory=$true)][string]$MaterialBuildRecord
)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$SourceWorkbook=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve existing package.'}
$sourceHash=(Get-FileHash -LiteralPath $SourceWorkbook).Hash
$materialBuild=Get-Content -LiteralPath $MaterialBuildRecord -Raw|ConvertFrom-Json
if($materialBuild.dll_sha256 -cne (Get-FileHash -LiteralPath $MaterialDll).Hash){throw 'Material build/DLL mismatch.'}
if($materialBuild.source_sha256 -cne (Get-FileHash -LiteralPath (Join-Path $repo 'native/soilfem_mat.c')).Hash){throw 'Material source/build mismatch.'}
foreach($option in '-O3','-mfpmath=387','-fexcess-precision=standard','-ffp-contract=off','-fno-fast-math','-fno-builtin','-nostdlib'){
 if($option -notin $materialBuild.options){throw "Missing material build option: $option"}
}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
foreach($dll in 'soilfem_native_x64.dll','soilfem_mat_x64.dll'){
 $dllSource=if($dll -eq 'soilfem_native_x64.dll'){(Resolve-Path -LiteralPath $NativeDll).Path}else{(Resolve-Path -LiteralPath $MaterialDll).Path}
 Copy-Item -LiteralPath $dllSource -Destination $OutputDirectory
}
$bookPath=Join-Path $OutputDirectory '2DSoilFEM_inco_element_20261020.xlsm'
& pwsh -NoProfile -File (Join-Path $PSScriptRoot 'make_native_workbook.ps1') -Src $SourceWorkbook -Dst $bookPath -ModDir (Join-Path $repo 'src/vba')
if($LASTEXITCODE -ne 0){throw 'Source verification/build failed.'}
for($wait=0;$wait -lt 25;$wait++){
 if(-not (Get-Process EXCEL -ErrorAction SilentlyContinue)){break}
 Start-Sleep -Seconds 1
}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel remains running; no settings update dispatched.'}
$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application
 $app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $book=$app.Workbooks.Open($bookPath,0,$false);$macro="'"+$book.Name+"'!"
 $app.Run($macro+'FEMWriteNumericSetting','IO_FAST_PATH',1.0)
 $app.Run($macro+'FEMWriteNumericSetting','EVENT_LOG_FAST',1.0)
 $app.Run($macro+'FEMWriteNumericSetting','NATIVE_ASSEMBLY',1.0)
 $app.Run($macro+'FEMWriteNumericSetting','NATIVE_GMRES_TUNED',1.0)
 $app.Run($macro+'FEMWriteNumericSetting','NATIVE_GMRES_PREPARED',1.0)
 foreach($key in 'NATIVE_GMRES_PREPARED','NATIVE_GMRES_TUNED','NATIVE_ASSEMBLY','EVENT_LOG_FAST','IO_FAST_PATH','NATIVE_STATE_COPY','NATIVE_GMRES_SIMD'){
  if([double]$app.Run($macro+'P6ReadSetting',$key,-999.0) -ne 1){throw "Setting missing: $key"}
 }
 if([string]$app.Run($macro+'P6ReadTextSetting','NATIVE_PRECISION','MISSING') -cne 'AUTO'){throw 'Expected AUTO precision.'}
 $book.Save();$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
 $book=$app.Workbooks.Open($bookPath,0,$true)
 foreach($key in 'NATIVE_GMRES_PREPARED','NATIVE_GMRES_TUNED','NATIVE_ASSEMBLY','EVENT_LOG_FAST','IO_FAST_PATH','NATIVE_STATE_COPY','NATIVE_GMRES_SIMD'){
  if([double]$app.Run($macro+'P6ReadSetting',$key,-999.0) -ne 1){throw "Setting did not persist: $key"}
 }
 if([string]$app.Run($macro+'P6ReadTextSetting','NATIVE_PRECISION','MISSING') -cne 'AUTO'){throw 'Precision setting did not persist.'}
 $book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
}finally{
 if($null -ne $book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
@'
2DSoilFEM INCONSISTENT 要素行列積・ILU構築高速化版 20261020_INCO_ELEMENT_01

XLSMとDLL2個を同じフォルダに置き、64bit Windows Excelで開いてください。
保存済み結果は元ブックのものです。入力を確認して再解析してください。
Double・収束条件・材料則・SRM探索条件を維持しています。
3項行列積のループ制御と、ILU構築時の繰返し列検索を削減しています。
材料DLLはGCCのO3を使用し、x87・丸め・FMA禁止・fast-math禁止の設定を維持しています。
直交化2回・従来の真残差確認・反復上限・許容値を維持します。
今回の追加変更は加算順序・Doubleへの丸め位置を維持します。前版との比較ではGMRES途中値を含むログ全体が一致しました。
NATIVE_GMRES_PREPARED=0で対角逆数を準備しないGMRES経路へ戻せます。キー欠落時も0です。
旧DLL・非対応環境では従来経路へ戻ります。材料DLLはx87・MSVCR100演算を維持しています。
検証条件はdocs/INCONSISTENT_ELEMENT_PERF.mdを参照してください。
'@ | Set-Content -LiteralPath (Join-Path $OutputDirectory 'README.txt') -Encoding UTF8
$modules=@{}
foreach($name in 'CONTROL','FEMCore','FEMSolver','FEMIo','FEMEngine','FEMNative','FEMNativeMat','FEMNativeKern'){
 $modules[$name]=(Get-FileHash -LiteralPath (Join-Path $repo "src/vba/$name.bas")).Hash
}
$files=@(Get-ChildItem -LiteralPath $OutputDirectory -File|ForEach-Object{@{name=$_.Name;sha256=(Get-FileHash -LiteralPath $_.FullName).Hash}})
@{build='20261020_INCO_ELEMENT_01';material_build=$materialBuild;native_gmres_prepared=1;native_gmres_tuned=1;native_assembly=1;event_log_fast=1;io_fast_path=1;precision='AUTO';native_state_copy=1;native_gmres_simd=1;source_workbook_sha256=$sourceHash;source_sha256_by_module=$modules;files=$files}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $OutputDirectory 'build.json') -Encoding UTF8
if((Get-FileHash -LiteralPath $SourceWorkbook).Hash -cne $sourceHash){throw 'Source workbook changed.'}
Write-Output "Built and verified $bookPath (fixed-triplet products and monotone ILU construction)"
