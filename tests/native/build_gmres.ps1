param(
 [string]$SourceWorkbook=(Join-Path $PSScriptRoot '../../workbook/inco_assembly_20261020/2DSoilFEM_inco_assembly_20261020.xlsm'),
 [Parameter(Mandatory=$true)][string]$OutputDirectory,
 [string]$NativeDll=(Join-Path $PSScriptRoot '../../native/bin/soilfem_native_x64.dll')
)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$SourceWorkbook=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve existing package.'}
$sourceHash=(Get-FileHash -LiteralPath $SourceWorkbook).Hash
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
foreach($dll in 'soilfem_native_x64.dll','soilfem_mat_x64.dll'){
 $dllSource=if($dll -eq 'soilfem_native_x64.dll'){(Resolve-Path -LiteralPath $NativeDll).Path}else{Join-Path (Split-Path $SourceWorkbook) $dll}
 Copy-Item -LiteralPath $dllSource -Destination $OutputDirectory
}
$bookPath=Join-Path $OutputDirectory '2DSoilFEM_inco_gmres_20261020.xlsm'
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
 foreach($key in 'NATIVE_GMRES_TUNED','NATIVE_ASSEMBLY','EVENT_LOG_FAST','IO_FAST_PATH','NATIVE_STATE_COPY','NATIVE_GMRES_SIMD'){
  if([double]$app.Run($macro+'P6ReadSetting',$key,-999.0) -ne 1){throw "Setting missing: $key"}
 }
 if([string]$app.Run($macro+'P6ReadTextSetting','NATIVE_PRECISION','MISSING') -cne 'AUTO'){throw 'Expected AUTO precision.'}
 $book.Save();$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
 $book=$app.Workbooks.Open($bookPath,0,$true)
 foreach($key in 'NATIVE_GMRES_TUNED','NATIVE_ASSEMBLY','EVENT_LOG_FAST','IO_FAST_PATH','NATIVE_STATE_COPY','NATIVE_GMRES_SIMD'){
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
2DSoilFEM INCONSISTENT GMRESカーネル高速化版 20261020_INCO_GMRES_03

XLSMとDLL2個を同じフォルダに置き、64bit Windows Excelで開いてください。
保存済み結果は元ブックのものです。入力を確認して再解析してください。
Double・収束条件・材料則・SRM探索条件を維持しています。
GMRESのILU前処理・疎行列積を倍精度SIMDで計算します。
直交化2回・従来の真残差確認・反復上限・許容値を維持します。
加算の順序が変わるため途中のGMRES残差・再スタート回数には差が出ます。
NATIVE_GMRES_TUNED=0で従来のGMRES経路へ戻せます。キー欠落時も0です。
旧DLL・非対応環境では従来経路へ戻ります。材料DLLは前版と同一です。
検証条件はdocs/INCONSISTENT_GMRES_PERF.mdを参照してください。
'@ | Set-Content -LiteralPath (Join-Path $OutputDirectory 'README.txt') -Encoding UTF8
$modules=@{}
foreach($name in 'CONTROL','FEMCore','FEMSolver','FEMIo','FEMEngine','FEMNative','FEMNativeMat','FEMNativeKern'){
 $modules[$name]=(Get-FileHash -LiteralPath (Join-Path $repo "src/vba/$name.bas")).Hash
}
$files=@(Get-ChildItem -LiteralPath $OutputDirectory -File|ForEach-Object{@{name=$_.Name;sha256=(Get-FileHash -LiteralPath $_.FullName).Hash}})
@{build='20261020_INCO_GMRES_03';native_gmres_tuned=1;native_assembly=1;event_log_fast=1;io_fast_path=1;precision='AUTO';native_state_copy=1;native_gmres_simd=1;source_workbook_sha256=$sourceHash;source_sha256_by_module=$modules;files=$files}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $OutputDirectory 'build.json') -Encoding UTF8
if((Get-FileHash -LiteralPath $SourceWorkbook).Hash -cne $sourceHash){throw 'Source workbook changed.'}
Write-Output "Built and verified $bookPath (native GMRES DLL; material DLL unchanged)"
