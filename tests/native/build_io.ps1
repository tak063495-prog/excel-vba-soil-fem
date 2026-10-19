param(
 [string]$SourceWorkbook=(Join-Path $PSScriptRoot '../../workbook/inco_native_20261020/2DSoilFEM_inco_native_20261020.xlsm'),
 [Parameter(Mandatory=$true)][string]$OutputDirectory
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
 Copy-Item -LiteralPath (Join-Path (Split-Path $SourceWorkbook) $dll) -Destination $OutputDirectory
}
$bookPath=Join-Path $OutputDirectory '2DSoilFEM_inco_io_20261020.xlsm'
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
 foreach($key in 'IO_FAST_PATH','NATIVE_STATE_COPY','NATIVE_GMRES_SIMD'){
  if([double]$app.Run($macro+'P6ReadSetting',$key,-999.0) -ne 1){throw "Setting missing: $key"}
 }
 if([string]$app.Run($macro+'P6ReadTextSetting','NATIVE_PRECISION','MISSING') -cne 'AUTO'){throw 'Expected AUTO precision.'}
 $book.Save();$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
 $book=$app.Workbooks.Open($bookPath,0,$true)
 foreach($key in 'IO_FAST_PATH','NATIVE_STATE_COPY','NATIVE_GMRES_SIMD'){
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
2DSoilFEM INCONSISTENT I/O高速化版 20261020_INCO_IO_01

XLSMとDLL2個を同じフォルダに置き、64bit Windows Excelで開いてください。
保存済みの結果は元ブックのものです。入力を確認して再解析してください。
Double・材料則・収束条件・SRM探索条件を維持しています。
前版のDLL2個は変更していません。AUTO/SSE2/STRICTの選択も維持します。
ログの出力先を再利用し、フォルダ確認を実書き込み時にまとめました。
保存先の変更は検知します。ロックで保留されたログは解除後に再試行します。
保留ログが残ると次の解析開始を拒否します。出力先のロックを解除して再試行してください。
比較条件とログ障害時の制限はリポジトリのdocs/INCONSISTENT_IO_PERF.mdを参照してください。
'@ | Set-Content -LiteralPath (Join-Path $OutputDirectory 'README.txt') -Encoding UTF8
$modules=@{}
foreach($name in 'CONTROL','FEMCore','FEMSolver','FEMIo','FEMEngine','FEMNative','FEMNativeMat','FEMNativeKern'){
 $modules[$name]=(Get-FileHash -LiteralPath (Join-Path $repo "src/vba/$name.bas")).Hash
}
$files=@(Get-ChildItem -LiteralPath $OutputDirectory -File|ForEach-Object{@{name=$_.Name;sha256=(Get-FileHash -LiteralPath $_.FullName).Hash}})
@{build='20261020_INCO_IO_01';io_fast_path=1;precision='AUTO';native_state_copy=1;native_gmres_simd=1;source_workbook_sha256=$sourceHash;source_sha256_by_module=$modules;files=$files}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $OutputDirectory 'build.json') -Encoding UTF8
if((Get-FileHash -LiteralPath $SourceWorkbook).Hash -cne $sourceHash){throw 'Source workbook changed.'}
Write-Output "Built and verified $bookPath (DLLs copied unchanged)"
