param(
 [string]$SourceWorkbook=(Join-Path $PSScriptRoot '../../workbook/2DSoilFEM_standard48_native_20261018.xlsm'),
 [Parameter(Mandatory=$true)][string]$OutputDirectory,
 [ValidateSet('STRICT','SSE2','AUTO')][string]$Precision='AUTO',
 [switch]$Inconsistent
)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before building.'}
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'OutputDirectory already exists; preserve prior builds.'}
$source=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$beforeHash=(Get-FileHash -LiteralPath $source).Hash
& (Join-Path $repo 'native/build.cmd')
if($LASTEXITCODE -ne 0){throw 'DLL build failed.'}
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
$bookName=if($Inconsistent){'2DSoilFEM_inco_native_20261020.xlsm'}else{'2DSoilFEM_simd_20261019.xlsm'}
$stamp=[regex]::Match((Get-Content (Join-Path $repo 'src/vba/FEMCore.bas') -Raw),'FEM_BUILD_STAMP As String = "([^"]+)"').Groups[1].Value
$bookPath=Join-Path $OutputDirectory $bookName
Copy-Item -LiteralPath (Join-Path $repo 'native/bin/soilfem_native_x64.dll') -Destination $OutputDirectory
Copy-Item -LiteralPath (Join-Path $repo 'workbook/soilfem_mat_x64.dll') -Destination $OutputDirectory
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'make_native_workbook.ps1') -Src $source -Dst $bookPath -ModDir (Join-Path $repo 'src/vba')
if($LASTEXITCODE -ne 0){throw 'Workbook build failed; preserve output for diagnostics.'}
for($closeWait=0;$closeWait -lt 25;$closeWait++){
 if(-not (Get-Process EXCEL -ErrorAction SilentlyContinue)){break}
 Start-Sleep -Seconds 1
}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel is running; no settings update dispatched.'}
$app=$null; $book=$null
try {
 $app=New-Object -ComObject Excel.Application
 $app.EnableEvents=$false; $app.DisplayAlerts=$false; $app.AutomationSecurity=1
 $book=$app.Workbooks.Open($bookPath,0,$false)
 $macro="'"+$book.Name+"'!"
 $app.Run($macro+'FEMWriteTextSetting','NATIVE_PRECISION',$Precision)
 $app.Run($macro+'FEMWriteNumericSetting','NATIVE_STATE_COPY',[double][int][bool]$Inconsistent)
 $app.Run($macro+'FEMWriteNumericSetting','NATIVE_GMRES_SIMD',[double][int][bool]$Inconsistent)
 # Place the choice on the existing settings page and bind its canonical key.
 $sheet=$book.Worksheets.Item('高速化設定'); $sheet.Unprotect()
 $lastKeyRow=$sheet.Cells.Item($sheet.Rows.Count,6).End(-4162).Row
 $row=$lastKeyRow+2
 $sheet.Range("B${lastKeyRow}:F${lastKeyRow}").Copy($sheet.Range("B${row}:F${row}"))
 $sheet.Cells.Item($row,2).Value2='DLLの演算方式（倍精度）'
 $sheet.Cells.Item($row,3).NumberFormat='@'; $sheet.Cells.Item($row,3).Value2=$Precision
 $sheet.Cells.Item($row,5).Value2='AUTO=CPU自動選択 / SSE2=倍精度 / STRICT=VBA互換'
 $sheet.Cells.Item($row,6).Value2='NATIVE_PRECISION'
 $sheet.Cells.Item($row,3).Validation.Delete()
 $sheet.Cells.Item($row,3).Validation.Add(3,1,1,'STRICT,SSE2,AUTO')
 $sheet.Cells.Item($row,3).Interior.Color=10092543
 $sheet.Cells.Item($row,3).MergeArea.Locked=$false
 $sheet.Rows.Item($row).RowHeight=30
 $sheet.Cells.Item($row,5).WrapText=$true
 $app.Run($macro+'FEMUiBindSettings')
 $app.Run($macro+'FEMUiSettingsChanged',$sheet,$sheet.Cells.Item($row,3))
 $sheet.Protect('', $false, $true, $true, $true)
 $app.Run($macro+'FEMInvalidateSettingCache')
 $status=[string]$app.Run($macro+'FEMNativeStatusText')
 if($status -notlike 'NATIVE(*'){throw ('Native DLL unavailable: '+$status)}
 $book.Save(); $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book); $book=$null
 $book=$app.Workbooks.Open($bookPath,0,$true)
 if([string]$app.Run($macro+'P6ReadTextSetting','NATIVE_PRECISION','MISSING') -cne $Precision){throw 'Precision setting did not persist.'}
 foreach($key in 'NATIVE_STATE_COPY','NATIVE_GMRES_SIMD'){
  if([double]$app.Run($macro+'P6ReadSetting',$key,-999.0) -ne [int][bool]$Inconsistent){throw "Setting did not persist: $key"}
 }
 $book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book); $book=$null
 @'
2DSoilFEM SIMD比較用ブック（ビルド番号はbuild.jsonを参照）

XLSMとDLL 2個を同じフォルダに置き、64bit Excelで開いてください。
旧ブックと比較する場合は、先にExcelをすべて閉じてください。
保存済みの結果は旧版の結果です。入力を確認し、再解析してから利用してください。

「高速化設定」の「DLLの演算方式（倍精度）」で設定します。
AUTO: 対応CPU・OSでAVX2/FMA、それ以外はSSE2。
SSE2: 倍精度SSE2。
STRICT: 従来のVBA互換演算。

Doubleを維持し、VBAとのビット一致条件だけを任意に解除しています。
収束許容値・材料則・SRM探索条件は従来のままです。
今回のSIMD対象は帯LUと前進・後退代入です。材料DLLは従来版です。
悪条件・限界状態付近では丸め差の影響が増幅される場合があります。
全解析条件での妥当性確認を済ませた標準版への置換ではなく、比較用の試験版です。
'@ | Set-Content -LiteralPath (Join-Path $OutputDirectory 'README.txt') -Encoding UTF8
 if($Inconsistent){
  @"
2DSoilFEM INCONSISTENT高速化版 $stamp

XLSMとDLL 2個を同じフォルダに置き、64bit Excelで開いてください。
保存済みの結果は旧版の結果です。入力を確認し、再解析してください。
「高速化設定」の「DLLの演算方式（倍精度）」はAUTOが既定です。
AUTO: AVX2/FMA対応CPUでは同命令、それ以外はSSE2。SSE2も選択可能。
STRICT: 従来のVBA互換演算。GMRESは従来のx87 DLLを使用します。
Double・収束許容値・材料則・SRM探索条件を維持しています。
GMRESと状態復元を高速化。材料DLLとログ出力処理は従来版です。
微小な丸め差を許容します。すべてのモデルで同じ高速化率を保証しません。
検証条件・比較結果・再現手順は docs/INCONSISTENT_NATIVE_PERF.md を参照。
"@ | Set-Content -LiteralPath (Join-Path $OutputDirectory 'README.txt') -Encoding UTF8
 }
 [pscustomobject]@{build=$stamp;precision=$Precision;state_copy=[bool]$Inconsistent;gmres_simd=[bool]$Inconsistent;native_status=$status;source_sha256=$beforeHash;files=@(Get-ChildItem -LiteralPath $OutputDirectory -File | ForEach-Object {@{name=$_.Name;sha256=(Get-FileHash -LiteralPath $_.FullName).Hash}})} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'build.json') -Encoding UTF8
} finally {
 if($null -ne $book){$book.Close($false); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
 if($null -ne $app){$app.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if((Get-FileHash -LiteralPath $source).Hash -cne $beforeHash){throw 'Source workbook changed.'}
Write-Output "Built $bookPath precision=$Precision"
