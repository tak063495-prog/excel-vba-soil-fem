param(
 [string]$SourceWorkbook=(Join-Path $PSScriptRoot '../workbook/inco_element_20261020/2DSoilFEM_inco_element_20261020.xlsm'),
 [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$SourceWorkbook=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Preserve existing output; choose a new directory.'}
$sourceHash=(Get-FileHash -LiteralPath $SourceWorkbook).Hash
New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
foreach($dll in 'soilfem_native_x64.dll','soilfem_mat_x64.dll'){
 Copy-Item -LiteralPath (Join-Path (Split-Path $SourceWorkbook) $dll) -Destination $OutputDirectory
}
$bookPath=Join-Path $OutputDirectory '2DSoilFEM_ui_20261020.xlsm'
& pwsh -NoProfile -File (Join-Path $PSScriptRoot 'native/make_native_workbook.ps1') -Src $SourceWorkbook -Dst $bookPath -ModDir (Join-Path $repo 'src/vba') -Modules FEMUiLayout
if($LASTEXITCODE -ne 0){throw 'VBA source verification/compilation failed.'}
for($wait=0;$wait -lt 20;$wait++){
 if(-not(Get-Process EXCEL -ErrorAction SilentlyContinue)){break}
 Start-Sleep -Seconds 1
}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel remains running; no second instance started.'}
$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application
 $app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $book=$app.Workbooks.Open($bookPath,0,$false);$macro="'"+$book.Name+"'!"
 $layoutResult=[string]$app.Run($macro+'FEMUiRefreshInputs')
 if($layoutResult -cne 'PASS'){throw $layoutResult}
 $app.Run($macro+'FEMUiValidateSettings')
 $app.Run($macro+'FEMUiShowHome')
 $book.Save();$book.Close($false)
 $book=$app.Workbooks.Open($bookPath,0,$true);$macro="'"+$book.Name+"'!"
 $layoutResult=[string]$app.Run($macro+'FEMUiRefreshInputs')
 if($layoutResult -cne 'PASS'){throw $layoutResult}
 $app.Run($macro+'FEMUiValidateSettings')
 foreach($key in 'NATIVE_STATE_COPY','NATIVE_ASSEMBLY','NATIVE_GMRES_SIMD','NATIVE_GMRES_TUNED','NATIVE_GMRES_PREPARED','IO_FAST_PATH','EVENT_LOG_FAST'){
  if([double]$app.Run($macro+'P6ReadSetting',$key,-999.0) -ne 1){throw "Recommended runtime flag not persisted: $key"}
 }
 foreach($key in 'NATIVE_DLL','NATIVE_PRECISION'){
  if([string]$app.Run($macro+'P6ReadTextSetting',$key,'MISSING') -cne 'AUTO'){throw "Recommended AUTO not persisted: $key"}
 }
 $compile=$app.VBE.CommandBars.FindControl(1,578)
 if($compile.Enabled){$compile.Execute()}
 if($app.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Compile remains enabled.'}
 $book.Close($false)
}finally{
 if($book){try{$book.Close($false)}catch{}}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
if((Get-FileHash -LiteralPath $SourceWorkbook).Hash -cne $sourceHash){throw 'Source workbook changed.'}
$files=@(Get-ChildItem -LiteralPath $OutputDirectory -File | ForEach-Object {@{name=$_.Name;bytes=$_.Length;sha256=(Get-FileHash -LiteralPath $_.FullName).Hash}})
@{source=$SourceWorkbook;source_sha256=$sourceHash;files=$files;compile='PASS';reopen='PASS';ui_stamp='20261020_UI_02'} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory 'BUILD_RECORD.json') -Encoding UTF8
Write-Output "PASS $bookPath"
