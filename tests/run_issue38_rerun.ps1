param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no dispatch.'}
$Book=(Resolve-Path $Book).Path;$hash=(Get-FileHash $Book).Hash;$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path $Out){throw 'Preserve previous evidence'}
New-Item -ItemType Directory $Out|Out-Null;$copy=Join-Path $Out 'rerun.xlsm';Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1;$wb=$xl.Workbooks.Open($copy,0,$false)
 $name=$wb.Name
 $null=$xl.Run("'$name'!FEMWriteTextSetting",'EXPORT_LOAD','')
 $first=$xl.Run("'$name'!WFRun");if($first -notlike 'PASS|*'){throw $first}
 $sheet=$wb.Worksheets.Item('結果塑性');$data=$sheet.UsedRange.Value2
 $oldMax=0.;for($r=2;$r -le $data.GetLength(0);$r++){if($data[$r,11] -is [double]){$oldMax=[Math]::Max($oldMax,[double]$data[$r,11])}}
 if($oldMax -le 1e-10){throw "First run did not plasticize: $oldMax"}
 $wb.Worksheets.Item('材料データ').Range('G2').Value2=1000000.
 $second=$xl.Run("'$name'!WFRun");if($second -notlike 'PASS|*'){throw $second}
 $data=$sheet.UsedRange.Value2;$seen=@{};$maxPlastic=0.;$rows=0
 for($r=2;$r -le $data.GetLength(0);$r++){
  if($null -eq $data[$r,1]){continue};$key="$($data[$r,1])/$($data[$r,2])/$($data[$r,3])"
  if($seen.ContainsKey($key)){throw "Duplicate stage/element/GP: $key"};$seen[$key]=$true;$rows++
  $maxPlastic=[Math]::Max($maxPlastic,[Math]::Abs([double]$data[$r,11]))
 }
 if($maxPlastic -gt 1e-12){throw "Old plastic accumulation remains: $maxPlastic"}
 $record=@{status='PASS';old_max_plastic=$oldMax;new_max_plastic=$maxPlastic;rows=$rows;unique_keys=$seen.Count;disp=@($xl.Run("'$name'!WFDisplacements"));stress=@($xl.Run("'$name'!WFStress"));metrics=@($xl.Run("'$name'!WFMetrics"));book_sha256=$hash}
 $record|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Out 'rerun.json') -Encoding utf8
 $wb.Close($false)
}finally{if($wb){try{$wb.Close($false)}catch{}};$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);[GC]::Collect();[GC]::WaitForPendingFinalizers()}
if((Get-FileHash $Book).Hash -ne $hash){throw 'Source workbook changed'}
