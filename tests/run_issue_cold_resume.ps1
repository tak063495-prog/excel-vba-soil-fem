param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Frame,[Parameter(Mandatory=$true)][string]$Out,[switch]$WaterObservables,[switch]$ViewOnly)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no dispatch'}
if(Test-Path $Out){throw 'Preserve evidence'}
$Book=(Resolve-Path $Book).Path;$Frame=(Resolve-Path $Frame).Path;$hash=(Get-FileHash $Book).Hash
$Out=[IO.Path]::GetFullPath($Out);New-Item -ItemType Directory $Out|Out-Null
$copy=Join-Path $Out 'cold_resume.xlsm';Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$xl=New-Object -ComObject Excel.Application
try{
 $xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
 # The prepared fixture book contains WaterFixtureDriver. Do not load a view
 # or any checkpoint first: WFRun is the first analysis call in this process.
 $null=$xl.Run($prefix+'FEMWriteTextSetting','EXPORT_LOAD',$Frame)
 $timer=[Diagnostics.Stopwatch]::StartNew()
 if($ViewOnly){$null=$xl.Run($prefix+'SetP0SilentMode',$true);$null=$xl.Run($prefix+'P3LoadTimeFrame',$Frame);$status='PASS|VIEW'}else{$status=$xl.Run($prefix+'WFRun')}
 $timer.Stop()
 if($WaterObservables){
  @($xl.Run($prefix+'WF50State'))|ConvertTo-Json -Depth 12|Set-Content (Join-Path $Out 'water_state.json') -Encoding utf8
  @($xl.Run($prefix+'WF50Loaded'))|ConvertTo-Json -Depth 12|Set-Content (Join-Path $Out 'water_loaded.json') -Encoding utf8
 }
 $a=$xl.Run($prefix+'WFOriginalNodes');$nodes=[Collections.Generic.List[object]]::new()
 for($n=0;$n -lt $a.GetLength(0);$n++){$nodes.Add(@($a[$n,0],$a[$n,1],$a[$n,2]))}
 @{status=$status;nodes=@($nodes);pressure=@($xl.Run($prefix+'WFPressure'));stress=@($xl.Run($prefix+'WFStress'));metrics=@($xl.Run($prefix+'WFMetrics'));
   seconds=$timer.Elapsed.TotalSeconds;source_book_sha256=$hash;build=$xl.Run($prefix+'WFBuild');excel_version=$xl.Version;excel_build=$xl.Build;fresh_excel_process=$true;preloaded_frame=$false}|
   ConvertTo-Json -Depth 10|Set-Content (Join-Path $Out 'continued.json') -Encoding utf8
 if($status -notlike 'PASS|*'){throw "Cold resume failed: $status"}
 $wb.Close($false)
}finally{
 if($wb){try{$wb.Close($false)}catch{}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);[GC]::Collect();[GC]::WaitForPendingFinalizers()
}
if((Get-FileHash $Book).Hash -ne $hash){throw 'Source changed'}
