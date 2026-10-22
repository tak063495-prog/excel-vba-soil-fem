param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$PackageRecord,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel running; no test started.'}
$Book=(Resolve-Path $Book).Path
$record=Get-Content $PackageRecord -Raw | ConvertFrom-Json
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($Book,0,$true)
 $codes=@{}
 foreach($name in $record.modules){$cm=$wb.VBProject.VBComponents.Item($name).CodeModule;$codes[$name]=$cm.Lines(1,$cm.CountOfLines)}
 $codes|ConvertTo-Json -Depth 3|Set-Content $Out -Encoding utf8
}finally{
 if($wb){$wb.Close($false)};$xl.Quit()
 foreach($o in @($cm,$wb,$xl)){if($o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
