param([Parameter(Mandatory=$true)][string]$Root,[Parameter(Mandatory=$true)][string]$Report)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Existing Excel; do not dispatch.'}
$Root=(Resolve-Path $Root).Path
$xl=New-Object -ComObject Excel.Application
$records=[Collections.Generic.List[object]]::new()
try{
 $xl.Visible=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 foreach($case in 'dry','wet','prescribed','continuous','prescribed_steady','soil_steady','partial','slip','opening','stages','srm','struct_pair','wet_davis','slip_davis'){
  $folder=Join-Path $Root $case
  $file=@(Get-ChildItem $folder -Filter '*.xlsm' | Where-Object {$_.Name -notlike '~$*'})
  if($file.Count -ne 1){throw "Expected one accepted book: $case"}
  $sha=(Get-FileHash $file[0].FullName).Hash
  $xl.EnableEvents=$true
  $wb=$xl.Workbooks.Open($file[0].FullName,0,$true)
  $xl.EnableEvents=$false
  try{
   $node=$wb.Worksheets.Item('結果節点');$element=$wb.Worksheets.Item('結果要素')
   if(-not [string]$node.Range('A2').Value2 -or -not [string]$element.Range('A2').Value2){throw "Empty saved mechanical results: $case"}
   $signature=$null
   if($case -ne 'dry'){
    $signature=[string]$xl.Run("'"+$wb.Name+"'!GWInputSignature")
    if($signature -ne [string]$wb.Worksheets.Item('水圧結果').Range('R5').Value2){throw "Stale water signature: $case"}
    if(-not [string]$wb.Worksheets.Item('接合結果').Range('A2').Value2){throw "Empty joint result: $case"}
   }
   if([string]$wb.Names.Item('_FEMViewerResultsDirty').RefersTo -ne '=FALSE'){throw "Saved result marked dirty: $case"}
   $records.Add(@{case=$case;status='PASS';book=$file[0].Name;sha256=$sha;node_rows=$node.Cells($node.Rows.Count,1).End(-4162).Row-1;element_rows=$element.Cells($element.Rows.Count,1).End(-4162).Row-1;signature=$signature;normal_workbook_open=$true})
  }finally{$wb.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}
  if((Get-FileHash $file[0].FullName).Hash -ne $sha){throw "Read-only verification changed book: $case"}
 }
 @{status='PASS';count=$records.Count;cases=@($records)}|ConvertTo-Json -Depth 8|Set-Content $Report -Encoding UTF8
}finally{
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
