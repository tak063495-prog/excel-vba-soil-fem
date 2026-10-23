param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out,[Parameter(Mandatory=$true)][string]$Frames)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running'}
if(Test-Path $Out){throw 'Preserve evidence'}
$Book=(Resolve-Path $Book).Path;$sourceHash=(Get-FileHash $Book).Hash;$Out=[IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory $Out|Out-Null;$copy=Join-Path $Out 'frame_forces.xlsm';Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$xl=New-Object -ComObject Excel.Application
try{
 $xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false);$history=$wb.Worksheets.Item('時間履歴');$folders=@{}
 for($r=2;$r -le $history.Cells($history.Rows.Count,1).End(-4162).Row;$r++){
  $p=[string]$history.Cells($r,11).Value2
  if($p -and -not [IO.Path]::IsPathRooted($p)){$p=Join-Path (Split-Path $Book) $p;$history.Cells($r,11).Value2=$p}
  if($p){$folders[[string]$history.Cells($r,1).Value2]=$p}
 }
 $mod=$wb.VBProject.VBComponents.Add(1);$mod.Name='Issue51FrameDriver'
 $mod.CodeModule.AddFromString(@'
Public Function IFShow(ByVal id As String) As String
 On Error GoTo Failed
 SuppressUserMessages = True
 FEM24Results
 ThisWorkbook.Worksheets("結果").Unprotect
 ThisWorkbook.Worksheets("結果").Range("C4").Value2 = "時刻別の図"
 FEM24Refresh ThisWorkbook.Worksheets("結果")
 ThisWorkbook.Worksheets("結果").Range("B8").Value2 = id
 FSShowTimeFrame
 IFShow = FEM24Error & "|" & FEMViewerLastError
 Exit Function
Failed:
 IFShow = "ERROR|" & Err.Description & "|" & FEMLastProc
End Function
Public Function IFRows() As Variant
 IFRows = ThisWorkbook.Worksheets("結果節点").UsedRange.Value2
End Function
Public Function IFMetrics() As Variant
 IFMetrics = Array(P6ReactionSumX, P6ReactionSumY, RelativeResidualFree)
End Function
'@)
 $prefix="'"+$wb.Name+"'!";$records=[Collections.Generic.List[object]]::new()
 foreach($id in $Frames.Split(',')){
  if(-not $folders.ContainsKey($id)){throw "Missing frame $id"}
  $status=[string]$xl.Run($prefix+'IFShow',$id)
  if($status -ne '|'){throw $status}
  $a=$xl.Run($prefix+'IFRows');$rows=[Collections.Generic.List[object]]::new()
  for($r=1;$r -le $a.GetLength(0);$r++){$row=[Collections.Generic.List[object]]::new();for($c=1;$c -le $a.GetLength(1);$c++){$row.Add($a[$r,$c])};$rows.Add(@($row))}
  $records.Add(@{id=$id;frame=$folders[$id];status='PASS';sheet_rows=@($rows);metrics=@($xl.Run($prefix+'IFMetrics'))})
 }
 @{status='EXECUTED';book_sha256=$sourceHash;excel_build=$xl.Build;records=@($records)}|ConvertTo-Json -Depth 10|Set-Content (Join-Path $Out 'frame_forces.json') -Encoding utf8
 $wb.Close($false)
}finally{if($wb){try{$wb.Close($false)}catch{}};$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);[GC]::Collect();[GC]::WaitForPendingFinalizers()}
if((Get-FileHash $Book).Hash -ne $sourceHash){throw 'Source changed'}
