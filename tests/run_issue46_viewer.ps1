param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out,[switch]$CheckCoordinateGuard)
$ErrorActionPreference='Stop'; if(Get-Process -Name EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; do not dispatch.'}
if(Test-Path $Out){throw 'Preserve previous evidence.'}; New-Item -ItemType Directory $Out|Out-Null
$Book=(Resolve-Path $Book).Path;$sourceHash=(Get-FileHash $Book).Hash
$Out=[IO.Path]::GetFullPath($Out)
$copy=Join-Path $Out ([IO.Path]::GetFileName($Book));Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1;$wb=$xl.Workbooks.Open($copy,0,$false);$m=$wb.VBProject.VBComponents.Add(1);$m.Name='Issue46Driver'
 $m.CodeModule.AddFromString(@'
Public Function I46Plot() As String
 On Error GoTo Failed
 SuppressUserMessages=True: FEM24Error=""
 FEM24Plot
 If Len(FEM24Error)>0 Then Err.Raise vbObjectError+8240,,FEM24Error
 If Len(FEMViewerLastSvg)=0 Then Err.Raise vbObjectError+8241,,"No result drawing"
 I46Plot="PASS"
 Exit Function
Failed: I46Plot="ERROR|" & Err.Description
End Function
Public Function I46Svg() As String
 I46Svg=FEMViewerLastSvg
End Function
'@)
 $n=$wb.Name;$plot=$xl.Run("'$n'!I46Plot")
 if($plot -notlike 'PASS'){throw "FEM24Plot failed: $plot"}
 $svg=[string]$xl.Run("'$n'!I46Svg");if($svg -notmatch '<svg'){throw 'Missing generated SVG'}
 [IO.File]::WriteAllText((Join-Path $Out 'result.svg'),$svg,[Text.UTF8Encoding]::new($false))
 $picture=$wb.Worksheets.Item('図').Shapes.Item('P1Agg_ViewSVG')
 if($picture.Type -ne 6 -or $picture.GroupItems.Count -le 1 -or $picture.Width -le 0 -or $picture.Height -le 0){throw 'Result drawing is not a valid native shape group'}
 $imageInfo=@{type=$picture.Type;native_shape_count=$picture.GroupItems.Count;width=$picture.Width;height=$picture.Height}
 $wb.Save();$wb.Close($false);$wb=$xl.Workbooks.Open($copy,0,$false)
 $reopen=$xl.Run("'$n'!I46Plot");if($reopen -ne 'PASS'){throw "Reopened drawing failed: $reopen"}
 $rawGuard='NOT_REQUESTED'
 if($CheckCoordinateGuard){
  $node=$wb.Worksheets.Item('節点データ');$old=$node.Range('B2').Value2;$node.Range('B2').Value2=[double]$old+0.123
  $rawGuard=$xl.Run("'$n'!I46Plot");$node.Range('B2').Value2=$old
  if($rawGuard -notlike 'ERROR|*' -or $rawGuard -notmatch '一致|座標|再解析'){throw "Saved-coordinate guard failed: $rawGuard"}
 }
 $xl.EnableEvents=$true
 $node=$wb.Worksheets.Item('節点データ');$old=$node.Range('B2').Value2;$node.Range('B2').Value2=[double]$old+0.123
 $guard=$xl.Run("'$n'!I46Plot");$node.Range('B2').Value2=$old
 if($guard -notlike 'ERROR|*' -or $guard -notmatch '一致|変更|再解析|入力'){throw "Coordinate/input guard was not observed: $guard"}
 @{status='PASS';plot=$plot;reopen=$reopen;native_figure=$imageInfo;coordinate_guard=$guard;saved_coordinate_guard=$rawGuard;book_sha256=$sourceHash;note='Prepared analysis with reordered IDs; saved/reopened test copy; real SheetChange input guard'}|ConvertTo-Json|Set-Content (Join-Path $Out 'issue46.json') -Encoding utf8
}finally{
 if($wb){try{$wb.Close($false)}catch{}}
 foreach($reference in @($picture,$node,$m,$wb)){
  if($null -ne $reference -and [Runtime.InteropServices.Marshal]::IsComObject($reference)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference)}
 }
 $picture=$null;$node=$null;$m=$null;$wb=$null
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
}
if((Get-FileHash $Book).Hash -ne $sourceHash){throw 'Source workbook was changed'}
