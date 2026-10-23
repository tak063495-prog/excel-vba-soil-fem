param(
 [Parameter(Mandatory=$true)][string]$Book,
 [Parameter(Mandatory=$true)][string]$Frame,
 [Parameter(Mandatory=$true)][string]$Out,
 [switch]$Screen
)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; do not dispatch.'}
$sourceBook=(Resolve-Path -LiteralPath $Book).Path
$Frame=(Resolve-Path -LiteralPath $Frame).Path
foreach($file in 'node.csv','water.csv','gauss.csv'){
 if(-not(Test-Path -LiteralPath (Join-Path $Frame $file))){throw "Frame is missing ${file}: $Frame"}
}
if(Test-Path -LiteralPath $Out){throw 'Preserve previous evidence.'}
$Out=[IO.Path]::GetFullPath($Out);New-Item -ItemType Directory -Path $Out|Out-Null
$sourceHash=(Get-FileHash -LiteralPath $sourceBook).Hash
$copy=Join-Path $Out 'fresh_results_input.xlsm'
Copy-Item -LiteralPath $sourceBook -Destination $copy
$dll=Join-Path (Split-Path -Parent $sourceBook) 'soilfem_native_x64.dll'
if(Test-Path -LiteralPath $dll){Copy-Item -LiteralPath $dll -Destination $Out}
function Trace([string]$message){Add-Content -LiteralPath (Join-Path $Out 'steps.log') -Value "$(Get-Date -Format o) $message" -Encoding utf8}
function Get-FrameHash([string]$folder){
 $manifest=(Get-ChildItem -LiteralPath $folder -File|Sort-Object Name|ForEach-Object { $_.Name+' '+(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) -join "`n"
 $sha=[Security.Cryptography.SHA256]::Create()
 try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($manifest))).Replace('-','')) }
 finally { $sha.Dispose() }
}
$frameHash=Get-FrameHash $Frame

$xl=$null;$wb=$null;$module=$null
try {
 $xl=New-Object -ComObject Excel.Application
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1;$xl.Visible=[bool]$Screen
 $wb=$xl.Workbooks.Open($copy,0,$false)
 # The saved history is relative to the source book, whereas the workbook copy lives under $Out.
 $history=$wb.Worksheets.Item('時間履歴')
 for($r=2;$r -le $history.Cells($history.Rows.Count,1).End(-4162).Row;$r++){
  $relative=[string]$history.Cells($r,11).Value2
  if($relative -and -not [IO.Path]::IsPathRooted($relative)){
   $absolute=[IO.Path]::GetFullPath((Join-Path (Split-Path -Parent $sourceBook) $relative))
   if(-not(Test-Path -LiteralPath $absolute)){throw "Saved history folder is missing: $absolute"}
   $history.Cells($r,11).Value2=$absolute
  }
 }
 foreach($component in @($wb.VBProject.VBComponents)){
  if($component.Name -eq 'FSFrameReopenDriver'){$wb.VBProject.VBComponents.Remove($component)}
 }
 $module=$wb.VBProject.VBComponents.Add(1);$module.Name='FSFrameReopenDriver'
 $module.CodeModule.AddFromString(@'
Public Function FSRShow(ByVal id As String) As String
 On Error GoTo Failed
 SuppressUserMessages = True
 FEM24Error = "": FEMViewerLastError = ""
 FEM24Results
 ThisWorkbook.Worksheets("結果").Unprotect
 ThisWorkbook.Worksheets("結果").Range("C4").Value2 = "時刻別の図"
 FEM24Refresh ThisWorkbook.Worksheets("結果")
 ThisWorkbook.Worksheets("結果").Unprotect
 ThisWorkbook.Worksheets("結果").Range("B8").Value2 = id
 FSShowTimeFrame
 If Len(FEM24Error) > 0 Then Err.Raise vbObjectError + 8291, , FEM24Error
 If Len(FEMViewerLastError) > 0 Then Err.Raise vbObjectError + 8292, , FEMViewerLastError
 With ThisWorkbook.Worksheets("結果").Shapes("UI24Plot")
  If .Type <> 6 Or .GroupItems.Count <= 1 Or .Width <= 0 Or .Height <= 0 Then Err.Raise vbObjectError + 8293, , "時刻別のネイティブ図形グループが不正です。"
 End With
 FSRShow = "PASS"
 Exit Function
Failed:
 FSRShow = "ERROR|" & Err.Description & "|" & FEM24Error & "|" & FEMViewerLastError & "|" & FEMLastProc
End Function
Public Function FSRErrors() As String
 FSRErrors = FEM24Error & "|" & FEMViewerLastError
End Function
Public Function FSRLastProc() As String
 FSRLastProc = FEMLastProc
End Function
Public Function FSRGroupInfo() As Variant
 Dim result(0 To 3) As Double
 With ThisWorkbook.Worksheets("結果").Shapes("UI24Plot")
  result(0) = .Type: result(1) = .GroupItems.Count: result(2) = .Width: result(3) = .Height
 End With
 FSRGroupInfo = result
End Function
Public Function FSRNodes() As Variant
 Dim result() As Double, n As Long, intNode As Long
 ReDim result(0 To NumberOfFreeNode - 1, 0 To 2)
 For n = 0 To NumberOfFreeNode - 1
  intNode = P6GetInternalFreeNode(n)
  result(n,0) = n + 1: result(n,1) = TDisp(2 * intNode): result(n,2) = TDisp(2 * intNode + 1)
 Next n
 FSRNodes = result
End Function
Public Function FSRPressure() As Variant
 Dim result() As Double, e As Long, g As Long, n As Long
 ReDim result(0 To NumberOfElement * 4 - 1)
 For e = 0 To NumberOfElement - 1: For g = 0 To 3
  result(n) = GWNewPressure(e, g): n = n + 1
 Next g, e
 FSRPressure = result
End Function
Public Function FSRStress() As Variant
 Dim result() As Double, e As Long, g As Long, j As Long, n As Long
 ReDim result(0 To NumberOfElement * 4 * 17 - 1)
 For e = 0 To NumberOfElement - 1: For g = 0 To 3: For j = 0 To 16
  result(n) = Elem(e).Stmat(j, g): n = n + 1
 Next j, g, e
 FSRStress = result
End Function
'@)
 $frameId=Split-Path -Leaf $Frame
 Trace "direct FSShowTimeFrame begin $frameId"
 $status=[string]$xl.Run("'"+$wb.Name+"'!FSRShow",$frameId)
 $errors=[string]$xl.Run("'"+$wb.Name+"'!FSRErrors")
 $lastProc=[string]$xl.Run("'"+$wb.Name+"'!FSRLastProc")
 Trace "direct FSShowTimeFrame end $status"
 if($status -ne 'PASS'){throw $status}
 if($errors -ne '|'){throw "Viewer errors after direct frame display: $errors"}
 $group=@($xl.Run("'"+$wb.Name+"'!FSRGroupInfo"))
 $actual=$xl.Run("'"+$wb.Name+"'!FSRNodes")
 $expected=@(Import-Csv -LiteralPath (Join-Path $Frame 'node.csv'))
 if($actual.GetLength(0) -ne $expected.Count){throw 'Node count mismatch'}
 $umax=0.0
 for($n=0;$n -lt $expected.Count;$n++){
  $umax=[Math]::Max($umax,[Math]::Abs([double]$actual[$n,1]-[double]$expected[$n].ux))
  $umax=[Math]::Max($umax,[Math]::Abs([double]$actual[$n,2]-[double]$expected[$n].uy))
 }
 $pressure=@($xl.Run("'"+$wb.Name+"'!FSRPressure"));$stress=@($xl.Run("'"+$wb.Name+"'!FSRStress"))
 $waterLines=Get-Content -LiteralPath (Join-Path $Frame 'water.csv')
 $waterStart=[Array]::IndexOf($waterLines,'element,gauss,pressure_kPa')
 if($waterStart -lt 0){throw 'Missing pressure table'}
 $waterRows=@($waterLines[($waterStart+1)..($waterLines.Count-1)]|Where-Object {$_}|ConvertFrom-Csv -Header element,gauss,pressure_kPa)
 if($waterRows.Count -ne $pressure.Count){throw 'Loaded pressure count mismatch'}
 $pmax=0.0
 foreach($row in $waterRows){$index=4*([int]$row.element-1)+[int]$row.gauss-1;$pmax=[Math]::Max($pmax,[Math]::Abs([double]$pressure[$index]-[double]$row.pressure_kPa))}
 $smax=0.0
 foreach($row in @(Import-Csv -LiteralPath (Join-Path $Frame 'gauss.csv'))){
  $index=17*(4*([int]$row.elem-1)+[int]$row.gp)
  foreach($pair in @(@(0,'sx'),@(1,'sy'),@(2,'txy'),@(16,'sz'))){$smax=[Math]::Max($smax,[Math]::Abs([double]$stress[$index+[int]$pair[0]]-[double]$row.($pair[1]))) }
 }
 if($umax -gt 1e-10 -or $pmax -gt 1e-8 -or $smax -gt 1e-8){throw "Loaded frame differs: u=$umax p=$pmax stress=$smax"}
 @{status='PASS';frame_id=$frameId;frame=$Frame;frame_sha256=$frameHash;excel_build=$xl.Build;book_sha256=$sourceHash;errors=$errors;FEMLastProc=$lastProc;native_group=@{type=$group[0];items=$group[1];width=$group[2];height=$group[3]};max_displacement_difference=$umax;max_pressure_difference=$pmax;max_stress_difference=$smax}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $Out 'frame_reopen.json') -Encoding utf8
 if($Screen){
  Set-Content -LiteralPath (Join-Path $Out 'screen_ready.txt') 'ready' -Encoding utf8
  while(-not(Test-Path -LiteralPath (Join-Path $Out 'screen_release.txt'))){Start-Sleep -Milliseconds 300}
 }
 $wb.Close($false);$wb=$null
} finally {
 if($wb){try{$wb.Close($false)}catch{}}
 if($xl){$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
if((Get-FileHash -LiteralPath $sourceBook).Hash -ne $sourceHash){throw 'Source workbook was changed'}
