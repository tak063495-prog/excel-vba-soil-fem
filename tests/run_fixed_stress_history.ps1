param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Frame,[Parameter(Mandatory=$true)][string]$Out,[switch]$Continue,[switch]$Plot,[switch]$Screen)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; do not dispatch.'}
$Book=(Resolve-Path $Book).Path;$Frame=(Resolve-Path $Frame).Path
if(Test-Path $Out){throw 'Preserve previous evidence.'}
$bookHash=(Get-FileHash $Book).Hash
$Out=[IO.Path]::GetFullPath($Out);New-Item -ItemType Directory $Out|Out-Null
function Trace([string]$message){Add-Content -LiteralPath (Join-Path $Out 'steps.log') -Value "$(Get-Date -Format o) $message" -Encoding utf8}
$sourceBook=$Book
$Book=Join-Path $Out 'history_input.xlsm'
Copy-Item -LiteralPath $sourceBook -Destination $Book
$dll=Join-Path (Split-Path $sourceBook) 'soilfem_native_x64.dll'
if(Test-Path $dll){Copy-Item -LiteralPath $dll -Destination $Out}
$xl=New-Object -ComObject Excel.Application
try {
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1;$xl.Visible=[bool]$Screen
 $wb=$xl.Workbooks.Open($Book,0,$false)
 # Preserve original result/checkpoint evidence: only the isolated copy is resumed.
 $history=$wb.Worksheets.Item('時間履歴')
 for($r=2;$r -le $history.Cells($history.Rows.Count,1).End(-4162).Row;$r++){
  $path=[string]$history.Cells($r,11).Value2
  if($path -and -not [IO.Path]::IsPathRooted($path)){$history.Cells($r,11).Value2=Join-Path (Split-Path $sourceBook) $path}
 }
 foreach($component in @($wb.VBProject.VBComponents)){if($component.Name -eq 'FSHistoryDriver'){$wb.VBProject.VBComponents.Remove($component)}}
 $module=$wb.VBProject.VBComponents.Add(1);$module.Name='FSHistoryDriver'
 $module.CodeModule.AddFromString(@'
Public Function FHLoad(ByVal folder As String) As String
 On Error GoTo Failed
 SuppressUserMessages = True
 P3LoadTimeFrame folder
 FHLoad = "PASS|" & CStr(FSTime) & "|" & CStr(FSAccepted) & "|" & CStr(FSResumePending)
 Exit Function
Failed:
 FHLoad = "ERROR|" & Err.Description & "|" & FEMLastProc
End Function
Public Function FHNodes() As Variant
 Dim result() As Double, n As Long, intNode As Long
 ReDim result(0 To NumberOfFreeNode - 1, 0 To 2)
 For n = 0 To NumberOfFreeNode - 1
  intNode = P6GetInternalFreeNode(n)
  result(n,0) = n + 1
  result(n,1) = TDisp(2 * intNode)
  result(n,2) = TDisp(2 * intNode + 1)
 Next n
 FHNodes = result
End Function
Public Function FHPressure() As Variant
 Dim result() As Double, e As Long, g As Long, n As Long
 ReDim result(0 To NumberOfElement * 4 - 1)
 For e = 0 To NumberOfElement - 1
  For g = 0 To 3: result(n) = GWNewPressure(e, g): n = n + 1: Next g
 Next e
 FHPressure = result
End Function
Public Function FHStress() As Variant
 Dim result() As Double, e As Long, g As Long, j As Long, n As Long
 ReDim result(0 To NumberOfElement * 4 * 17 - 1)
 For e = 0 To NumberOfElement - 1: For g = 0 To 3: For j = 0 To 16
  result(n) = Elem(e).Stmat(j, g): n = n + 1
 Next j, g, e
 FHStress = result
End Function
Public Function FHActive() As Variant
 Dim result() As Double, e As Long
 ReDim result(0 To NumberOfElement - 1)
 For e = 0 To NumberOfElement - 1: result(e) = Abs(CDbl(P3ElementActive(e))): Next e
 FHActive = result
End Function
Public Function FHRun() As String
 On Error GoTo Failed
 SuppressUserMessages = True
 P0_RunAnalysis
 FHRun = ResultStatus & "|" & AnalysisMessage
 Exit Function
Failed:
 FHRun = "ERROR|" & Err.Description
End Function
Public Function FHPlot(ByVal id As String) As String
 On Error GoTo Failed
 SuppressUserMessages = True
 FEM24Install
 ThisWorkbook.Worksheets("結果").Unprotect
 ThisWorkbook.Worksheets("結果").Range("C4").Value2 = "時刻別の図"
 FEM24Refresh ThisWorkbook.Worksheets("結果")
 ThisWorkbook.Worksheets("結果").Unprotect
 ThisWorkbook.Worksheets("結果").Range("B8").Value2 = id
 FSShowTimeFrame
 If Len(FEM24Error) > 0 Then Err.Raise vbObjectError + 8190, , FEM24Error
 FHPlot = "PASS|" & CStr(ThisWorkbook.Worksheets("結果").Shapes.Count) & "|" & FSStoredCaption()
 Exit Function
Failed:
 FHPlot = "ERROR|" & Err.Description
End Function
'@)
 Trace 'FHLoad begin'
 $status=$xl.Run("'"+$wb.Name+"'!FHLoad",$Frame)
 Trace "FHLoad end $status"
 if($status -notlike 'PASS|*'){throw $status}
 $actual=$xl.Run("'"+$wb.Name+"'!FHNodes")
 $expected=@(Import-Csv (Join-Path $Frame 'node.csv'))
 if($actual.GetLength(0) -ne $expected.Count){throw 'Node count mismatch'}
 $max=0.0
 for($n=0;$n -lt $expected.Count;$n++){
  $max=[Math]::Max($max,[Math]::Abs([double]$actual[$n,1]-[double]$expected[$n].ux))
  $max=[Math]::Max($max,[Math]::Abs([double]$actual[$n,2]-[double]$expected[$n].uy))
 }
 if($max -gt 1e-10){throw "Loaded displacement differs: $max"}
 $pressure=@($xl.Run("'"+$wb.Name+"'!FHPressure"));$stress=@($xl.Run("'"+$wb.Name+"'!FHStress"));$active=@($xl.Run("'"+$wb.Name+"'!FHActive"))
 $pmax=0.0;$smax=0.0
 $waterLines=Get-Content (Join-Path $Frame 'water.csv')
 $waterStart=[Array]::IndexOf($waterLines,'element,gauss,pressure_kPa')
 if($waterStart -lt 0){throw 'Missing pressure table'}
 $waterRows=@($waterLines[($waterStart+1)..($waterLines.Count-1)] | Where-Object {$_} | ConvertFrom-Csv -Header element,gauss,pressure_kPa)
 if($waterRows.Count -ne $pressure.Count){throw 'Loaded pressure count mismatch'}
 foreach($r in $waterRows){$k=4*([int]$r.element-1)+[int]$r.gauss-1;$pmax=[Math]::Max($pmax,[Math]::Abs([double]$pressure[$k]-[double]$r.pressure_kPa))}
 $gauss=@(Import-Csv (Join-Path $Frame 'gauss.csv'))
 foreach($r in $gauss){
  $k=17*(4*([int]$r.elem-1)+[int]$r.gp)
  foreach($pair in @(@(0,'sx'),@(1,'sy'),@(2,'txy'),@(16,'sz'))){$smax=[Math]::Max($smax,[Math]::Abs([double]$stress[$k+[int]$pair[0]]-[double]$r.($pair[1])))}
 }
 if($pmax -gt 1e-8 -or $smax -gt 1e-8){throw "Loaded fields differ: p=$pmax stress=$smax"}
 @{status=$status;max_displacement_difference=$max;max_pressure_difference=$pmax;max_stress_difference=$smax;pressure=$pressure;stress=$stress;active=$active;frame=$Frame;excel_build=$xl.Build;book_sha256=$bookHash}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Out 'load.json') -Encoding utf8
 if($Plot){
  Trace 'FHPlot begin'
  $plotStatus=$xl.Run("'"+$wb.Name+"'!FHPlot",(Split-Path $Frame -Leaf))
  if($plotStatus -notlike 'PASS|*'){throw $plotStatus}
  Trace "FHPlot end $plotStatus"
  @{status=$plotStatus;frame=$Frame}|ConvertTo-Json|Set-Content (Join-Path $Out 'plot.json') -Encoding utf8
 }
 if($Screen){
  Set-Content (Join-Path $Out 'screen_ready.txt') 'ready' -Encoding utf8
  while(-not(Test-Path (Join-Path $Out 'screen_release.txt'))){Start-Sleep -Milliseconds 300}
 }
 if($Continue){
  Trace 'EXPORT_LOAD write begin'
  $null=$xl.Run("'"+$wb.Name+"'!FEMWriteTextSetting",'EXPORT_LOAD',$Frame)
  Trace 'EXPORT_LOAD write end; FHRun begin'
  $status=$xl.Run("'"+$wb.Name+"'!FHRun")
  Trace "FHRun end $status"
  if($status -notlike 'PASS|*'){throw $status}
  $actual=$xl.Run("'"+$wb.Name+"'!FHNodes")
  $pressure=@($xl.Run("'"+$wb.Name+"'!FHPressure"));$stress=@($xl.Run("'"+$wb.Name+"'!FHStress"));$active=@($xl.Run("'"+$wb.Name+"'!FHActive"))
  $rows=[Collections.Generic.List[object]]::new()
  for($n=0;$n -lt $actual.GetLength(0);$n++){$rows.Add(@($actual[$n,0],$actual[$n,1],$actual[$n,2]))}
  @{status=$status;nodes=@($rows);pressure=$pressure;stress=$stress;active=$active}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Out 'continued.json') -Encoding utf8
 }
 $wb.Close($false)
}finally{
 if($wb){try{$wb.Close($false)}catch{}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}



if((Get-FileHash $sourceBook).Hash -ne $bookHash){throw 'Source evidence workbook changed'}
