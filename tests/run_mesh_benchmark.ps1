param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out,[int]$Repeats=3,[switch]$Extended,[switch]$Direct)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Existing Excel; no test started.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path $Out){throw 'Preserve existing evidence.'}
New-Item -ItemType Directory -Path $Out | Out-Null
$copy=Join-Path $Out 'mesh_bench.xlsm';Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$true;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!";$xl.EnableEvents=$false
 # Record an otherwise swallowed error in silent benchmarks; algorithm unchanged.
 $cm=$wb.VBProject.VBComponents.Item('CONTROL').CodeModule
 $control=$cm.Lines(1,$cm.CountOfLines)
 $control=$control.Replace('errorDescription = Err.description', 'errorDescription = Err.description'+"`r`n    FEMUiMeshPreviewError = errorDescription")
 $cm.DeleteLines(1,$cm.CountOfLines);$cm.AddFromString($control)
 $helper=$wb.VBProject.VBComponents.Add(1);$helper.Name='MeshBench'
 $helper.CodeModule.AddFromString(@'
Option Explicit
Public Function BenchMesh(ByVal spacing As Double, ByVal policy As Long, ByVal flat As Boolean, ByVal geometryKind As Long) As String
 On Error GoTo Bad
 Dim ws As Worksheet, pts As Variant, i As Long, savedGuard As Variant
 SuppressUserMessages=True
 FEMInvalidateSettingCache
 FEMWriteNumericSetting "MESH_GEN_COMPARE",policy
 FEMWriteNumericSetting "MESH_REPAIR_TERRAIN_COMPARE",1
 Set ws=ThisWorkbook.Worksheets("形状入力")
 ws.Range("B8:D2000").ClearContents:ws.Range("G8:J2000").ClearContents
 If geometryKind = 2 Then
  pts=Array(Array(1,0,0),Array(2,20,0),Array(3,20,8),Array(4,0,8),Array(5,8,2),Array(6,12,2),Array(7,12,4),Array(8,8,4))
 ElseIf geometryKind = 3 Then
  pts=Array(Array(1,0,0),Array(2,20,1),Array(3,20,8),Array(4,0,8))
 ElseIf geometryKind = 4 Then
  pts=Array(Array(1,0,0),Array(2,20,0),Array(3,20,4),Array(4,10,4),Array(5,12,8),Array(6,0,8))
 ElseIf geometryKind = 1 Then
  pts=Array(Array(1,0,0),Array(2,20,0),Array(3,20,4),Array(4,6.4,4),Array(5,6,8),Array(6,0,8))
 ElseIf flat Then
  pts=Array(Array(1,0,0),Array(2,20,0),Array(3,20,8),Array(4,0,8))
 Else
  pts=Array(Array(1,0,0),Array(2,20,0),Array(3,20,4),Array(4,12,4),Array(5,6,8),Array(6,0,8))
 End If
 For i=0 To UBound(pts)
  ws.Cells(i+8,2).Value2=pts(i)(0):ws.Cells(i+8,3).Value2=pts(i)(1):ws.Cells(i+8,4).Value2=pts(i)(2)
 Next i
 ws.Cells(8,7).Value2=1:ws.Cells(8,8).Value2="外側"
 If flat Or geometryKind = 2 Then ws.Cells(8,9).Value2="1,2,3,4" Else ws.Cells(8,9).Value2="1,2,3,4,5,6"
 ws.Cells(8,10).Value2=spacing
 If geometryKind = 2 Then
  ws.Cells(9,7).Value2=2:ws.Cells(9,8).Value2="内側":ws.Cells(9,9).Value2="5,6,7,8":ws.Cells(9,10).Value2=spacing
 End If
 If geometryKind = 6 Then
  savedGuard=ThisWorkbook.Worksheets("載荷").Range("A2:J2").Value2
  ThisWorkbook.Worksheets("載荷").Range("A2:D2").Value2=Array(1,"1",0,"FORCE")
 ElseIf geometryKind = 7 Then
  savedGuard=ThisWorkbook.Worksheets("接合").Range("A2:G2").Value2
  ThisWorkbook.Worksheets("接合").Range("A2:E2").Value2=Array(1,2,3,1,"test guard")
 ElseIf geometryKind = 8 Then
  savedGuard=ThisWorkbook.Worksheets("水条件").Range("A216:F216").Value2
  ThisWorkbook.Worksheets("水条件").Range("A216:F216").Value2=Array(1,1,"HEAD","",0,"test guard")
 ElseIf geometryKind = 9 Then
  ws.Cells(8,9).Value2="6,5,4,3,2,1"
 End If
 MeshPerfReset True
 FEMUiGenerateShape
 If geometryKind = 6 Then
  ThisWorkbook.Worksheets("載荷").Range("A2:J2").Value2=savedGuard
 ElseIf geometryKind = 7 Then
  ThisWorkbook.Worksheets("接合").Range("A2:G2").Value2=savedGuard
 ElseIf geometryKind = 8 Then
  ThisWorkbook.Worksheets("水条件").Range("A216:F216").Value2=savedGuard
 End If
 If Not FEMShapeGenerationOK Or Len(FEMUiMeshPreviewError)>0 Then BenchMesh="ERROR|" & FEMUiMeshPreviewError:Exit Function
 BenchMesh="PASS|" & FEMReadTextSetting("MESH_GEN_LAST_METHOD","") & "|" & P5QualitySummary():Exit Function
Bad:BenchMesh="ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
'@)
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()}
 if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Compile failed.'}
 $records=[Collections.Generic.List[object]]::new()
 $cases=@(@{name='slope2';h=2.0;flat=$false;policy=2},@{name='slope1';h=1.0;flat=$false;policy=2},@{name='flat1';h=1.0;flat=$true;policy=2},@{name='slope2_quality';h=2.0;flat=$false;policy=0})
 if($Extended){$cases=@(@{name='steep_fallback';h=2.0;flat=$false;policy=2;geometryKind=1},@{name='hole_fallback';h=2.0;flat=$true;policy=2;geometryKind=2},@{name='slope2_delaunay';h=2.0;flat=$false;policy=1})}
 if($Extended -and $Direct){
  # Holes are now supported by the complex terrain regression, not a rejection.
  $cases=@($cases|Where-Object {$_.geometryKind -ne 2})
  $cases[0].reject=$true
  $cases+=@(@{name='sloping_bottom_blocks';h=2.0;flat=$true;policy=2;geometryKind=3},@{name='overhang_rejected';h=2.0;flat=$false;policy=2;geometryKind=4;reject=$true})
  $cases+=@(@{name='force_rejected';h=2.0;flat=$false;policy=2;geometryKind=6;reject=$true},@{name='joint_rejected';h=2.0;flat=$false;policy=2;geometryKind=7;reject=$true},@{name='water_boundary_rejected';h=2.0;flat=$false;policy=2;geometryKind=8;reject=$true},@{name='too_fine_rejected';h=1e-12;flat=$false;policy=2;reject=$true},@{name='clockwise';h=2.0;flat=$false;policy=2;geometryKind=9})
 }
 function MeshSnapshot {
  $snapshot=@{}
  foreach($name in @('節点データ','要素データ')){$sheet=$wb.Worksheets.Item($name);$last=$sheet.Cells($sheet.Rows.Count,1).End(-4162).Row;$width=if($name -eq '節点データ'){14}else{10};$snapshot[$name]=$sheet.Range($sheet.Cells(1,1),$sheet.Cells($last,$width)).Value2}
  return ($snapshot|ConvertTo-Json -Depth 6 -Compress)
 }
 foreach($case in $cases){
  for($i=0;$i -lt $Repeats;$i++){
   if($case.reject){$beforeSnapshot=MeshSnapshot}
   $watch=[Diagnostics.Stopwatch]::StartNew();$result=[string]$xl.Run($prefix+'BenchMesh',[double]$case.h,[int]$case.policy,[bool]$case.flat,[int]$case.geometryKind);$watch.Stop()
   if($case.reject){
    if(-not $result.StartsWith('ERROR|') -or $result.Length -lt 10){throw ('Expected explicit rejection: '+$result)}
    if((MeshSnapshot) -cne $beforeSnapshot){throw 'Rejected shape modified existing mesh.'}
    $records.Add(@{case=$case.name;repeat=$i;seconds=$watch.Elapsed.TotalSeconds;result=$result;status='EXPECTED_REJECT';mesh_unchanged=$true;profile=[string]$xl.Run($prefix+'MeshPerfReport')})
    $records|ConvertTo-Json -Depth 5 | Set-Content (Join-Path $Out 'summary.json') -Encoding UTF8
    Write-Output ($case.name+' EXPECTED_REJECT mesh unchanged '+$result)
    continue
   }
   if(-not $result.StartsWith('PASS|')){throw $result}
   $profile=[string]$xl.Run($prefix+'MeshPerfReport')
   $n=$wb.Worksheets.Item('節点データ');$e=$wb.Worksheets.Item('要素データ')
   $nn=[int]$n.Cells($n.Rows.Count,1).End(-4162).Row;$ne=[int]$e.Cells($e.Rows.Count,1).End(-4162).Row
   $record=@{case=$case.name;repeat=$i;seconds=$watch.Elapsed.TotalSeconds;result=$result;profile=$profile;nodes=$n.Range("A1:N$nn").Value2;elements=$e.Range("A1:J$ne").Value2}
   $record|ConvertTo-Json -Depth 6 | Set-Content (Join-Path $Out ($case.name+'_'+$i+'.json')) -Encoding UTF8
   $records.Add(@{case=$case.name;repeat=$i;seconds=$watch.Elapsed.TotalSeconds;result=$result;profile=$profile;nodes=$nn-1;elements=$ne-1})
   $records|ConvertTo-Json -Depth 5 | Set-Content (Join-Path $Out 'summary.json') -Encoding UTF8
   Write-Output ($case.name+' '+$i+' '+$watch.Elapsed.TotalSeconds+' '+$result);Write-Output $profile
  }
 }
}finally{
 if($wb){$wb.Close($false)}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
