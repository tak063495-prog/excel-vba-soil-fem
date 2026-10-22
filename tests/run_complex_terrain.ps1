param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out,[switch]$Visible,[string]$CaseName="")
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Existing Excel; no test started.'}
if(Test-Path $Out){throw 'Preserve previous evidence.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
New-Item -ItemType Directory -Path $Out | Out-Null
$copy=Join-Path $Out 'complex_mesh.xlsm';Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$fixtures=Get-Content "$PSScriptRoot/fixtures/complex_terrain.json" -Raw -Encoding UTF8 | ConvertFrom-Json
if($CaseName){$fixtures=@($fixtures|Where-Object {$_.name -eq $CaseName});if($fixtures.Count -ne 1){throw "Unknown case."}}
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$true;$xl.AutomationSecurity=1;$xl.Visible=[bool]$Visible
 $wb=$xl.Workbooks.Open($copy,0,$false);$xl.EnableEvents=$false;$prefix="'"+$wb.Name+"'!"
 $mod=$wb.VBProject.VBComponents.Add(1);$mod.Name='ComplexTerrainTest'
 $mod.CodeModule.AddFromString(@'
Option Explicit
Public Function CTGenerate() As String
 On Error GoTo Failed
 SuppressUserMessages=True
 FEMWriteNumericSetting "MESH_GEN_COMPARE",2
 FEMInvalidateSettingCache
 FEMUiGenerateShape
 If Not FEMShapeGenerationOK Or Len(FEMUiMeshPreviewError)>0 Then CTGenerate="ERROR|" & FEMUiMeshPreviewError:Exit Function
 CTGenerate="PASS|" & MeshAdaptLastSummary
 Exit Function
Failed:CTGenerate="ERROR|" & Err.Description
End Function
'@)
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()}
 if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Compile failed.'}
 $summary=[Collections.Generic.List[object]]::new()
 function Snapshot{
  $a=@{};foreach($s in '節点データ','要素データ'){$ws=$wb.Worksheets.Item($s);$last=$ws.Cells($ws.Rows.Count,1).End(-4162).Row;$a[$s]=$ws.Range($ws.Cells(1,1),$ws.Cells($last,14)).Value2}
  return ($a|ConvertTo-Json -Depth 6 -Compress)
 }
 foreach($case in $fixtures){
  $ws=$wb.Worksheets.Item('形状入力');$ws.Range('B8:D2000').ClearContents();$ws.Range('G8:J2000').ClearContents()
  $r=8;foreach($point in $case.points){for($c=0;$c -lt 3;$c++){$ws.Cells($r,$c+2).Value2=[double]$point[$c]};$r++}
  $r=8;foreach($loop in $case.loops){$ws.Cells($r,7).Value2=[double]($r-7);$ws.Cells($r,8).Value2=[string]$loop.kind;$ws.Cells($r,9).Value2=[string]$loop.ids;$ws.Cells($r,10).Value2=[double]$case.h;$r++}
  $before=Snapshot
  $watch=[Diagnostics.Stopwatch]::StartNew();$result=[string]$xl.Run($prefix+'CTGenerate');$watch.Stop()
  if($case.reject){
   if(-not $result.StartsWith('ERROR|')){throw ('Expected rejection: '+$case.name+' '+$result)}
   if((Snapshot) -cne $before){throw 'Rejected input modified the existing mesh.'}
   $record=@{case=$case.name;status='EXPECTED_REJECT';mesh_unchanged=$true;result=$result;seconds=$watch.Elapsed.TotalSeconds}
  }else{
   if(-not $result.StartsWith('PASS|')){throw ($case.name+' '+$result)}
   $n=$wb.Worksheets.Item('節点データ');$e=$wb.Worksheets.Item('要素データ')
   $nn=$n.Cells($n.Rows.Count,1).End(-4162).Row;$ne=$e.Cells($e.Rows.Count,1).End(-4162).Row
   $record=@{case=$case.name;status='PASS';seconds=$watch.Elapsed.TotalSeconds;result=$result;nodes=$n.Range("A1:N$nn").Value2;elements=$e.Range("A1:J$ne").Value2}
   $wb.SaveCopyAs((Join-Path $Out ($case.name+'.xlsm')))
  }
  $record|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Out ($case.name+'.json')) -Encoding UTF8
  $summary.Add(@{case=$case.name;status=$record.status;seconds=$record.seconds;result=$result})
  $summary|ConvertTo-Json -Depth 6|Set-Content (Join-Path $Out 'summary.json') -Encoding UTF8
  Write-Output ($case.name+' '+$record.status+' '+$record.seconds+' '+$result)
 }
 if($Visible){
  Set-Content (Join-Path $Out 'ready.txt') 'Own test copy ready for visual inspection.'
  while(-not (Test-Path (Join-Path $Out 'close.txt'))){Start-Sleep -Milliseconds 500}
 }
}finally{
 if($wb){$wb.Close($false)}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
