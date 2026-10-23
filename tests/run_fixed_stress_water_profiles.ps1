param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; do not dispatch.'}
$Book=(Resolve-Path $Book).Path
$sourceHash=(Get-FileHash $Book).Hash
if(Test-Path $Out){throw 'Preserve previous evidence.'}
New-Item -ItemType Directory $Out|Out-Null
$frame=Get-ChildItem (Split-Path $Book) -Recurse -Filter fixed_stress.csv | Where-Object {$_.Directory.Name -like 'C*'} | Sort-Object {$_.Directory.Name} | Select-Object -Last 1
if(-not $frame){throw 'No accepted frame'}
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($Book,0,$false);$m=$wb.VBProject.VBComponents.Add(1);$m.Name='FSWaterProfileDriver'
 $m.CodeModule.AddFromString(@'
Option Explicit
Public Function WPLoad(ByVal folder As String) As String
 On Error GoTo Failed
 SuppressUserMessages=True: P3LoadTimeFrame folder: WPLoad="PASS": Exit Function
Failed: WPLoad="ERROR|" & Err.Description
End Function
Public Function WP(ByVal f As Double, ByVal face As Boolean) As Variant
 Dim v() As Double,i As Long, fy As Double,mz As Double
 ReDim v(lastDof)
 GWTransientProfile "W","W2",f
 If face Then
  GWAddProfileFaceLoad 7,2,"*",v
 Else
  GWTransientForce v
 End If
 For i=0 To lastDof Step 2
  fy=fy+v(i+1): mz=mz+XXX(i)*v(i+1)-XXX(i+1)*v(i)
 Next i
 WP=Array(GWProfileLevel("*",.5),GWProfileLevel("W2",.5),fy,mz)
End Function
'@)
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()}
 if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Driver compile failed'}
 $loaded=$xl.Run("'$($wb.Name)'!WPLoad",$frame.DirectoryName);if($loaded -ne 'PASS'){throw $loaded}
 $w=$wb.Worksheets.Item('水条件');$records=[Collections.Generic.List[object]]::new()
 function SetProfiles($rows){
  $w.Range('A12:D211').ClearContents()
  $values=[object[,]]::new($rows.Count,4)
  for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt 4;$c++){$values[$r,$c]=$rows[$r][$c]}}
  $w.Range($w.Cells(12,1),$w.Cells((11+$rows.Count),4)).Value2=$values
 }
 function CheckProfile([string]$kind,[double]$f,[bool]$face,[double[]]$expected){
  $v=@($xl.Run("'$($wb.Name)'!WP",$f,$face))
  if($v.Count -ne 4){throw "Bad result count $($v.Count)"}
  $err=0.0;for($i=0;$i -lt 4;$i++){$err=[Math]::Max($err,[Math]::Abs([double]$v[$i]-$expected[$i]))}
  if($err -gt 1e-9){throw "$kind fraction=$f mismatch $err actual=$v expected=$expected"}
  $records.Add(@{kind=$kind;fraction=$f;face=$face;actual=$v;expected=$expected;max_error=$err})
 }
 SetProfiles @(@('W',0,4,''),@('W',1,4,''),@('W2',0,5,''),@('W2',1,5,''))
 foreach($f in @(0,.25,.5,.75,1)){
  foreach($face in @($false,$true)){CheckProfile 'uniform' $f $face @((4+$f),5,(-9.81*(2+$f)),(-9.81*(2+$f)/2))}
 }
 SetProfiles @(@('W',0,4,''),@('W',.25,4.5,''),@('W',1,4,''),@('W2',0,5,''),@('W2',.75,5.5,''),@('W2',1,5,''))
 foreach($f in @(.25,.5,.75)){
  foreach($face in @($false,$true)){CheckProfile 'union-knots' $f $face @((4+1.0/3+$f),(5+1.0/3),(-9.81*(2.25+$f)),(-9.81*((2+$f)/2+.25*(1.25+.5*$f)/3)))}
 }
 $w.Cells(216,4).Value2='W2'
 CheckProfile 'explicit-target-fixed' .25 $false @((4+1.0/3+.25),(5+1.0/3),(-9.81*3.25),(-9.81*(1.5+.25*1.75/3)))
 @{status='PASS';checks=@($records);tolerance=1e-9;scope='profile/head, external force and moment, prescribed JOINT face-load path; uniform and different polyline knots';book_sha256=$sourceHash}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Out 'water_profiles.json') -Encoding utf8
}catch{Write-Output $_.ScriptStackTrace;Write-Output $_.Exception.ToString();throw}finally{if($wb){try{$wb.Close($false)}catch{}};$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);[GC]::Collect();[GC]::WaitForPendingFinalizers()}


if((Get-FileHash $Book).Hash -ne $sourceHash){throw 'Source evidence workbook changed'}
