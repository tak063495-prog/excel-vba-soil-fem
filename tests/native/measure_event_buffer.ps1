param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
if(Test-Path -LiteralPath $Out){throw 'Preserve evidence.'}
$Book=(Resolve-Path -LiteralPath $Book).Path
$dir=[IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Out))
if(-not(Test-Path -LiteralPath $dir)){New-Item -ItemType Directory -Path $dir | Out-Null}
$copy=Join-Path $dir 'buffer_measure.xlsm'
if(Test-Path -LiteralPath $copy){throw 'Preserve test copy.'}
Copy-Item -LiteralPath $Book -Destination $copy
$harness=@'
Private Declare PtrSafe Function EventQpc Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef value As Currency) As Long
Private Declare PtrSafe Function EventFreq Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef value As Currency) As Long
Public Function EventBufferMeasure(ByVal method As Long, ByVal width As Long) As Variant
 Dim parts(0 To 511) As String, line As String, blockText As String
 Dim i As Long, j As Long, total As Double, start As Currency, finish As Currency, freq As Currency
 Dim required As Long, capacity As Long, grown As Long
 On Error GoTo Failed
 line=String$(width,"X"):mEvBuf="":mEvUsed=0:mEvCount=0:mEventFast=(method=1)
 EventFreq freq:EventQpc start
 For i=0 To 65535
  If method=2 Then
   parts(i Mod 512)=line & vbCrLf
  ElseIf method=0 Then
   mEvBuf=mEvBuf & line & vbCrLf
  Else
   required=mEvUsed+Len(line)+2:capacity=Len(mEvBuf)
   If required>capacity Then
    grown=capacity:If grown<65536 Then grown=65536
    Do While grown<required:grown=grown*2:Loop
    mEvBuf=mEvBuf & Space$(grown-capacity)
   End If
   Mid$(mEvBuf,mEvUsed+1,Len(line))=line
   Mid$(mEvBuf,mEvUsed+Len(line)+1,2)=vbCrLf
   mEvUsed=required
  End If
  If i Mod 512=511 Then
   Select Case method
    Case 0:blockText=mEvBuf:mEvBuf=""
    Case 1:blockText=left$(mEvBuf,mEvUsed):mEvUsed=0
    Case 2:blockText=Join(parts,vbNullString):For j=0 To 511:parts(j)="":Next j
   End Select
   total=total+Len(blockText)
  End If
 Next i
 EventQpc finish
 If total<>65536#*(width+2) Then Err.Raise vbObjectError+6401,,"output length mismatch"
 EventBufferMeasure=Array((finish-start)/freq,total)
 Exit Function
Failed:
 EventBufferMeasure=Array(-1,Err.Number,Err.Description)
End Function
'@
$app=$null;$wb=$null;$module=$null;$runs=@()
try{
 $app=New-Object -ComObject Excel.Application
 $app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $wb=$app.Workbooks.Open($copy,0,$false)
 $module=$wb.VBProject.VBComponents.Item('FEMIo').CodeModule;$module.AddFromString($harness)
 $compile=$app.VBE.CommandBars.FindControl(1,578);if($compile.Enabled){$compile.Execute()}
 if($app.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Microbenchmark compilation failed.'}
 foreach($width in 128,512){foreach($pair in 1,2,3){
  $order=if($pair%2){@(0,1,2)}else{@(2,1,0)}
  foreach($method in $order){
   $r=@($app.Run("'"+$wb.Name+"'!EventBufferMeasure",$method,$width))
   if([double]$r[0] -le 0){throw ('Microbenchmark failed: '+($r -join '|'))}
   $runs+=@{width=$width;pair=$pair;method=$method;seconds=[double]$r[0];characters=[double]$r[1]}
  }
 }}
 $runs | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $Out -Encoding UTF8
 $runs | Format-Table
 $wb.Close($false)
}finally{
 if($null -ne $module){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($module)}
 if($null -ne $wb){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)}
 if($null -ne $app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
