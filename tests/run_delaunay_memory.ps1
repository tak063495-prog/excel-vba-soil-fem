param(
    [Parameter(Mandatory=$true)][string]$Book,
    [Parameter(Mandatory=$true)][string]$Out
)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Existing Excel; no test started.'}
$Book=(Resolve-Path $Book).Path; $Out=[IO.Path]::GetFullPath($Out)
if(Test-Path $Out){throw 'Preserve existing evidence.'}
New-Item -ItemType Directory -Path $Out | Out-Null
$copy=Join-Path $Out 'delaunay_memory_probe.xlsm'; Copy-Item $Book $copy

$xl=New-Object -ComObject Excel.Application
try {
    $xl.DisplayAlerts=$false; $xl.EnableEvents=$true; $xl.AutomationSecurity=1
    $wb=$xl.Workbooks.Open($copy,0,$false)
    $xl.EnableEvents=$false
    $prefix="'"+$wb.Name+"'!"
    $mesh=$wb.VBProject.VBComponents.Item('Mesh').CodeModule
    $source=$mesh.Lines(1,$mesh.CountOfLines)
    $before='    Set imno = CreateObject("Scripting.Dictionary")'; $beforeLabel='before_sparse_tables'
    if($source.IndexOf($before,[StringComparison]::Ordinal) -lt 0){$before='    ReDim imno(1 To IIf(NOB + NIB > 200, NOB + NIB, 200), 1 To IIf(NOB + NIB > 200, NOB + NIB, 200))'; $beforeLabel='before_dense_tables'}
    $after='    pinStage = "細分化ノード生成"'
    if($source.IndexOf($before,[StringComparison]::Ordinal) -lt 0){throw 'imno allocation marker not found.'}
    if($source.IndexOf($after,[StringComparison]::Ordinal) -lt 0){throw 'Subdivision marker not found.'}
    $source=$source.Replace($before,('    MeshMemoryProbe.Record "'+$beforeLabel+'"')+"`r`n"+$before)
    $source=$source.Replace($after,$after+"`r`n"+'    MeshMemoryProbe.Record "after_tables"')
    $mesh.DeleteLines(1,$mesh.CountOfLines); $mesh.AddFromString($source)
    $helper=$wb.VBProject.VBComponents.Add(1); $helper.Name='MeshMemoryProbe'
    $helper.CodeModule.AddFromString(@'
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal hProcess As LongPtr, ByRef counters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
#Else
Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal hProcess As Long, ByRef counters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
#End If
Private Type PROCESS_MEMORY_COUNTERS
 cb As Long: PageFaultCount As Long
 PeakWorkingSetSize As LongPtr: WorkingSetSize As LongPtr
 QuotaPeakPagedPoolUsage As LongPtr: QuotaPagedPoolUsage As LongPtr
 QuotaPeakNonPagedPoolUsage As LongPtr: QuotaNonPagedPoolUsage As LongPtr
 PagefileUsage As LongPtr: PeakPagefileUsage As LongPtr
End Type
Public outputPath As String
Public currentN As Long
Public probeCase As String
Public Sub Record(ByVal label As String)
 Dim c As PROCESS_MEMORY_COUNTERS, f As Integer
 c.cb=LenB(c): If GetProcessMemoryInfo(GetCurrentProcess(),c,c.cb)=0 Then Exit Sub
 f=FreeFile: Open outputPath For Append As #f
 Print #f, probeCase & "," & CStr(currentN) & "," & label & "," & CStr(Timer) & "," & CStr(c.WorkingSetSize) & "," & CStr(c.PagefileUsage)
 Close #f
End Sub
Public Sub PreparePolygon(ByVal n As Long, ByVal path As String)
 Dim ws As Worksheet, i As Long, theta As Double, pi As Double
 pi=4#*Atn(1#): Set ws=ThisWorkbook.Worksheets("図形定義")
 ws.Range("B2:C70000").ClearContents: ws.Range("E3:BZL3").ClearContents
 ws.Range("H1").Value2=n: ws.Range("J1").Value2=0: ws.Range("D3").Value2=1: ws.Range("D24").Value2=0
 ws.Range("E3").Value2=n: ws.Range("F3").Value2=1000000#
 For i=1 To n
  theta=2#*pi*(i-1)/n: ws.Cells(i+1,2).Value2=100#*Cos(theta): ws.Cells(i+1,3).Value2=100#*Sin(theta): ws.Cells(3,6+i).Value2=i
 Next i
 currentN=n: probeCase="polygon": outputPath=path: Record "before_prepare": P5PrepareGeometryDefinitionWorkspace
 PINPUT NEX,NIN,ibex,ibin,ibno,NOB,NIB,px,py,delx,xlok
 DumpMesh path
 Record "after_pinput": P5ReleaseGeometryDefinitionWorkspace
End Sub
Public Sub DumpMesh(ByVal path As String)
 Dim f As Integer, i As Long, j As Long, total As Long, count As Long
 f=FreeFile: Open path & ".nodes.csv" For Append As #f: total=NOB+NIB
 For i=1 To total: Print #f, probeCase & "," & currentN & "," & i & "," & px(i) & "," & py(i) & "," & xlok(i): Next i
 Close #f
 f=FreeFile: Open path & ".boundary.csv" For Append As #f
 For i=1 To NEX+NIN
  If i<=NEX Then count=ibex(i) Else count=ibin(i-NEX)
  For j=1 To count+1: Print #f, probeCase & "," & currentN & "," & i & "," & j & "," & ibno(i,j): Next j
 Next i
 Close #f
End Sub
Public Sub PrepareShared(ByVal path As String)
 Dim ws As Worksheet, pts As Variant, ids As Variant, i As Long, r As Long, j As Long
 Set ws=ThisWorkbook.Worksheets("図形定義")
 ws.Range("B2:C70000").ClearContents: ws.Range("E3:BZL4").ClearContents
 ws.Range("H1").Value2=6: ws.Range("J1").Value2=0: ws.Range("D3").Value2=2: ws.Range("D24").Value2=0
 pts=Array(Array(0#,0#),Array(4#,0#),Array(4#,4#),Array(0#,4#),Array(8#,0#),Array(8#,4#))
 For i=0 To 5: ws.Cells(i+2,2).Value2=pts(i)(0): ws.Cells(i+2,3).Value2=pts(i)(1): Next i
 For r=3 To 4
  If r=3 Then ids=Array(1,2,3,4) Else ids=Array(2,5,6,3)
  ws.Cells(r,5).Value2=4: ws.Cells(r,6).Value2=IIf(r=3,1#,.5)
  For j=0 To 3: ws.Cells(r,j+7).Value2=ids(j): Next j
 Next r
 currentN=6: probeCase="shared_reverse_subdivided": outputPath=path: Record "before_prepare"
 P5PrepareGeometryDefinitionWorkspace
 PINPUT NEX,NIN,ibex,ibin,ibno,NOB,NIB,px,py,delx,xlok
 DumpMesh path: Record "after_pinput": P5ReleaseGeometryDefinitionWorkspace
End Sub
'@)
    $memoryFile=Join-Path $Out 'memory.csv'; Set-Content $memoryFile 'case,n,label,timer,working_set_bytes,pagefile_bytes' -Encoding ASCII
    $helper.CodeModule.AddFromString("Public Sub SetProbePath(ByVal path As String)`r`n outputPath=path`r`nEnd Sub")
    $xl.VBE.CommandBars.FindControl(1,578).Execute()
    if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Compile failed.'}
    foreach($n in @(200,1000,1800)){
        [void]$xl.Run($prefix+'MeshMemoryProbe.SetProbePath',$memoryFile)
        [void]$xl.Run($prefix+'MeshMemoryProbe.PreparePolygon',[int]$n,$memoryFile)
    }
    [void]$xl.Run($prefix+'MeshMemoryProbe.PrepareShared',$memoryFile)
    @{status='PASS';book_sha256=(Get-FileHash $Book).Hash;compile='PASS';excel_version=$xl.Version;excel_build=$xl.Build;allocation=$beforeLabel;cases=@('polygon200','polygon1000','polygon1800','shared_reverse_subdivided')} | ConvertTo-Json | Set-Content (Join-Path $Out 'result.json') -Encoding utf8
    [IO.File]::WriteAllText((Join-Path $Out 'README.txt'),"Delaunay allocation probe completed. Compare rows before_sparse_tables/after_sparse_tables. This driver is for controlled Excel acceptance only.",[Text.Encoding]::UTF8)
} finally {
    if($wb){$wb.Close($false)}; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
}
