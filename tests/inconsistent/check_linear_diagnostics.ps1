param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'; New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null
$tmp=Join-Path ([IO.Path]::GetTempPath()) ('linear_diag_'+[guid]::NewGuid().ToString('N')+'.xlsm')
$before=(Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash; Copy-Item -LiteralPath $SourceWorkbook -Destination $tmp -Force
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1; $wb=$null
try {
  $wb=$xl.Workbooks.Open($tmp,0,$false); $cm=$wb.VBProject.VBComponents.Item('FEMSolver').CodeModule
  $cm.AddFromString(@'
Public Function LinearDiagHarness(ByVal caseId As Long) As Variant
  On Error GoTo Failed
  Dim i As Long, n As Long, r(0 To 4) As Variant
  SetP0SilentMode True
  n=2: nDof=n: lastDof=1: P6CSRNNZ=4: P6CSRReady=True: P6CSRILUReady=False
  ReDim P6CSRRowPtr(0 To n): ReDim P6CSRColumnIndex(0 To 3): ReDim P6CSRValues(0 To 3): ReDim P6CSRDiagonalPosition(0 To 1)
  P6CSRRowPtr(0)=0: P6CSRRowPtr(1)=2: P6CSRRowPtr(2)=4: P6CSRColumnIndex(0)=0: P6CSRColumnIndex(1)=1: P6CSRColumnIndex(2)=0: P6CSRColumnIndex(3)=1: P6CSRDiagonalPosition(0)=0: P6CSRDiagonalPosition(1)=3
  If caseId=1 Then
    P6CSRValues(0)=2#: P6CSRValues(1)=1#: P6CSRValues(2)=1#: P6CSRValues(3)=3#
  ElseIf caseId=2 Then
    P6CSRValues(0)=0#: P6CSRValues(1)=1#: P6CSRValues(2)=1#: P6CSRValues(3)=3#
  ElseIf caseId=3 Then
    P6CSRValues(0)=2#: P6CSRValues(1)=1#: P6CSRValues(2)=1#: P6CSRValues(3)=0.5#
  Else
    P6CSRValues(0)=1E+301: P6CSRValues(1)=1#: P6CSRValues(2)=1#: P6CSRValues(3)=3#
  End If
  r(0)=P6BuildCSRILU(): r(1)=P6CSRILUReady: r(2)=P6CSRILUFailRow: r(3)=P6CSRILUFailColumn: r(4)=P6CSRILUMinAbsDiag: LinearDiagHarness=r: Exit Function
Failed: LinearDiagHarness=Array(False,False,-99,-99,0#)
End Function
'@)
  $p="'"+$wb.Name+"'!"; $rows=@(); foreach($id in 1,2,3,4){$v=@($xl.Run($p+'LinearDiagHarness',$id)); if($id -eq 1 -and (-not [bool]$v[0] -or -not [bool]$v[1])){throw 'well-conditioned ILU did not pass'}; if($id -ge 2 -and [bool]$v[1]){throw 'invalid ILU factor was accepted'}; $rows += [ordered]@{case=$id;build_ok=[bool]$v[0];ilu_ready=[bool]$v[1];fail_row=$v[2];fail_column=$v[3];min_abs_diag=$v[4]}}
  [IO.File]::WriteAllText((Join-Path $OutputRoot 'check_linear_diagnostics.json'),([ordered]@{status='PASS';source_sha256=$before;results=$rows}|ConvertTo-Json -Depth 6))
} finally {if($wb){$wb.Close($false)};$xl.Quit();[Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null;Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue}
if((Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash -ne $before){throw 'SourceWorkbook changed'}
