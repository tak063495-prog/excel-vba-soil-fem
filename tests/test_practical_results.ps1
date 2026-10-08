$ErrorActionPreference='Stop'
$rootTask=Split-Path -Parent $PSScriptRoot
Push-Location $rootTask
try {
$copyTask=Join-Path $pwd 'tests/tmp/practical_results.xlsm'
Copy-Item workbook/2DSoilFEM_20261008_practical.xlsm $copyTask -Force
$xlTask=New-Object -ComObject Excel.Application
try {
  $xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1
  $wbTask=$xlTask.Workbooks.Open($copyTask,0,$false)
  $tmTask=$wbTask.VBProject.VBComponents.Add(1);$tmTask.Name='PracticalResultTest'
  $tmTask.CodeModule.AddFromString(@'
Option Explicit
Private passed As Long
Private Sub Assert(ByVal condition As Boolean, ByVal label As String)
  If Not condition Then Err.Raise vbObjectError + 3951, "PracticalResultTest", label
  passed = passed + 1
End Sub
Private Sub Near(ByVal actual As Double, ByVal expected As Double, ByVal label As String)
  Assert Abs(actual - expected) < 0.000000001 * (1# + Abs(expected)), label
End Sub
Public Function TestResults() As String
  Dim g As Long, i As Long, j As Long, ws As Worksheet
  Dim gx As Variant, gy As Variant, status As String, areaRejected As Boolean
  On Error GoTo Failed
  NumberOfElement = 3: NumberOfFreeNode = 14: NumberOfNode = 14: NumberOfMaterial = 3: lastDof = 27
  ReDim Elem(0 To 2): ReDim Material(0 To 2): ReDim TDisp(0 To 27)
  ReDim P3ElementActive(0 To 2): P3ElementActiveReady = True
  P3ElementActive(0) = True: P3ElementActive(1) = False: P3ElementActive(2) = True
  Material(0).thickness = 1#: Material(1).thickness = 1#: Material(2).thickness = 1#
  gx = Array(0#, 1#, 1#, 0#, 0.5, 1#, 0.5, 0#): gy = Array(0#, 0#, 1#, 1#, 0#, 0.5, 1#, 0.5)
  For i = 0 To 1
    Elem(i).MatNo = i
    For j = 0 To 7
      Elem(i).node(j) = j: Elem(i).x(j) = gx(j): Elem(i).y(j) = gy(j)
    Next j
    For g = 0 To 3
      Elem(i).dj(g) = 0.25
      If i = 0 Then
        Elem(i).Stmat(0, g) = 10#: Elem(i).Stmat(1, g) = 20#: Elem(i).Stmat(2, g) = 5#
      End If
    Next g
  Next i
  Elem(2).IsJoint = True: Elem(2).MatNo = 2
  For j = 0 To 5
    Elem(2).node(j) = j + 8: Elem(2).x(j) = CDbl(j): Elem(2).y(j) = 7#
  Next j
  Elem(2).node(6) = -1: Elem(2).node(7) = -1
  AnalysisOK = True: ModelRevision = 7
  BuildP1Results
  For j = 0 To 7
    Near NodalGlobalResult(j, 0), 10#, "active stress not diluted by DEATH element"
    Near NodalGlobalResult(j, 1), 20#, "active Y stress"
    Near NodalGlobalResult(j, 2), 5#, "active shear"
    Near NodalResultWeight(j), 1#, "active continuum weight only"
  Next j
  For j = 8 To 13
    Near NodalResultWeight(j), 0#, "joint nodes excluded from continuum smoothing"
    Near NodalX(j), CDbl(j - 8), "unweighted node X retained"
    Near NodalY(j), 7#, "unweighted node Y retained"
  Next j
  Assert P1ResultReady And P1ResultModelRevision = 7, "P1 readiness"
  For j = 0 To 7: Elem(0).x(j) = 0#: Elem(0).y(j) = 0#: Next j
  On Error Resume Next
  BuildP1Results
  areaRejected = (Err.Number <> 0): Err.Clear
  On Error GoTo Failed
  Assert areaRejected, "zero area active continuum still rejected"

  NumberOfElement = 1: NumberOfMaterial = 1: NumberOfNode = 6: NumberOfFreeNode = 6: lastDof = 11
  ReDim Elem(0 To 0): ReDim Material(0 To 0): ReDim TDisp(0 To 11): ReDim P3ElementActive(0 To 0)
  P3ElementActiveReady = True: P3ElementActive(0) = True
  Elem(0).IsJoint = True: Elem(0).MatNo = 0
  For j = 0 To 5: Elem(0).node(j) = j: Next j
  Elem(0).node(6) = -1: Elem(0).node(7) = -1
  Elem(0).x(0) = 0#: Elem(0).x(1) = 1#: Elem(0).x(2) = 0.5
  Elem(0).x(3) = 0#: Elem(0).x(4) = 1#: Elem(0).x(5) = 0.5
  Material(0).Kn = 100#: Material(0).Young = 100#: Material(0).Ks = 50#: Material(0).thickness = 1#
  Material(0).cohesion = 1#: Material(0).Kind = "JOINT": Material(0).AllowTension = False
  For j = 3 To 5: TDisp(2 * j) = 0.02: TDisp(2 * j + 1) = -0.01: Next j
  P3JointEval 0, False
  For g = 0 To 2
    Assert Elem(0).JointContact(g) = 1, "opening detaches contact"
    Near Elem(0).Stmat(0, g), 0#, "detached normal force"
    Near Elem(0).Stmat(1, g), 0#, "detached shear force"
    Near Elem(0).Stmat(2, g), 0.01, "positive opening"
  Next g
  For j = 3 To 5: TDisp(2 * j) = 0.01: TDisp(2 * j + 1) = 0.01: Next j
  P3JointEval 0, False
  For g = 0 To 2
    Assert Elem(0).JointContact(g) = 0, "compression sticks"
    Near Elem(0).Stmat(0, g), 1#, "compression positive"
    Near Elem(0).Stmat(1, g), 0.5, "shear Ks times slip"
    Near Elem(0).Stmat(2, g), -0.01, "compression closes gap"
  Next g
  For j = 3 To 5: TDisp(2 * j) = 0.1: Next j
  P3JointEval 0, False
  For g = 0 To 2
    Assert Elem(0).JointContact(g) = 2, "shear reaches slip state"
    Near Elem(0).Stmat(1, g), 1#, "Coulomb shear capped"
    Near Elem(0).Stmat(10, g), 0.1, "slip retained"
  Next g
  Material(0).AllowTension = True
  For j = 3 To 5: TDisp(2 * j) = 0#: TDisp(2 * j + 1) = -0.01: Next j
  P3JointEval 0, False
  For g = 0 To 2
    Assert Elem(0).JointContact(g) = 0, "tension allowed remains bonded"
    Near Elem(0).Stmat(0, g), -1#, "tension negative"
  Next g
  FEMWriteTextSetting "OUTPUT_STAGE_MODE", "FINAL"
  FEMWriteTextSetting "OUTPUT_INACTIVE_ELEMENTS", "SKIP"
  P3LastCompletedStage = 1
  SaveStress
  Set ws = ThisWorkbook.Worksheets("接合結果")
  Assert ws.Cells(ws.Rows.Count, 1).End(xlUp).Row = 4, "three joint Gauss rows"
  Assert ThisWorkbook.Worksheets("結果要素").Cells(2, 5).Value2 = "接合結果参照", "joint stress marker"
  For g = 0 To 2
    Near CDbl(ws.Cells(g + 2, 7).Value2), Elem(0).Stmat(0, g), "export normal traction"
    Near CDbl(ws.Cells(g + 2, 8).Value2), Elem(0).Stmat(1, g), "export shear traction"
    Near CDbl(ws.Cells(g + 2, 9).Value2), Elem(0).Stmat(2, g), "export gap without CalcStrain overwrite"
    Near CDbl(ws.Cells(g + 2, 10).Value2), Elem(0).Stmat(10, g), "export slip without overwrite"
  Next g
  SaveStress
  Assert ws.Cells(ws.Rows.Count, 1).End(xlUp).Row = 4, "FINAL overwrites rather than appends"
  FEMWriteTextSetting "OUTPUT_STAGE_MODE", "ALL"
  FEMPracticalClearJointResults
  P3LastCompletedStage = 1: FEMPracticalSaveJointResults
  P3LastCompletedStage = 2: FEMPracticalSaveJointResults
  Assert ws.Cells(ws.Rows.Count, 1).End(xlUp).Row = 7, "ALL appends three points per stage"
  Assert ws.Cells(2, 1).Value2 = 1 And ws.Cells(5, 1).Value2 = 2, "ALL keeps stage numbers"
  FEMWriteTextSetting "OUTPUT_STAGE_MODE", "FINAL"
  P3ElementActive(0) = False
  FEMPracticalSaveJointResults
  Assert ws.Cells(ws.Rows.Count, 1).End(xlUp).Row = 1, "SKIP removes inactive joint rows"
  FEMWriteTextSetting "OUTPUT_INACTIVE_ELEMENTS", "INCLUDE"
  FEMPracticalSaveJointResults
  Assert ws.Cells(ws.Rows.Count, 1).End(xlUp).Row = 4, "INCLUDE keeps inactive joint metadata"
  For g = 0 To 2
    Assert IsEmpty(ws.Cells(g + 2, 7).Value2) And IsEmpty(ws.Cells(g + 2, 10).Value2), "inactive joint has no false zero results"
    Assert ws.Cells(g + 2, 12).Value2 = "無効" And ws.Cells(g + 2, 13).Value2 = 0, "inactive joint labeled"
  Next g
  Cleardata
  Assert ws.Cells(ws.Rows.Count, 1).End(xlUp).Row = 1, "clear removes previous joint results"
  TestResults = "PASS|" & CStr(passed)
  Exit Function
Failed:
  TestResults = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
'@)
  $resultTask=$xlTask.Run("'"+$wbTask.Name+"'!TestResults")
  Write-Output $resultTask
  if(-not $resultTask.StartsWith('PASS|')){throw $resultTask}
  $resultTask|Set-Content tests/results/practical_result_checks_20261008.txt -Encoding UTF8
  $wbTask.Close($false)
} finally {
  if($null -ne $wbTask){try{$wbTask.Close($false)}catch{}}
  try{$xlTask.Quit()}catch{}
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)|Out-Null
}

} finally { Pop-Location }
