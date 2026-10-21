Option Explicit

' Synthetic observer test, not a constitutive or solver convergence test.
Public Function ManualReviewObserverProbe() As String
  Dim gp As Long, c As Long, j As Long, mode As Variant, saved As Double
  Dim ws As Worksheet, d As Object, text As String, legend As String
  Dim phase As String
  On Error GoTo Failed
  phase = "initialize"
  NumberOfElement = 1: ReDim Elem(0): lastDof = 15
  ReDim TDisp(15): ReDim P3CommittedDisp(15)
  P3ElementActiveReady = False: P3LastCompletedStage = 1
  For j = 0 To 15: Elem(0).ElNode(j) = j: Next j
  For gp = 0 To 3
    Elem(0).Bmat(0, 0, gp) = -1#
    For c = 0 To 3: Elem(0).P2PlasticStrain(c, gp) = (c + 1) * 0.001: Next c
  Next gp
  FEMIncrementReset
  FEMIncrementSave
  Set ws = ThisWorkbook.Worksheets("結果増分")
  FEMWriteTextSetting "OUTPUT_STAGE_MODE", "FINAL"
  phase = "settings"
  For Each mode In Array("OFF", "AUTO")
    FEMWriteTextSetting "NATIVE_DLL", CStr(mode)
    phase = "capture nonzero"
    TDisp(0) = 0.01: CurrentIncrement = 1
    FEMIncrementCapture 1#: Call FEMIncrementSnapshot: FEMIncrementSave
    If ws.Cells(2, 9).Value2 <> 0.001 Then Err.Raise 5, , "Nonzero plastic increment lost"
    If ws.Cells(2, 6).Value2 <> -0.01 Then Err.Raise 5, , "Affine expansion must be compression-negative"
    saved = Elem(0).P2PlasticStrain(0, 0)
    phase = "capture zero"
    TDisp(0) = 1E-31: CurrentIncrement = 2
    FEMIncrementCapture 1#: FEMIncrementSave
    For gp = 2 To 5
      For c = 9 To 16
        If c <> 13 And ws.Cells(gp, c).Value2 <> 0# Then Err.Raise 5, , "Stale plastic observer value"
      Next c
    Next gp
    If Elem(0).P2PlasticStrain(0, 0) <> saved Then Err.Raise 5, , "Material history modified"
    Call FEMIncrementRestore: FEMIncrementSave
    If ws.Cells(2, 9).Value2 <> saved Or ws.Cells(2, 19).Value2 <> 1 Then Err.Raise 5, , "Observer snapshot changed plastic=" & CStr(ws.Cells(2, 9).Value2) & " inc=" & CStr(ws.Cells(2, 19).Value2) & " expected=" & CStr(saved)
    TDisp(0) = 1E-29: CurrentIncrement = 3
    FEMIncrementCapture 1#: FEMIncrementSave
    If ws.Cells(2, 9).Value2 <> saved Then Err.Raise 5, , "Nonzero threshold value suppressed"
  Next mode
  phase = "SVG"
  Set d = CreateObject("Scripting.Dictionary")
  d(1) = Array(0#, 0#, 100#, 0#, 100#, 100#, 0#, 100#, 0#, 0.001, 0#, 0#)
  text = FEMIncrementOverlaySvg(d, False, True, False, 1#, legend)
  If InStr(text, "#1f77b4") = 0 Or InStr(text, "x2='50") = 0 Then Err.Raise 5, , "Compression must be blue and inward"
  ManualReviewObserverProbe = "PASS: zero/nonzero threshold, history unchanged, snapshot, compression sign/color/inward arrow"
  Exit Function
Failed:
  ManualReviewObserverProbe = "FAIL: " & phase & " " & Err.Number & " " & Err.Description
End Function
