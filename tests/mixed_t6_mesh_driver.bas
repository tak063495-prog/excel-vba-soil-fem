Option Explicit

Public Function M6SetFallback(ByVal enabled As Long) As String
  On Error GoTo Failed
  FEMWriteNumericSetting "MESH_T6_FALLBACK", CDbl(enabled)
  FEMInvalidateSettingCache
  M6SetFallback = "PASS|" & CStr(FEMReadSetting("MESH_T6_FALLBACK", 0#))
  Exit Function
Failed: M6SetFallback = "ERROR|" & Err.Description
End Function

Public Function M6ApplyFallback() As String
  On Error GoTo Failed
  P5ApplyT6FallbackCurrentSheet
  M6ApplyFallback = "PASS"
  Exit Function
Failed: M6ApplyFallback = "ERROR|" & Err.Description
End Function

Public Function M6Validate(ByVal expectedQ8 As Long, ByVal expectedT6 As Long, ByVal expectedNodes As Long) As String
  Dim ws As Worksheet, lastRow As Long, r As Long, q8 As Long, t6 As Long, typ As String
  On Error GoTo Failed
  Set ws = ThisWorkbook.Worksheets("要素データ")
  lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
  For r = 2 To lastRow
    typ = UCase$(Trim$(CStr(ws.Cells(r, 11).Value2)))
    If typ = "T6" Then t6 = t6 + 1 Else q8 = q8 + 1
  Next r
  If expectedQ8 >= 0 And q8 <> expectedQ8 Then Err.Raise vbObjectError + 3981, , "Q8件数不一致=" & CStr(q8)
  If expectedT6 >= 0 And t6 <> expectedT6 Then Err.Raise vbObjectError + 3982, , "T6件数不一致=" & CStr(t6)
  If expectedNodes >= 0 Then
    If ThisWorkbook.Worksheets("節点データ").Cells(ThisWorkbook.Worksheets("節点データ").Rows.Count, 1).End(xlUp).Row - 1 <> expectedNodes Then Err.Raise vbObjectError + 3983, , "節点数不一致"
  End If
  P5EvaluateCurrentSheet "T6受入検査"
  ValidateInputData
  M6Validate = "PASS|Q8=" & CStr(q8) & "|T6=" & CStr(t6)
  Exit Function
Failed: M6Validate = "ERROR|" & Err.Description
End Function

Public Function M6Preview() As String
  On Error GoTo Failed
  SuppressUserMessages = True
  FEMViewAll
  If Len(FEMViewerLastError) > 0 Then Err.Raise vbObjectError + 3984, , FEMViewerLastError
  If FEMMeshSheetHasT6() And InStr(FEMViewerLastSvg, "#008b6d") = 0 Then Err.Raise vbObjectError + 3984, , "T6の輪郭表示がありません。"
  If ThisWorkbook.Worksheets("図").Shapes.count = 0 Then Err.Raise vbObjectError + 3984, , "図形がありません。"
  M6Preview = "PASS|SVG=" & CStr(Len(FEMViewerLastSvg))
  Exit Function
Failed: M6Preview = "ERROR|" & Err.Description
End Function

Public Function M6Quiet() As String
  SuppressUserMessages = True
  FEM24Error = "": FEMViewerLastError = ""
  M6Quiet = "PASS"
End Function

Public Function M6ReadFallback() As String
  FEMInvalidateSettingCache
  M6ReadFallback = CStr(FEMReadSetting("MESH_T6_FALLBACK", -1#))
End Function

Public Function M6Run() As String
  On Error GoTo Failed
  SuppressUserMessages = True
  P0_RunAnalysis
  If Not AnalysisOK Then Err.Raise vbObjectError + 3985, , AnalysisMessage
  If Abs(TotalSelfWeight - 45) > 0.000001 Or Abs(P6ReactionSumY - 45) > 0.000001 Then Err.Raise vbObjectError + 3985, , "自重・反力の独立検算が不一致です。"
  If RelativeResidualFree > 0.0000001 Then Err.Raise vbObjectError + 3985, , "残差が大きすぎます。"
  FEMViewResult
  If Len(FEMViewerLastError) > 0 Then Err.Raise vbObjectError + 3985, , FEMViewerLastError
  M6Run = "PASS|weight=" & CStr(TotalSelfWeight) & "|reaction=" & CStr(P6ReactionSumY) & "|residual=" & CStr(RelativeResidualFree)
  Exit Function
Failed: M6Run = "ERROR|" & Err.Description
End Function

Public Function M6Generate(ByVal fallback As Long, ByVal invalid As Boolean, Optional ByVal method As Long = 2) As String
  On Error GoTo Failed
  Dim ws As Worksheet, pts As Variant, i As Long
  SuppressUserMessages = True
  Set ws = ThisWorkbook.Worksheets("形状入力")
  ws.Range("B8:D2000").ClearContents: ws.Range("G8:J2000").ClearContents
  pts = Array(Array(1, 0, 0), Array(2, 2, 0), Array(3, 2, 1), Array(4, 0, 1))
  For i = 0 To 3
    ws.Cells(i + 8, 2).Value2 = pts(i)(0): ws.Cells(i + 8, 3).Value2 = pts(i)(1): ws.Cells(i + 8, 4).Value2 = pts(i)(2)
  Next i
  ws.Cells(8, 7).Value2 = 1: ws.Cells(8, 8).Value2 = "外側"
  ws.Cells(8, 9).Value2 = IIf(invalid, "1,3,2,4", "1,2,3,4"): ws.Cells(8, 10).Value2 = 1#
  FEMWriteNumericSetting "MESH_GEN_COMPARE", CDbl(method)
  FEMWriteNumericSetting "MESH_T6_FALLBACK", CDbl(fallback)
  If fallback <> 0 Then FEMWriteNumericSetting "MESH_QUALITY_MIN_JACOBIAN_RATIO", 0.03 Else FEMWriteNumericSetting "MESH_QUALITY_MIN_JACOBIAN_RATIO", 0.000001
  FEMUiGenerateShape
  If Not FEMShapeGenerationOK Or Len(FEMUiMeshPreviewError) > 0 Then Err.Raise vbObjectError + 3986, , FEMUiMeshPreviewError
  M6Generate = "PASS"
  Exit Function
Failed: M6Generate = "ERROR|" & Err.Description
End Function

Public Function M6ShuffledRows() As String
  Dim ws As Worksheet, saved As Variant, temp As Variant, expected As Double, result As String
  On Error GoTo Failed
  Set ws = ThisWorkbook.Worksheets("要素データ")
  saved = ws.Range("A2:K4").Value2
  temp = ws.Range("A2:K2").Value2
  ws.Range("A2:K2").Value2 = ws.Range("A3:K3").Value2
  ws.Range("A3:K3").Value2 = temp
  expected = (1# + 1# / Sqr(3#)) / 4#
  If Abs(FEMGaussFraction(0, 0, True) - 0.25) > 0.000000000001 Then Err.Raise vbObjectError + 3987, , "Q8保存積分重みが不正です。"
  If Abs(FEMGaussFraction(1, 0, True) - expected) > 0.000000000001 Then Err.Raise vbObjectError + 3987, , "T6保存積分重みが不正です。"
  FEMInvalidateRemeshedModel
  result = M6Run()
  If InStr(result, "連番") = 0 Or Left$(result, 6) <> "ERROR|" Then Err.Raise vbObjectError + 3987, , "行順不正が拒否されません。" & result
  ws.Range("A2:K4").Value2 = saved
  FEMInvalidateRemeshedModel
  M6ShuffledRows = "PASS"
  Exit Function
Failed:
  result = Err.Description
  On Error Resume Next
  ws.Range("A2:K4").Value2 = saved
  FEMInvalidateRemeshedModel
  M6ShuffledRows = "ERROR|" & result
End Function

Public Function M6PrimeQualityScale() As String
  Dim xx(1 To 8) As Double, yy(1 To 8) As Double, conn(1 To 1, 1 To 8) As Long, j As Long
  On Error GoTo Failed
  xx(2) = 100#: xx(3) = 100#: yy(3) = 100#: yy(4) = 100#
  xx(5) = 50#: xx(6) = 100#: yy(6) = 50#: xx(7) = 50#: yy(7) = 100#: yy(8) = 50#
  For j = 1 To 8: conn(1, j) = j: xx(j) = xx(j) * 10000000#: yy(j) = yy(j) * 10000000#: Next j
  P5ValidateMeshQuality 8, 1, conn, xx, yy
  M6PrimeQualityScale = "PASS"
  Exit Function
Failed: M6PrimeQualityScale = "ERROR|" & Err.Description
End Function
