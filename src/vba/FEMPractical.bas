Option Explicit
Public Const FEM_PRACTICAL_STAMP As String = "20261009_PRACTICAL_04"

Public Sub FEMPracticalAfterLayout()
  Dim ws As Worksheet, shp As shape, target As range
  Set ws = ThisWorkbook.Worksheets("操作パネル")
  ws.range("B35:H35").Merge
  ws.range("B35").value2 = "実務動作修正版 " & FEM_PRACTICAL_STAMP & " / INCONSISTENT接線・非収束判定修正"
  ws.range("B35").Font.size = 10
  ws.range("B35").Font.Color = RGB(97, 112, 128)
  On Error Resume Next
  Set shp = ws.Shapes("UI_JointResults")
  On Error GoTo 0
  Set target = ws.range("F33:G33")
  If shp Is Nothing Then
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, target.left + 2, target.Top + 2, target.width - 4, target.Height - 4)
    shp.name = "UI_JointResults"
  End If
  shp.TextFrame2.TextRange.text = "接合結果"
  shp.TextFrame2.TextRange.Font.name = "Arial"
  shp.TextFrame2.TextRange.Font.size = 11
  shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
  shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
  shp.fill.ForeColor.RGB = RGB(231, 237, 244)
  shp.TextFrame2.TextRange.Font.fill.ForeColor.RGB = RGB(36, 52, 71)
  shp.line.Visible = msoFalse
  shp.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!FEMPracticalShowJointResults"
  ThisWorkbook.Worksheets("接合結果").Move After:=ThisWorkbook.Worksheets("ステージ結果")
  ThisWorkbook.Worksheets("接合結果").Visible = xlSheetVisible
End Sub

Public Sub FEMPracticalShowJointResults()
  With ThisWorkbook.Worksheets("接合結果")
    .Visible = xlSheetVisible: .Activate
    Application.Goto .range("A1"), True
  End With
End Sub

Public Sub FEMPracticalClearJointResults()
  Dim ws As Worksheet, lastRow As Long
  Set ws = ThisWorkbook.Worksheets("接合結果")
  lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
  If lastRow >= 2 Then ws.range(ws.Cells(2, 1), ws.Cells(lastRow, 15)).ClearContents
End Sub

Public Sub FEMPracticalSaveJointResults()
  Dim ws As Worksheet, e As Long, gp As Long, a As Long, rowNo As Long, count As Long
  Dim startRow As Long, stageNo As Long, xi As Double, xValue As Double, yValue As Double
  Dim buffer() As Variant, shape(0 To 2) As Double, xiGauss(0 To 2) As Double
  Dim active As Boolean, skipInactive As Boolean, appendMode As Boolean, kn As Double, ks As Double
  Set ws = ThisWorkbook.Worksheets("接合結果")
  appendMode = (UCase$(Trim$(FEMReadTextSetting("OUTPUT_STAGE_MODE", "FINAL"))) = "ALL")
  skipInactive = (UCase$(Trim$(FEMReadTextSetting("OUTPUT_INACTIVE_ELEMENTS", "SKIP"))) <> "INCLUDE")
  If appendMode Then
    startRow = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    If startRow < 2 Then startRow = 2
  Else
    FEMPracticalClearJointResults
    startRow = 2
  End If
  For e = 0 To NumberOfElement - 1
    If Elem(e).IsJoint Then
      If P3IsElementActive(e) Or Not skipInactive Then count = count + 3
    End If
  Next e
  If count = 0 Then Exit Sub
  If startRow + count - 1 > ws.rows.count Then Err.Raise vbObjectError + 3901, "FEMPracticalSaveJointResults", "接合結果がExcelの行数上限を超えます。"
  ReDim buffer(1 To count, 1 To 15)
  xiGauss(0) = -0.774596669241483: xiGauss(1) = 0#: xiGauss(2) = 0.774596669241483
  stageNo = P3ResultStageNumber()
  For e = 0 To NumberOfElement - 1
    If Elem(e).IsJoint Then
      active = P3IsElementActive(e)
      If active Or Not skipInactive Then
        kn = Material(Elem(e).MatNo).kn
        If kn <= 0# Then kn = Material(Elem(e).MatNo).young
        ks = Material(Elem(e).MatNo).ks
        If ks <= 0# Then ks = 0.1 * kn
        For gp = 0 To 2
          rowNo = rowNo + 1: xi = xiGauss(gp)
          shape(0) = 0.5 * xi * (xi - 1#): shape(1) = 0.5 * xi * (xi + 1#): shape(2) = 1# - xi * xi
          xValue = 0#: yValue = 0#
          For a = 0 To 2
            xValue = xValue + shape(a) * Elem(e).x(a): yValue = yValue + shape(a) * Elem(e).y(a)
          Next a
          buffer(rowNo, 1) = stageNo: buffer(rowNo, 2) = e + 1
          buffer(rowNo, 3) = Elem(e).MatNo + 1: buffer(rowNo, 4) = gp + 1
          buffer(rowNo, 5) = xValue: buffer(rowNo, 6) = yValue
          If active Then
            buffer(rowNo, 7) = Elem(e).Stmat(0, gp): buffer(rowNo, 8) = Elem(e).Stmat(1, gp)
            buffer(rowNo, 9) = Elem(e).Stmat(2, gp): buffer(rowNo, 10) = Elem(e).Stmat(10, gp)
            buffer(rowNo, 11) = Elem(e).JointContact(gp)
            Select Case Elem(e).JointContact(gp)
              Case 0: buffer(rowNo, 12) = "密着"
              Case 1: buffer(rowNo, 12) = "剥離"
              Case 2: buffer(rowNo, 12) = "すべり"
              Case Else: buffer(rowNo, 12) = "未設定"
            End Select
          Else
            buffer(rowNo, 12) = "無効"
          End If
          buffer(rowNo, 13) = Abs(CLng(active)): buffer(rowNo, 14) = kn: buffer(rowNo, 15) = ks
        Next gp
      End If
    End If
  Next e
  ws.range(ws.Cells(startRow, 1), ws.Cells(startRow + count - 1, 15)).value2 = buffer
End Sub
