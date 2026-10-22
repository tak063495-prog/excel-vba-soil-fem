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
  shp.left = target.left + 2: shp.Top = target.Top + 2
  shp.width = target.width - 4: shp.Height = target.Height - 4
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
  lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row
  If lastRow >= 2 Then ws.range(ws.cells(2, 1), ws.cells(lastRow, 20)).ClearContents
End Sub

Public Sub FEMPracticalSaveJointResults(Optional ByVal restoredStage As Long = 0)
  Dim ws As Worksheet, e As Long, gp As Long, a As Long, rowNo As Long, count As Long
  Dim startRow As Long, stageNo As Long, xi As Double, xValue As Double, yValue As Double
  Dim buffer() As Variant, shape(0 To 2) As Double, xiGauss(0 To 2) As Double
  Dim active As Boolean, skipInactive As Boolean, appendMode As Boolean, kn As Double, ks As Double
  Set ws = ThisWorkbook.Worksheets("接合結果")
  FEMPracticalEnsureJointLayout
  appendMode = (UCase$(Trim$(FEMReadTextSetting("OUTPUT_STAGE_MODE", "FINAL"))) = "ALL")
  skipInactive = (UCase$(Trim$(FEMReadTextSetting("OUTPUT_INACTIVE_ELEMENTS", "SKIP"))) <> "INCLUDE")
  If appendMode Then
    startRow = ws.cells(ws.rows.count, 1).End(xlUp).row + 1
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
  ReDim buffer(1 To count, 1 To 20)
  xiGauss(0) = -0.774596669241483: xiGauss(1) = 0#: xiGauss(2) = 0.774596669241483
  stageNo = P3ResultStageNumber()
  If restoredStage > 0 Then stageNo = restoredStage
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
            buffer(rowNo, 16) = GWJointPressure(e, gp)
            buffer(rowNo, 17) = Elem(e).Stmat(0, gp) + GWJointPressure(e, gp)
            buffer(rowNo, 18) = Elem(e).Stmat(0, gp)
            buffer(rowNo, 19) = Material(Elem(e).MatNo).cohesion + WorksheetFunction.Max(0#, Elem(e).Stmat(0, gp)) * Tan(Material(Elem(e).MatNo).fai * 3.14159265358979 / 180#)
            If Elem(e).JointContact(gp) = 1 Then buffer(rowNo, 19) = 0#
            buffer(rowNo, 20) = GJSource(e)
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
  ws.range(ws.cells(startRow, 1), ws.cells(startRow + count - 1, 20)).value2 = buffer
End Sub

Public Sub FEMPracticalEnsureJointLayout()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets("接合結果")
  If CStr(ws.range("P1").value2) <> "接合面水圧[kPa]" Then ws.range("Q1:R5").ClearContents
  ws.range("P1:T1").value2 = Array("接合面水圧[kPa]", "全法線応力[kPa]", "有効接触応力[kPa]", "摩擦上限[kPa]", "水圧元")
  ws.columns("P:T").ColumnWidth = 22
  ws.range("V1").value2 = "接合の結果": ws.range("W1").value2 = "3積分点。G/R列は固体の有効接触応力、P列は物理水圧、Q列は両者の和。圧縮を正。"
  ws.range("V2").value2 = "接触状態": ws.range("W2").value2 = "0=密着、1=開離、2=すべり。開離中でも設定した水圧は作用します。"
  ws.range("V3").value2 = "変位・強度": ws.range("W3").value2 = "I列は開口を正、J列は相対すべり。S列は現在Fsで低減した摩擦上限。"
  ws.range("V4").value2 = "水圧元": ws.range("W4").value2 = "T列はSOIL/CONTINUOUS/PRESCRIBED。地盤間水頭だけを連続にしても変位自由度は分離。"
  ws.range("V5").value2 = "保存範囲": ws.range("W5").value2 = "出力・表示のFINAL/ALL、SKIP/INCLUDEに従います。水条件NONEではP列0、Q=R=従来G。"
  ws.columns("V").ColumnWidth = 22: ws.columns("W").ColumnWidth = 80
  ws.range("P1:T1").WrapText = True: ws.range("V1:W5").WrapText = True
  ws.rows(1).RowHeight = 42
End Sub
