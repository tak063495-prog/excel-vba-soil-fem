Option Explicit
' Input cache, element kinematics/stiffness, results and reports.
Private Sub ClearDataRows(ByVal sheetName As String, ByVal firstColumn As Long, ByVal lastColumn As Long)
  Dim ws As Worksheet
  Dim lastRow As Long
  Set ws = ThisWorkbook.Worksheets(sheetName)
  lastRow = ws.Cells(ws.rows.count, firstColumn).End(xlUp).row
  If lastRow >= 2 Then
    ws.range(ws.Cells(2, firstColumn), ws.Cells(lastRow, lastColumn)).ClearContents
  End If
End Sub

Sub Cleardata()
  ClearDataRows "結果節点", 1, 12
  ClearDataRows "結果要素", 1, 80
  FEMPracticalClearJointResults
End Sub
Sub ClearEle()
  ClearDataRows "節点データ", 1, 14
  ClearDataRows "要素データ", 1, 10
  FEMInvalidateInputCache
End Sub
Sub DrawElement(Eno, XX1, YY1, XX2, YY2, XX3, YY3, XX4, YY4, XX5, YY5, XX6, YY6, XX7, YY7, XX8, YY8)
Dim x1 As Variant: Dim x2 As Variant: Dim x3 As Variant: Dim x4 As Variant: Dim x5 As Variant
Dim x6 As Variant: Dim x7 As Variant: Dim x8 As Variant: Dim y1 As Variant: Dim y2 As Variant
Dim y3 As Variant: Dim y4 As Variant: Dim y5 As Variant: Dim y6 As Variant: Dim y7 As Variant
Dim y8 As Variant
Dim xmin As Variant: Dim yMax As Variant: Dim sc As Variant
'メッシュ＆メッシュ番号描写
  xmin = P6ReadSetting("VIEW_MIN_X", 0#)
  yMax = P6ReadSetting("VIEW_MAX_Y", 0#)
  sc = P6ReadSetting("VIEW_SCALE", 1#)
  x1 = xmin + XX1 * sc: y1 = yMax - YY1 * sc '3点の座標変換
  x2 = xmin + XX2 * sc: y2 = yMax - YY2 * sc
  x3 = xmin + XX3 * sc: y3 = yMax - YY3 * sc
  x4 = xmin + XX4 * sc: y4 = yMax - YY4 * sc
  x5 = xmin + XX5 * sc: y5 = yMax - YY5 * sc
  x6 = xmin + XX6 * sc: y6 = yMax - YY6 * sc
  x7 = xmin + XX7 * sc: y7 = yMax - YY7 * sc
  x8 = xmin + XX8 * sc: y8 = yMax - YY8 * sc
  
  With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, x1, y1)
       .AddNodes msoSegmentLine, msoEditingAuto, x1, y1
       .AddNodes msoSegmentLine, msoEditingAuto, x2, y2
       .AddNodes msoSegmentLine, msoEditingAuto, x3, y3
       .AddNodes msoSegmentLine, msoEditingAuto, x4, y4
       .AddNodes msoSegmentLine, msoEditingAuto, x5, y5
       .AddNodes msoSegmentLine, msoEditingAuto, x6, y6
       .AddNodes msoSegmentLine, msoEditingAuto, x7, y7
       .AddNodes msoSegmentLine, msoEditingAuto, x8, y8
       .AddNodes msoSegmentLine, msoEditingAuto, x1, y1
       .ConvertToShape.Select
  End With
  selection.name = "E_" & Eno '図形に名をつける
  selection.ShapeRange.fill.Visible = msoFalse '塗潰し無し
  selection.ShapeRange.line.ForeColor.SchemeColor = 0
  
  'XM = (X1 + X2 + X3) / 3 - 5: YM = (Y1 + Y2 + Y3) / 3 - 5
  'ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, XM, YM, 30, 20).Select
  With selection
     .ShapeRange.fill.Visible = msoFalse
     '.ShapeRange.Line.Visible = msoFalse
     .ShapeRange.line.ForeColor.SchemeColor = 0
  End With
End Sub
Sub DrawDispElement(Eno, XX1, YY1, XX2, YY2, XX3, YY3, XX4, YY4, XX5, YY5, XX6, YY6, XX7, YY7, XX8, YY8)
Dim x1 As Variant: Dim x2 As Variant: Dim x3 As Variant: Dim x4 As Variant: Dim x5 As Variant
Dim x6 As Variant: Dim x7 As Variant: Dim x8 As Variant: Dim y1 As Variant: Dim y2 As Variant
Dim y3 As Variant: Dim y4 As Variant: Dim y5 As Variant: Dim y6 As Variant: Dim y7 As Variant
Dim y8 As Variant
Dim xmin As Variant: Dim yMax As Variant: Dim sc As Variant
'変位用要素の描写
  xmin = P6ReadSetting("VIEW_MIN_X", 0#)
  yMax = P6ReadSetting("VIEW_MAX_Y", 0#)
  sc = P6ReadSetting("VIEW_SCALE", 1#)
  x1 = xmin + XX1 * sc: y1 = yMax - YY1 * sc '3点の座標変換
  x2 = xmin + XX2 * sc: y2 = yMax - YY2 * sc
  x3 = xmin + XX3 * sc: y3 = yMax - YY3 * sc
  x4 = xmin + XX4 * sc: y4 = yMax - YY4 * sc
  x5 = xmin + XX5 * sc: y5 = yMax - YY5 * sc
  x6 = xmin + XX6 * sc: y6 = yMax - YY6 * sc
  x7 = xmin + XX7 * sc: y7 = yMax - YY7 * sc
  x8 = xmin + XX8 * sc: y8 = yMax - YY8 * sc
  
  With ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, x1, y1)
       .AddNodes msoSegmentLine, msoEditingAuto, x1, y1
       .AddNodes msoSegmentLine, msoEditingAuto, x2, y2
       .AddNodes msoSegmentLine, msoEditingAuto, x3, y3
       .AddNodes msoSegmentLine, msoEditingAuto, x4, y4
       .AddNodes msoSegmentLine, msoEditingAuto, x5, y5
       .AddNodes msoSegmentLine, msoEditingAuto, x6, y6
       .AddNodes msoSegmentLine, msoEditingAuto, x7, y7
       .AddNodes msoSegmentLine, msoEditingAuto, x8, y8
       .AddNodes msoSegmentLine, msoEditingAuto, x1, y1
       .ConvertToShape.Select
  End With
  selection.name = "E_" & Eno '図形に名をつける
  
  With selection.ShapeRange
       .fill.Visible = msoFalse
       .line.ForeColor.SchemeColor = 0
  End With

End Sub
Private Function ReadNumericCell(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal colNo As Long, _
                                 ByVal fieldName As String, ByVal allowBlank As Boolean) As Double
  Dim cellValue As Variant
  cellValue = ws.Cells(rowNo, colNo).value2
  If IsError(cellValue) Then
    Err.Raise vbObjectError + 3001, "FEM.ReadNumericCell", fieldName & " にExcelエラー値があります。行=" & CStr(rowNo)
  End If
  If IsEmpty(cellValue) Then
    If allowBlank Then
      ReadNumericCell = 0#
      Exit Function
    End If
    Err.Raise vbObjectError + 3002, "FEM.ReadNumericCell", fieldName & " が空白です。行=" & CStr(rowNo)
  End If
  If VarType(cellValue) = vbString Then
    If Len(Trim$(CStr(cellValue))) = 0 Then
      If allowBlank Then
        ReadNumericCell = 0#
        Exit Function
      End If
      Err.Raise vbObjectError + 3002, "FEM.ReadNumericCell", fieldName & " が空白です。行=" & CStr(rowNo)
    End If
  End If
  If Not IsNumeric(cellValue) Then
    Err.Raise vbObjectError + 3003, "FEM.ReadNumericCell", fieldName & " は数値で指定してください。行=" & CStr(rowNo)
  End If
  ReadNumericCell = CDbl(cellValue)
End Function

Private Function ReadIntegerCell(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal colNo As Long, _
                                 ByVal fieldName As String, ByVal allowBlank As Boolean) As Long
  Dim numericValue As Double
  numericValue = ReadNumericCell(ws, rowNo, colNo, fieldName, allowBlank)
  If numericValue <> Fix(numericValue) Then
    Err.Raise vbObjectError + 3004, "FEM.ReadIntegerCell", fieldName & " は整数で指定してください。行=" & CStr(rowNo)
  End If
  ReadIntegerCell = CLng(numericValue)
End Function

Private Function P6BulkValueIsBlank(ByVal value As Variant) As Boolean
  If IsError(value) Or IsEmpty(value) Then
    P6BulkValueIsBlank = True
  ElseIf VarType(value) = vbString Then
    P6BulkValueIsBlank = (Len(Trim$(CStr(value))) = 0)
  Else
    P6BulkValueIsBlank = False
  End If
End Function

Private Function P6BulkReadNumeric(ByRef values As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long, _
                                   ByVal fieldName As String, ByVal allowBlank As Boolean) As Double
  Dim value As Variant
  value = values(rowIndex, columnIndex)
  If IsError(value) Then
    Err.Raise vbObjectError + 3001, "FEM.P6BulkReadNumeric", fieldName & " にエラー値があります。行=" & CStr(rowIndex + 1)
  End If
  If P6BulkValueIsBlank(value) Then
    If allowBlank Then
      P6BulkReadNumeric = 0#
      Exit Function
    End If
    Err.Raise vbObjectError + 3002, "FEM.P6BulkReadNumeric", fieldName & " が空白です。行=" & CStr(rowIndex + 1)
  End If
  If Not IsNumeric(value) Then
    Err.Raise vbObjectError + 3003, "FEM.P6BulkReadNumeric", fieldName & " は数値で指定してください。行=" & CStr(rowIndex + 1)
  End If
  P6BulkReadNumeric = CDbl(value)
End Function

Private Function P6BulkReadInteger(ByRef values As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long, _
                                   ByVal fieldName As String, ByVal allowBlank As Boolean) As Long
  Dim numericValue As Double
  numericValue = P6BulkReadNumeric(values, rowIndex, columnIndex, fieldName, allowBlank)
  If numericValue <> Fix(numericValue) Then
    Err.Raise vbObjectError + 3004, "FEM.P6BulkReadInteger", fieldName & " は整数で指定してください。行=" & CStr(rowIndex + 1)
  End If
  P6BulkReadInteger = CLng(numericValue)
End Function

Private Function P6ParseOnOff(ByVal rawValue As Variant, ByVal defaultOn As Boolean) As Boolean
  Dim token As String
  P6ParseOnOff = defaultOn
  If IsError(rawValue) Then Exit Function
  If P6BulkValueIsBlank(rawValue) Then Exit Function
  If IsNumeric(rawValue) Then
    P6ParseOnOff = (Abs(CDbl(rawValue)) >= 0.5)
    Exit Function
  End If
  token = UCase$(Trim$(CStr(rawValue)))
  If token = "0" Or token = "OFF" Or token = "NO" Or token = "N" Or token = "しない" Or token = "無効" Or token = "FALSE" Then
    P6ParseOnOff = False
  ElseIf token = "1" Or token = "ON" Or token = "YES" Or token = "Y" Or token = "する" Or token = "有効" Or token = "TRUE" Then
    P6ParseOnOff = True
  End If
End Function

Private Function P6ParseMaterialKind(ByVal rawValue As Variant) As String
  Dim token As String
  P6ParseMaterialKind = "SOIL"
  If IsError(rawValue) Then Exit Function
  If P6BulkValueIsBlank(rawValue) Then Exit Function
  token = UCase$(Trim$(CStr(rawValue)))
  If token = "JOINT" Or token = "接合" Or token = "IF" Or token = "INTERFACE" Then
    P6ParseMaterialKind = "JOINT"
  ElseIf token = "STRUCT" Or token = "構造" Or token = "CONCRETE" Or token = "RC" Then
    P6ParseMaterialKind = "STRUCT"
  End If
End Function

Public Function P6MaterialKindOf(ByVal materialIndex As Long) As String
  If materialIndex < 0 Or materialIndex > NumberOfMaterial - 1 Then
    P6MaterialKindOf = "SOIL"
    Exit Function
  End If
  If Len(Material(materialIndex).kind) = 0 Then
    P6MaterialKindOf = "SOIL"
  Else
    P6MaterialKindOf = Material(materialIndex).kind
  End If
End Function

Private Sub P6LoadInputCache()
  Dim wsNode As Worksheet, wsMaterial As Worksheet, wsElement As Worksheet
  Dim nodeValues As Variant, materialValues As Variant, elementValues As Variant
  Dim lastRow As Long, rowIndex As Long, dataId As Long, j As Long, k As Long
  Dim conditionValue As Double
  Dim topologyFingerprint As Double, geometryFingerprint As Double, materialFingerprint As Double
  Dim seen() As Boolean

  P6InputCacheReady = False
  Set wsNode = ThisWorkbook.Worksheets("節点データ")
  Set wsMaterial = ThisWorkbook.Worksheets("材料データ")
  Set wsElement = ThisWorkbook.Worksheets("要素データ")

  lastRow = wsNode.Cells(wsNode.rows.count, 1).End(xlUp).row
  If lastRow < 2 Then Err.Raise vbObjectError + 3005, "FEM.P6LoadInputCache", "節点データ にデータがありません。"
  P6InputNodeCount = lastRow - 1
  If P6InputNodeCount > FEM_FIXED_DATA_CAPACITY Then Err.Raise vbObjectError + 3011, "FEM.P6LoadInputCache", "節点数が固定配列容量を超えています。"
  geometryFingerprint = P6FingerprintStep(CDbl(P6InputNodeCount), 17#)
  nodeValues = wsNode.range(wsNode.Cells(2, 1), wsNode.Cells(lastRow, 9)).value2
  ReDim P6InputNodeId(1 To P6InputNodeCount)
  ReDim P6InputNodeX(1 To P6InputNodeCount)
  ReDim P6InputNodeY(1 To P6InputNodeCount)
  ReDim P6InputNodeCondX(1 To P6InputNodeCount)
  ReDim P6InputNodeCondY(1 To P6InputNodeCount)
  ReDim P6InputNodeDispX(1 To P6InputNodeCount)
  ReDim P6InputNodeDispY(1 To P6InputNodeCount)
  ReDim P6InputNodeForceX(1 To P6InputNodeCount)
  ReDim P6InputNodeForceY(1 To P6InputNodeCount)
  ReDim seen(1 To P6InputNodeCount)
  For rowIndex = 1 To P6InputNodeCount
    dataId = P6BulkReadInteger(nodeValues, rowIndex, 1, "節点番号", False)
    If dataId < 1 Or dataId > P6InputNodeCount Then Err.Raise vbObjectError + 3006, "FEM.P6LoadInputCache", "節点番号は1からの連番で指定してください。行=" & CStr(rowIndex + 1)
    If seen(dataId) Then Err.Raise vbObjectError + 3007, "FEM.P6LoadInputCache", "節点番号が重複しています。ID=" & CStr(dataId)
    seen(dataId) = True
    P6InputNodeId(rowIndex) = dataId
    P6InputNodeX(rowIndex) = P6BulkReadNumeric(nodeValues, rowIndex, 2, "X座標", False)
    P6InputNodeY(rowIndex) = P6BulkReadNumeric(nodeValues, rowIndex, 3, "Y座標", False)
    geometryFingerprint = P6FingerprintStep(geometryFingerprint, CDbl(dataId))
    geometryFingerprint = P6FingerprintStep(geometryFingerprint, P6QuantizeScaled(P6InputNodeX(rowIndex), 1000000#))
    geometryFingerprint = P6FingerprintStep(geometryFingerprint, P6QuantizeScaled(P6InputNodeY(rowIndex), 1000000#))
    conditionValue = P6BulkReadNumeric(nodeValues, rowIndex, 4, "X拘束条件", True)
    If conditionValue <> Fix(conditionValue) Then Err.Raise vbObjectError + 3014, "FEM.P6LoadInputCache", "X拘束条件は整数で指定してください。節点=" & CStr(dataId)
    P6InputNodeCondX(rowIndex) = CLng(conditionValue)
    conditionValue = P6BulkReadNumeric(nodeValues, rowIndex, 5, "Y拘束条件", True)
    If conditionValue <> Fix(conditionValue) Then Err.Raise vbObjectError + 3015, "FEM.P6LoadInputCache", "Y拘束条件は整数で指定してください。節点=" & CStr(dataId)
    P6InputNodeCondY(rowIndex) = CLng(conditionValue)
    P6InputNodeDispX(rowIndex) = P6BulkReadNumeric(nodeValues, rowIndex, 6, "X強制変位", True)
    P6InputNodeDispY(rowIndex) = P6BulkReadNumeric(nodeValues, rowIndex, 7, "Y強制変位", True)
    P6InputNodeForceX(rowIndex) = P6BulkReadNumeric(nodeValues, rowIndex, 8, "X節点力", True)
    P6InputNodeForceY(rowIndex) = P6BulkReadNumeric(nodeValues, rowIndex, 9, "Y節点力", True)
  Next rowIndex
  For dataId = 1 To P6InputNodeCount
    If Not seen(dataId) Then Err.Raise vbObjectError + 3008, "FEM.P6LoadInputCache", "節点番号に欠番があります。ID=" & CStr(dataId)
  Next dataId

  lastRow = wsMaterial.Cells(wsMaterial.rows.count, 1).End(xlUp).row
  If lastRow < 2 Then Err.Raise vbObjectError + 3005, "FEM.P6LoadInputCache", "材料データ にデータがありません。"
  P6InputMaterialCount = lastRow - 1
  If P6InputMaterialCount > 32767 Then Err.Raise vbObjectError + 3013, "FEM.P6LoadInputCache", "材料数が許容範囲を超えています。"
  materialFingerprint = P6FingerprintStep(CDbl(P6InputMaterialCount), 17#)
  materialValues = wsMaterial.range(wsMaterial.Cells(2, 1), wsMaterial.Cells(lastRow, 13)).value2
  ReDim P6InputMaterialId(1 To P6InputMaterialCount)
  ReDim P6InputMaterialYoung(1 To P6InputMaterialCount)
  ReDim P6InputMaterialPoisson(1 To P6InputMaterialCount)
  ReDim P6InputMaterialThickness(1 To P6InputMaterialCount)
  ReDim P6InputMaterialWeight(1 To P6InputMaterialCount)
  ReDim P6InputMaterialFriction(1 To P6InputMaterialCount)
  ReDim P6InputMaterialCohesion(1 To P6InputMaterialCount)
  ReDim P6InputMaterialDilation(1 To P6InputMaterialCount)
  ReDim P6InputMaterialReduce(1 To P6InputMaterialCount)
  ReDim P6InputMaterialInitialActive(1 To P6InputMaterialCount)
  ReDim P6InputMaterialKind(1 To P6InputMaterialCount)
  ReDim P6InputMaterialKs(1 To P6InputMaterialCount)
  ReDim P6InputMaterialAllowTension(1 To P6InputMaterialCount)
  Erase seen
  ReDim seen(1 To P6InputMaterialCount)
  For rowIndex = 1 To P6InputMaterialCount
    dataId = P6BulkReadInteger(materialValues, rowIndex, 1, "材料番号", False)
    If dataId < 1 Or dataId > P6InputMaterialCount Then Err.Raise vbObjectError + 3006, "FEM.P6LoadInputCache", "材料番号は1からの連番で指定してください。行=" & CStr(rowIndex + 1)
    If seen(dataId) Then Err.Raise vbObjectError + 3007, "FEM.P6LoadInputCache", "材料番号が重複しています。ID=" & CStr(dataId)
    seen(dataId) = True
    P6InputMaterialId(rowIndex) = dataId
    P6InputMaterialYoung(rowIndex) = P6BulkReadNumeric(materialValues, rowIndex, 2, "ヤング率", False)
    P6InputMaterialPoisson(rowIndex) = P6BulkReadNumeric(materialValues, rowIndex, 3, "ポアソン比", False)
    P6InputMaterialThickness(rowIndex) = P6BulkReadNumeric(materialValues, rowIndex, 4, "板厚", False)
    P6InputMaterialWeight(rowIndex) = P6BulkReadNumeric(materialValues, rowIndex, 5, "単位体積重量", False)
    P6InputMaterialFriction(rowIndex) = P6BulkReadNumeric(materialValues, rowIndex, 6, "内部摩擦角", False)
    P6InputMaterialCohesion(rowIndex) = P6BulkReadNumeric(materialValues, rowIndex, 7, "粘着力", False)
    P6InputMaterialDilation(rowIndex) = P6BulkReadNumeric(materialValues, rowIndex, 8, "ダイレタンシー角", False)
    P6InputMaterialReduce(rowIndex) = P6ParseOnOff(materialValues(rowIndex, 9), True)
    P6InputMaterialInitialActive(rowIndex) = P6ParseOnOff(materialValues(rowIndex, 10), True)
    P6InputMaterialKind(rowIndex) = P6ParseMaterialKind(materialValues(rowIndex, 11))
    If P6InputMaterialKind(rowIndex) = "JOINT" Then
      P6InputMaterialKs(rowIndex) = P6BulkReadNumeric(materialValues, rowIndex, 12, "せん断剛性", True)
      If P6InputMaterialThickness(rowIndex) <= 0# Then P6InputMaterialThickness(rowIndex) = 1#
      If P6InputMaterialKs(rowIndex) <= 0# Then P6InputMaterialKs(rowIndex) = 0.1 * P6InputMaterialYoung(rowIndex)
    Else
      ' SOIL/STRUCT: ks is unused. Ignore leftover legend text in column L.
      If P6BulkValueIsBlank(materialValues(rowIndex, 12)) Then
        P6InputMaterialKs(rowIndex) = 0#
      ElseIf IsNumeric(materialValues(rowIndex, 12)) Then
        P6InputMaterialKs(rowIndex) = CDbl(materialValues(rowIndex, 12))
      Else
        P6InputMaterialKs(rowIndex) = 0#
      End If
    End If
    P6InputMaterialAllowTension(rowIndex) = P6ParseOnOff(materialValues(rowIndex, 13), False)
    materialFingerprint = P6FingerprintStep(materialFingerprint, CDbl(dataId))
    materialFingerprint = P6FingerprintStep(materialFingerprint, P6QuantizeScaled(P6InputMaterialYoung(rowIndex), 1000000#))
    materialFingerprint = P6FingerprintStep(materialFingerprint, P6QuantizeScaled(P6InputMaterialPoisson(rowIndex), 1000000000#))
    materialFingerprint = P6FingerprintStep(materialFingerprint, P6QuantizeScaled(P6InputMaterialThickness(rowIndex), 1000000#))
    materialFingerprint = P6FingerprintStep(materialFingerprint, P6QuantizeScaled(P6InputMaterialWeight(rowIndex), 1000000#))
    materialFingerprint = P6FingerprintStep(materialFingerprint, P6QuantizeScaled(P6InputMaterialFriction(rowIndex), 1000000#))
    materialFingerprint = P6FingerprintStep(materialFingerprint, P6QuantizeScaled(P6InputMaterialCohesion(rowIndex), 1000000#))
    materialFingerprint = P6FingerprintStep(materialFingerprint, P6QuantizeScaled(P6InputMaterialDilation(rowIndex), 1000000#))
    materialFingerprint = P6FingerprintStep(materialFingerprint, P6QuantizeScaled(P6InputMaterialKs(rowIndex), 1000000#))
    If P6InputMaterialKind(rowIndex) = "JOINT" Then
      materialFingerprint = P6FingerprintStep(materialFingerprint, 2#)
    ElseIf P6InputMaterialKind(rowIndex) = "STRUCT" Then
      materialFingerprint = P6FingerprintStep(materialFingerprint, 1#)
    End If
    If P6InputMaterialYoung(rowIndex) <= 0# Then Err.Raise vbObjectError + 3016, "FEM.P6LoadInputCache", "ヤング率は正で指定してください。材料=" & CStr(dataId)
    If P6InputMaterialKind(rowIndex) <> "JOINT" Then
      If P6InputMaterialPoisson(rowIndex) <= -1# Or P6InputMaterialPoisson(rowIndex) >= 0.5 Then Err.Raise vbObjectError + 3017, "FEM.P6LoadInputCache", "平面ひずみのポアソン比は-1より大きく0.5未満で指定してください。材料=" & CStr(dataId) & " 値=" & Format$(P6InputMaterialPoisson(rowIndex), "0.##########")
    End If
    If P6InputMaterialThickness(rowIndex) <= 0# Then Err.Raise vbObjectError + 3018, "FEM.P6LoadInputCache", "板厚は正で指定してください。材料=" & CStr(dataId)
  Next rowIndex
  For dataId = 1 To P6InputMaterialCount
    If Not seen(dataId) Then Err.Raise vbObjectError + 3008, "FEM.P6LoadInputCache", "材料番号に欠番があります。ID=" & CStr(dataId)
  Next dataId

  lastRow = wsElement.Cells(wsElement.rows.count, 1).End(xlUp).row
  If lastRow < 2 Then Err.Raise vbObjectError + 3005, "FEM.P6LoadInputCache", "要素データ にデータがありません。"
  P6InputElementCount = lastRow - 1
  If P6InputElementCount > FEM_FIXED_DATA_CAPACITY Then Err.Raise vbObjectError + 3012, "FEM.P6LoadInputCache", "要素数が固定配列容量を超えています。"
  topologyFingerprint = P6FingerprintStep(CDbl(P6InputNodeCount), CDbl(P6InputElementCount))
  elementValues = wsElement.range(wsElement.Cells(2, 1), wsElement.Cells(lastRow, 10)).value2
  ReDim P6InputElementId(1 To P6InputElementCount)
  ReDim P6InputElementNode(1 To P6InputElementCount, 1 To 8)
  ReDim P6InputElementMaterial(1 To P6InputElementCount)
  Erase seen
  ReDim seen(1 To P6InputElementCount)
  Dim kindById() As String, isJointElem As Boolean, nodeUse As Long
  ReDim kindById(1 To P6InputMaterialCount)
  For rowIndex = 1 To P6InputMaterialCount
    kindById(P6InputMaterialId(rowIndex)) = P6InputMaterialKind(rowIndex)
  Next rowIndex
  For rowIndex = 1 To P6InputElementCount
    dataId = P6BulkReadInteger(elementValues, rowIndex, 1, "要素番号", False)
    If dataId < 1 Or dataId > P6InputElementCount Then Err.Raise vbObjectError + 3006, "FEM.P6LoadInputCache", "要素番号は1からの連番で指定してください。行=" & CStr(rowIndex + 1)
    If seen(dataId) Then Err.Raise vbObjectError + 3007, "FEM.P6LoadInputCache", "要素番号が重複しています。ID=" & CStr(dataId)
    seen(dataId) = True
    P6InputElementId(rowIndex) = dataId
    topologyFingerprint = P6FingerprintStep(topologyFingerprint, CDbl(dataId))
    P6InputElementMaterial(rowIndex) = P6BulkReadInteger(elementValues, rowIndex, 10, "要素材料番号", False)
    If P6InputElementMaterial(rowIndex) < 1 Or P6InputElementMaterial(rowIndex) > P6InputMaterialCount Then Err.Raise vbObjectError + 3021, "FEM.P6LoadInputCache", "未定義の材料番号です。要素=" & CStr(dataId)
    isJointElem = (kindById(P6InputElementMaterial(rowIndex)) = "JOINT")
    If isJointElem Then nodeUse = 6 Else nodeUse = 8
    For j = 1 To 8
      If isJointElem And j > 6 Then
        P6InputElementNode(rowIndex, j) = P6BulkReadInteger(elementValues, rowIndex, j + 1, "接合要素節点番号", True)
      Else
        P6InputElementNode(rowIndex, j) = P6BulkReadInteger(elementValues, rowIndex, j + 1, "Q8要素節点番号", False)
      End If
      topologyFingerprint = P6FingerprintStep(topologyFingerprint, CDbl(P6InputElementNode(rowIndex, j)))
      If j <= nodeUse Then
        If P6InputElementNode(rowIndex, j) < 1 Or P6InputElementNode(rowIndex, j) > P6InputNodeCount Then Err.Raise vbObjectError + 3019, "FEM.P6LoadInputCache", "未定義の節点番号です。要素=" & CStr(dataId)
        For k = 1 To j - 1
          If P6InputElementNode(rowIndex, j) = P6InputElementNode(rowIndex, k) Then Err.Raise vbObjectError + 3020, "FEM.P6LoadInputCache", "要素内で節点番号が重複しています。要素=" & CStr(dataId)
        Next k
      End If
    Next j
    topologyFingerprint = P6FingerprintStep(topologyFingerprint, CDbl(P6InputElementMaterial(rowIndex)))
  Next rowIndex
  For dataId = 1 To P6InputElementCount
    If Not seen(dataId) Then Err.Raise vbObjectError + 3008, "FEM.P6LoadInputCache", "要素番号に欠番があります。ID=" & CStr(dataId)
  Next dataId

  P6InputHorizontalLoad = 0#
  P6InputVerticalLoad = -1#
  P6InputMode = 1
  P6InputFSS = 1#
  If Not P6InputFingerprintReady Or topologyFingerprint <> P6InputTopologyFingerprint Then
    P6MeshRevision = P6NextRevision(P6MeshRevision)
    P6InputTopologyFingerprint = topologyFingerprint
  End If
  If Not P6InputFingerprintReady Or geometryFingerprint <> P6InputGeometryFingerprint Then
    P6GeometryRevision = P6NextRevision(P6GeometryRevision)
    P6InputGeometryFingerprint = geometryFingerprint
  End If
  If Not P6InputFingerprintReady Or materialFingerprint <> P6InputMaterialFingerprint Then
    P6MaterialInputRevision = P6NextRevision(P6MaterialInputRevision)
    P6InputMaterialFingerprint = materialFingerprint
  End If
  P6InputFingerprintReady = True
  NumberOfNode = P6InputNodeCount
  NumberOfMaterial = P6InputMaterialCount
  NumberOfElement = P6InputElementCount
  P6InputCacheReady = True
End Sub

Public Sub FEMInvalidateInputCache()
  P6InputCacheReady = False
  'メッシュ／入力シートの変更後に、図形表示が旧Elem配列を再利用しないよう件数を無効化する。
  '解析開始時は直後のP6LoadInputCacheで正しい件数へ再設定される。
  NumberOfNode = 0
  NumberOfMaterial = 0
  NumberOfElement = 0
  NumberOfFreeNode = 0
End Sub

Public Sub FEMInvalidateRemeshedModel()
  ' 再メッシュ採用後。入力キャッシュだけでなく RCM・帯域因子・弾性K・SRM スナップも捨てる。
  FEMInvalidateInputCache
  P6InputFingerprintReady = False
  P6InputTopologyFingerprint = 0#
  P6InputGeometryFingerprint = 0#
  P6InputMaterialFingerprint = 0#
  P6RCMCacheSignature = vbNullString
  P6RCMCacheReady = False
  P6MeshInvariantReady = False
  P6MeshInvariantSignature = vbNullString
  P6CSRPatternReady = False
  P6CSRReady = False
  P3ElementActiveReady = False
  P3SupportCondReady = False
  P3SrmSnapReady = False
  P6InvalidateTangentGeneration
  P6InvalidateActiveDependentCaches
End Sub

Function SetDataCount(ByVal sheetName As String) As Long
  If Not P6InputCacheReady Then P6LoadInputCache
  Select Case sheetName
    Case "節点データ": SetDataCount = P6InputNodeCount
    Case "材料データ": SetDataCount = P6InputMaterialCount
    Case "要素データ": SetDataCount = P6InputElementCount
    Case Else: Err.Raise vbObjectError + 3009, "FEM.SetDataCount", "未対応のデータシートです。シート=" & sheetName
  End Select
End Function

Function SetNumberOfNode() As Long
  SetNumberOfNode = SetDataCount("節点データ")
End Function
Function SetNumberOfMaterial() As Long
  SetNumberOfMaterial = SetDataCount("材料データ")
End Function
Function SetNumberOfElement() As Long
  SetNumberOfElement = SetDataCount("要素データ")
End Function

Public Sub ValidateInputData()
  P6LoadInputCache
  If NumberOfNode < 1 Then Err.Raise vbObjectError + 3457, "FEM.ValidateInputData", "節点数は1以上で指定してください。"
  If NumberOfElement < 1 Then Err.Raise vbObjectError + 3458, "FEM.ValidateInputData", "要素数は1以上で指定してください。"
  If NumberOfMaterial < 1 Then Err.Raise vbObjectError + 3459, "FEM.ValidateInputData", "材料数は1以上で指定してください。"
  P5ValidateAnalysisMesh
End Sub

' P5: 解析入力のメッシュを、出力を書き換える前に検証する。
Public Sub P5ValidateAnalysisMesh()
  Dim connectivity() As Long, materialNo() As Long
  Dim usedNode() As Boolean, materialSeen() As Boolean
  Dim rowNo As Long, i As Long, j As Long, nodeId As Long, elementId As Long
  Dim materialId As Long, minX As Double, maxX As Double, minY As Double, maxY As Double
  Dim firstPoint As Boolean

  If Not P6InputCacheReady Then P6LoadInputCache
  If NumberOfNode < 1 Or NumberOfElement < 1 Then Exit Sub
  ReDim usedNode(1 To NumberOfNode)
  ReDim connectivity(1 To NumberOfElement, 1 To 8)
  ReDim materialNo(1 To NumberOfElement)
  ReDim materialSeen(1 To NumberOfMaterial)
  firstPoint = True
  For rowNo = 1 To NumberOfNode
    nodeId = P6InputNodeId(rowNo)
    If nodeId <> rowNo Then Err.Raise vbObjectError + 3460, "FEM.P5ValidateAnalysisMesh", "節点番号は1からの連番で指定してください。行=" & CStr(rowNo + 1)
    If firstPoint Then
      minX = P6InputNodeX(rowNo): maxX = minX
      minY = P6InputNodeY(rowNo): maxY = minY
      firstPoint = False
    Else
      If P6InputNodeX(rowNo) < minX Then minX = P6InputNodeX(rowNo)
      If P6InputNodeX(rowNo) > maxX Then maxX = P6InputNodeX(rowNo)
      If P6InputNodeY(rowNo) < minY Then minY = P6InputNodeY(rowNo)
      If P6InputNodeY(rowNo) > maxY Then maxY = P6InputNodeY(rowNo)
    End If
  Next rowNo

  P5ModelScale = maxX - minX
  If maxY - minY > P5ModelScale Then P5ModelScale = maxY - minY
  If P5ModelScale <= 0# Then Err.Raise vbObjectError + 3461, "FEM.P5ValidateAnalysisMesh", "全節点が同一点でモデル寸法を決定できません。"
  P5GeometryTolerance = P5ModelScale * 0.0000000001
  If P5GeometryTolerance < 0.000000000001 Then P5GeometryTolerance = 0.000000000001
  P5AreaTolerance = P5ModelScale * P5ModelScale * 0.000000000001
  If P5AreaTolerance < 1E-24 Then P5AreaTolerance = 1E-24
  P5MinElementArea = 1E+308
  P5MinJacobian = 1E+308
  P5MaxAspectRatio = 0#
  P5MinAngleDeg = 180#
  P5MaxAngleDeg = 0#
  P5BoundaryEdgeCount = 0
  P5SharedEdgeCount = 0
  P5MaterialInterfaceCount = 0
  P5IsolatedNodeCount = 0
  P5ConnectedComponentCount = 0

  For rowNo = 1 To NumberOfMaterial
    materialId = P6InputMaterialId(rowNo)
    If materialId < 1 Or materialId > NumberOfMaterial Then Err.Raise vbObjectError + 3462, "FEM.P5ValidateAnalysisMesh", "材料番号が連番範囲外です。材料=" & CStr(materialId)
    If materialSeen(materialId) Then Err.Raise vbObjectError + 3463, "FEM.P5ValidateAnalysisMesh", "材料番号が重複しています。材料=" & CStr(materialId)
    materialSeen(materialId) = True
  Next rowNo

  Dim kindByIdV() As String, isJointElem As Boolean
  ReDim kindByIdV(1 To NumberOfMaterial)
  For rowNo = 1 To NumberOfMaterial
    kindByIdV(P6InputMaterialId(rowNo)) = P6InputMaterialKind(rowNo)
  Next rowNo
  For rowNo = 1 To NumberOfElement
    elementId = P6InputElementId(rowNo)
    If elementId <> rowNo Then Err.Raise vbObjectError + 3464, "FEM.P5ValidateAnalysisMesh", "要素番号は1からの連番で指定してください。行=" & CStr(rowNo + 1)
    materialId = P6InputElementMaterial(rowNo)
    If materialId < 1 Or materialId > NumberOfMaterial Then Err.Raise vbObjectError + 3466, "FEM.P5ValidateAnalysisMesh", "要素材料番号が範囲外です。要素=" & CStr(elementId)
    materialNo(elementId) = materialId
    isJointElem = (kindByIdV(materialId) = "JOINT")
    For j = 1 To 8
      nodeId = P6InputElementNode(rowNo, j)
      If nodeId > 0 Then
        If nodeId > NumberOfNode Then Err.Raise vbObjectError + 3465, "FEM.P5ValidateAnalysisMesh", "要素接続番号が範囲外です。要素=" & CStr(elementId)
        connectivity(elementId, j) = nodeId
        usedNode(nodeId) = True
      Else
        connectivity(elementId, j) = 0
        If (Not isJointElem) Or j <= 6 Then Err.Raise vbObjectError + 3465, "FEM.P5ValidateAnalysisMesh", "要素接続番号が範囲外です。要素=" & CStr(elementId)
      End If
    Next j
    If Not isJointElem Then P5EvaluateQuadQuality elementId, connectivity, P6InputNodeX, P6InputNodeY
  Next rowNo
  For i = 1 To NumberOfNode
    If Not usedNode(i) Then
      P5IsolatedNodeCount = P5IsolatedNodeCount + 1
    End If
  Next i
  If P5IsolatedNodeCount > 0 Then Err.Raise vbObjectError + 3467, "FEM.P5ValidateAnalysisMesh", "要素に接続していない節点があります。件数=" & CStr(P5IsolatedNodeCount)
  P5ValidateElementTopology connectivity, materialNo, NumberOfNode, NumberOfElement
End Sub

Private Sub P5EvaluateQuadQuality(ByVal elementId As Long, ByRef connectivity() As Long, ByRef x() As Double, ByRef y() As Double)
  Dim nodeNo(1 To 8) As Long, edgeLength(1 To 4) As Double
  Dim i As Long, previousIndex As Long, nextIndex As Long
  Dim area As Double, minJacobian As Double, jacobianValue As Double
  Dim minEdge As Double, maxEdge As Double, angleDeg As Double
  Dim dx1 As Double, dy1 As Double, dx2 As Double, dy2 As Double, dotValue As Double, crossValue As Double
  Dim xiValues(1 To 4) As Double, etaValues(1 To 4) As Double

  For i = 1 To 8: nodeNo(i) = connectivity(elementId, i): Next i
  For i = 1 To 7
    For nextIndex = i + 1 To 8
      If nodeNo(i) = nodeNo(nextIndex) Then Err.Raise vbObjectError + 3468, "FEM.P5ValidateAnalysisMesh", "Q8要素内で節点番号が重複しています。要素=" & CStr(elementId)
    Next nextIndex
  Next i
  area = 0.5 * ((x(nodeNo(1)) * y(nodeNo(2)) + x(nodeNo(2)) * y(nodeNo(3)) + x(nodeNo(3)) * y(nodeNo(4)) + x(nodeNo(4)) * y(nodeNo(1))) - _
                   (y(nodeNo(1)) * x(nodeNo(2)) + y(nodeNo(2)) * x(nodeNo(3)) + y(nodeNo(3)) * x(nodeNo(4)) + y(nodeNo(4)) * x(nodeNo(1))))
  If area <= P5AreaTolerance Then Err.Raise vbObjectError + 3469, "FEM.P5ValidateAnalysisMesh", "要素の向きが反時計回りでないか、面積が退化しています。要素=" & CStr(elementId)
  If area < P5MinElementArea Then P5MinElementArea = area

  minEdge = 1E+308: maxEdge = 0#
  For i = 1 To 4
    nextIndex = (i Mod 4) + 1
    edgeLength(i) = Sqr((x(nodeNo(nextIndex)) - x(nodeNo(i))) ^ 2 + (y(nodeNo(nextIndex)) - y(nodeNo(i))) ^ 2)
    If edgeLength(i) <= P5GeometryTolerance Then Err.Raise vbObjectError + 3470, "FEM.P5ValidateAnalysisMesh", "極短辺を持つ要素です。要素=" & CStr(elementId) & "、辺=" & CStr(i)
    If edgeLength(i) < minEdge Then minEdge = edgeLength(i)
    If edgeLength(i) > maxEdge Then maxEdge = edgeLength(i)
    previousIndex = ((i + 2) Mod 4) + 1
    dx1 = x(nodeNo(previousIndex)) - x(nodeNo(i)): dy1 = y(nodeNo(previousIndex)) - y(nodeNo(i))
    dx2 = x(nodeNo(nextIndex)) - x(nodeNo(i)): dy2 = y(nodeNo(nextIndex)) - y(nodeNo(i))
    dotValue = dx1 * dx2 + dy1 * dy2
    crossValue = dx1 * dy2 - dy1 * dx2
    angleDeg = P1Atan2(Abs(crossValue), dotValue) * 180# / 3.14159265358979
    If angleDeg < P5MinAngleDeg Then P5MinAngleDeg = angleDeg
    If angleDeg > P5MaxAngleDeg Then P5MaxAngleDeg = angleDeg
    If angleDeg < 0.001 Or angleDeg > 179.999 Then Err.Raise vbObjectError + 3471, "FEM.P5ValidateAnalysisMesh", "要素角度が退化しています。要素=" & CStr(elementId) & "、頂点=" & CStr(i)
  Next i
  If minEdge > 0# Then
    If maxEdge / minEdge > P5MaxAspectRatio Then P5MaxAspectRatio = maxEdge / minEdge
  End If

  xiValues(1) = -1#: etaValues(1) = -1#: xiValues(2) = 1#: etaValues(2) = -1#
  xiValues(3) = 1#: etaValues(3) = 1#: xiValues(4) = -1#: etaValues(4) = 1#
  minJacobian = 1E+308
  For i = 1 To 4
    jacobianValue = P5Q8Jacobian(x, y, nodeNo, xiValues(i), etaValues(i))
    If jacobianValue < minJacobian Then minJacobian = jacobianValue
  Next i
  xiValues(1) = -0.577350269189626: etaValues(1) = -0.577350269189626
  xiValues(2) = 0.577350269189626: etaValues(2) = -0.577350269189626
  xiValues(3) = 0.577350269189626: etaValues(3) = 0.577350269189626
  xiValues(4) = -0.577350269189626: etaValues(4) = 0.577350269189626
  For i = 1 To 4
    jacobianValue = P5Q8Jacobian(x, y, nodeNo, xiValues(i), etaValues(i))
    If jacobianValue < minJacobian Then minJacobian = jacobianValue
  Next i
  If minJacobian < P5MinJacobian Then P5MinJacobian = minJacobian
  If minJacobian <= P5AreaTolerance Then Err.Raise vbObjectError + 3472, "FEM.P5ValidateAnalysisMesh", "要素Jacobianが正でないか小さすぎます。要素=" & CStr(elementId)
End Sub

Private Function P5Q8Jacobian(ByRef x() As Double, ByRef y() As Double, ByRef nodeNo() As Long, ByVal xi As Double, ByVal eta As Double) As Double
  Dim dNxi(1 To 8) As Double, dNeta(1 To 8) As Double
  Dim i As Long, dXdxi As Double, dYdxi As Double, dXdeta As Double, dYdeta As Double
  dNxi(1) = 0.25 * (1# - eta) * (2# * xi + eta): dNeta(1) = 0.25 * (1# - xi) * (xi + 2# * eta)
  dNxi(2) = 0.25 * (1# - eta) * (2# * xi - eta): dNeta(2) = 0.25 * (1# + xi) * (-xi + 2# * eta)
  dNxi(3) = 0.25 * (1# + eta) * (2# * xi + eta): dNeta(3) = 0.25 * (1# + xi) * (xi + 2# * eta)
  dNxi(4) = 0.25 * (1# + eta) * (2# * xi - eta): dNeta(4) = 0.25 * (1# - xi) * (-xi + 2# * eta)
  dNxi(5) = -xi * (1# - eta): dNeta(5) = -0.5 * (1# - xi * xi)
  dNxi(6) = 0.5 * (1# - eta * eta): dNeta(6) = -eta * (1# + xi)
  dNxi(7) = -xi * (1# + eta): dNeta(7) = 0.5 * (1# - xi * xi)
  dNxi(8) = -0.5 * (1# - eta * eta): dNeta(8) = -eta * (1# - xi)
  For i = 1 To 8
    dXdxi = dXdxi + dNxi(i) * x(nodeNo(i)): dYdxi = dYdxi + dNxi(i) * y(nodeNo(i))
    dXdeta = dXdeta + dNeta(i) * x(nodeNo(i)): dYdeta = dYdeta + dNeta(i) * y(nodeNo(i))
  Next i
  P5Q8Jacobian = dXdxi * dYdeta - dYdxi * dXdeta
End Function

Private Sub P5ValidateElementTopology(ByRef connectivity() As Long, ByRef materialNo() As Long, ByVal nodeCount As Long, ByVal elementCount As Long)
  Dim edgeMap As Object, parent() As Long
  Dim i As Long, j As Long, a As Long, b As Long, rootA As Long, rootB As Long
  Dim key As String, info As Variant, edgeCount As Long, firstMaterial As Long
  Set edgeMap = CreateObject("Scripting.Dictionary")
  edgeMap.CompareMode = 0
  ReDim parent(1 To nodeCount)
  For i = 1 To nodeCount: parent(i) = i: Next i
  For i = 1 To elementCount
    If connectivity(i, 7) = 0 Or connectivity(i, 8) = 0 Then
      a = 0
      For j = 1 To 6
        b = connectivity(i, j)
        If b > 0 Then
          If a > 0 Then P5Union parent, a, b Else a = b
        End If
      Next j
      GoTo NextTopoElement
    End If
    For j = 1 To 4
      a = connectivity(i, j): b = connectivity(i, (j Mod 4) + 1)
      If a < b Then key = CStr(a) & ":" & CStr(b) Else key = CStr(b) & ":" & CStr(a)
      P5Union parent, a, b
      If edgeMap.Exists(key) Then
        info = Split(CStr(edgeMap.item(key)), "|")
        edgeCount = CLng(info(0))
        firstMaterial = CLng(info(2))
        If edgeCount >= 2 Then Err.Raise vbObjectError + 3473, "FEM.P5ValidateAnalysisMesh", "辺を3要素以上で共有しています。辺=" & key
        P5SharedEdgeCount = P5SharedEdgeCount + 1
        If firstMaterial <> materialNo(i) Then P5MaterialInterfaceCount = P5MaterialInterfaceCount + 1
        edgeMap.item(key) = CStr(edgeCount + 1) & "|" & CStr(CLng(info(1))) & "|" & CStr(firstMaterial)
      Else
        edgeMap.Add key, "1|" & CStr(i) & "|" & CStr(materialNo(i))
        P5BoundaryEdgeCount = P5BoundaryEdgeCount + 1
      End If
    Next j
    For j = 5 To 8
      a = connectivity(i, j)
      If a <= 0 Then GoTo NextMidUnion
      If j = 5 Then b = connectivity(i, 1)
      If j = 6 Then b = connectivity(i, 2)
      If j = 7 Then b = connectivity(i, 3)
      If j = 8 Then b = connectivity(i, 4)
      If b > 0 Then P5Union parent, a, b
NextMidUnion:
    Next j
NextTopoElement:
  Next i
  For i = 1 To nodeCount
    If P5FindRoot(parent, i) = i Then P5ConnectedComponentCount = P5ConnectedComponentCount + 1
  Next i
  If P5ConnectedComponentCount > 1 Then Err.Raise vbObjectError + 3474, "FEM.P5ValidateAnalysisMesh", "メッシュが複数の非接続領域に分かれています。領域数=" & CStr(P5ConnectedComponentCount)
End Sub

Private Sub P5Union(ByRef parent() As Long, ByVal leftNode As Long, ByVal rightNode As Long)
  Dim leftRoot As Long, rightRoot As Long
  leftRoot = P5FindRoot(parent, leftNode): rightRoot = P5FindRoot(parent, rightNode)
  If leftRoot <> rightRoot Then parent(rightRoot) = leftRoot
End Sub

Private Function P5FindRoot(ByRef parent() As Long, ByVal nodeNo As Long) As Long
  Dim rootNo As Long, nextNo As Long
  rootNo = nodeNo
  Do While parent(rootNo) <> rootNo
    rootNo = parent(rootNo)
  Loop
  nextNo = nodeNo
  Do While parent(nextNo) <> nextNo
    nodeNo = parent(nextNo): parent(nextNo) = rootNo: nextNo = nodeNo
  Loop
  P5FindRoot = rootNo
End Function
Sub DrawNo(No, XX1, YY1)
Dim sc As Variant: Dim xm As Variant: Dim ym As Variant
Dim xmin As Variant: Dim yMax As Variant
  '節点番号描写
  xmin = P6ReadSetting("VIEW_MIN_X", 0#)
  yMax = P6ReadSetting("VIEW_MAX_Y", 0#)
  sc = P6ReadSetting("VIEW_SCALE", 1#)
  xm = xmin + XX1 * sc
  ym = yMax - YY1 * sc
  ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, xm, ym, 30, 20).Select
                            
   With selection
     .Characters.text = No
     .ShapeRange.fill.Visible = msoFalse
     .ShapeRange.line.Visible = msoFalse
     .Characters.Font.size = 9
     .name = "Node_" & No
  End With
End Sub
Sub SetNodeData()
  Dim rowNo As Long, nodeId As Long, nodeIndex As Long
  If Not P6InputCacheReady Then ValidateInputData
  ReDim x(NumberOfNode - 1), y(NumberOfNode - 1)
  For rowNo = 1 To NumberOfNode
    nodeId = P6InputNodeId(rowNo)
    nodeIndex = nodeId - 1
    x(nodeIndex) = P6InputNodeX(rowNo)
    y(nodeIndex) = P6InputNodeY(rowNo)
  Next rowNo
End Sub

Sub SetMaterialData()
  Dim rowNo As Long, materialId As Long, materialIndex As Long
  Dim piValue As Double, lameLambda As Double, shearModulus As Double
  Dim materialSignature As String, failMessage As String
  If Not P6InputCacheReady Then ValidateInputData
  piValue = 3.14159265358979
  nh = 0#
  nv = -1#
  mode = P6InputMode
  FSS = 1#
  ReDim Material(NumberOfMaterial - 1)
  For rowNo = 1 To NumberOfMaterial
    materialId = P6InputMaterialId(rowNo)
    materialIndex = materialId - 1
    Material(materialIndex).young = P6InputMaterialYoung(rowNo)
    Material(materialIndex).Poisson = P6InputMaterialPoisson(rowNo)
    Material(materialIndex).thickness = P6InputMaterialThickness(rowNo)
    Material(materialIndex).weight = P6InputMaterialWeight(rowNo)
    Material(materialIndex).fai = P6InputMaterialFriction(rowNo)
    Material(materialIndex).cohesion = P6InputMaterialCohesion(rowNo)
    Material(materialIndex).psai = P6InputMaterialDilation(rowNo)
    Material(materialIndex).ReduceStrength = P6InputMaterialReduce(rowNo)
    Material(materialIndex).InitialActive = P6InputMaterialInitialActive(rowNo)
    Material(materialIndex).kind = P6InputMaterialKind(rowNo)
    Material(materialIndex).kn = Material(materialIndex).young
    Material(materialIndex).ks = P6InputMaterialKs(rowNo)
    Material(materialIndex).AllowTension = P6InputMaterialAllowTension(rowNo)
    If Material(materialIndex).kind = "JOINT" Then
      If Material(materialIndex).ks <= 0# Then Material(materialIndex).ks = 0.1 * Material(materialIndex).kn
      If Not P3ValidateJointMaterial(Material(materialIndex), failMessage) Then Err.Raise vbObjectError + 3025, "FEM.SetMaterialData", failMessage & " 材料=" & CStr(materialId)
    Else
      If Not P3ValidateSolidMaterial(Material(materialIndex), failMessage) Then Err.Raise vbObjectError + 3025, "FEM.SetMaterialData", failMessage & " 材料=" & CStr(materialId)
    End If
    If Material(materialIndex).kind = "JOINT" Then
      Material(materialIndex).ElasticD00 = Material(materialIndex).kn
      Material(materialIndex).ElasticD01 = 0#
      Material(materialIndex).ElasticD22 = Material(materialIndex).ks
    Else
      shearModulus = Material(materialIndex).young / (2# * (1# + Material(materialIndex).Poisson))
      lameLambda = Material(materialIndex).young * Material(materialIndex).Poisson / ((1# + Material(materialIndex).Poisson) * (1# - 2# * Material(materialIndex).Poisson))
      Material(materialIndex).ElasticD00 = lameLambda + 2# * shearModulus
      Material(materialIndex).ElasticD01 = lameLambda
      Material(materialIndex).ElasticD22 = shearModulus
    End If
    Material(materialIndex).SinFriction = Sin(Material(materialIndex).fai * piValue / 180#)
    Material(materialIndex).CosFriction = Cos(Material(materialIndex).fai * piValue / 180#)
    Material(materialIndex).SinDilation = Sin(Material(materialIndex).psai * piValue / 180#)
     Material(materialIndex).CosDilation = Cos(Material(materialIndex).psai * piValue / 180#)
     Material(materialIndex).MaterialCacheReady = True
   Next rowNo
   materialSignature = P6ComputeMaterialSignature()
   If materialSignature <> P6MaterialSignature Then
     P6MaterialGeneration = P6MaterialGeneration + 1
     If P6MaterialGeneration <= 0 Then P6MaterialGeneration = 1
     P6MaterialSignature = materialSignature
   End If
End Sub

Private Function P6ComputeMaterialSignature() As String
  ' 材料値の検出は入力配列を読み込む時点で行い、ここでは整数世代だけを返す。
  P6ComputeMaterialSignature = CStr(P6MaterialInputRevision)
End Function
Private Function FEMGetFreeNode(ByVal originalNode As Long) As Long
    If originalNode < LBound(FEMFreeNodeMap) Or originalNode > UBound(FEMFreeNodeMap) Then
        Err.Raise vbObjectError + 3430, "FEM.FEMGetFreeNode", "Q8節点番号が範囲外です。"
    End If
    If FEMFreeNodeMap(originalNode) < 0 Then
        Err.Raise vbObjectError + 3431, "FEM.FEMGetFreeNode", "要素に接続していない節点を参照しました。"
    End If
    FEMGetFreeNode = FEMFreeNodeMap(originalNode)
End Function

Sub ReadFreeNode()
    Dim usedNode() As Boolean
    Dim i As Long, j As Long, rowNo As Long
    Dim nodeId As Long

    If NumberOfNode <= 0 Or NumberOfElement <= 0 Then
        Err.Raise vbObjectError + 3434, "FEM.ReadFreeNode", "節点数または要素数が0です。"
    End If
    nn = NumberOfElement - 1
    NAB = NumberOfNode - 1
    ReDim FEMFreeNodeMap(0 To NAB)
    ReDim usedNode(0 To NAB)
    For i = 0 To NAB
        FEMFreeNodeMap(i) = -1
    Next i

    For rowNo = 1 To NumberOfElement
        For j = 1 To 8
            If P6InputElementNode(rowNo, j) > 0 Then
                nodeId = P6InputElementNode(rowNo, j) - 1
                If nodeId < 0 Or nodeId > NAB Then Err.Raise vbObjectError + 3435, "FEM.ReadFreeNode", "要素接続節点番号が節点範囲外です。要素行=" & CStr(rowNo + 1)
                usedNode(nodeId) = True
            End If
        Next j
    Next rowNo

    NumberOfFreeNode = 0
    For i = 0 To NAB
        If usedNode(i) Then
            FEMFreeNodeMap(i) = NumberOfFreeNode
            NumberOfFreeNode = NumberOfFreeNode + 1
        End If
    Next i
End Sub

Sub SetElementData()
    Dim i As Long, j As Long, ii As Long, ax As Long, ix As Long
    Dim originalNode As Long
    Dim CoutN As Long

    ' Elem()を再生成するため、Bmat/Smat/DJとCSR散布表は再利用不可。
    ' 弾性Kキャッシュ自体は形状・材料署名が一致すれば再利用できる。
    P6ElasticOperatorCacheReady = False
    P6CSRScatterCount = 0
    P6CSRScatterElementCount = -1
    P6CSRScatterTangentGeneration = -1
    P6CSRScatterMaterialGeneration = -1
    P6CSRScatterPatternSignature = vbNullString
    Erase P6CSRScatterStart
    Erase P6CSRScatterLocalRow
    Erase P6CSRScatterLocalColumn
    Erase P6CSRScatterPosition

    If Not P6InputCacheReady Then ValidateInputData
    If SuppressUserMessages Then SaveP0Progress "element_validate_complete"
    ReadFreeNode
    If SuppressUserMessages Then SaveP0Progress "free_node_complete"

    nn = NumberOfElement - 1
    NAB = NumberOfNode - 1
    ReDim Elem(nn)
    nDof = NumberOfFreeNode * 2
    lastDof = nDof - 1
    P6KrylovLast = lastDof
    CoutN = lastDof
    ReDim Force(CoutN): ReDim Disp(CoutN): ReDim UDisp(CoutN)
    ReDim XXX(CoutN): ReDim NodeCond(CoutN): ReDim RForce(CoutN): ReDim eForce(CoutN)
    ReDim TDisp(CoutN): ReDim NForce(CoutN): ReDim mRisp(CoutN): ReDim iNForce(CoutN)
    ReDim OriginalForce(CoutN)
    P6EnsureReactionWorkspace
    If SuppressUserMessages Then SaveP0Progress "element_arrays_complete"
    P3HasJointElements = False

    For i = 1 To NumberOfElement
        ii = P6InputElementId(i) - 1
        If ii < 0 Or ii > nn Then Err.Raise vbObjectError + 3437, "FEM.SetElementData", "要素番号が連番範囲外です。"
        Elem(ii).ElNo = ii
        Elem(ii).MatNo = P6InputElementMaterial(i) - 1
        If Elem(ii).MatNo < 0 Or Elem(ii).MatNo > NumberOfMaterial - 1 Then
            Err.Raise vbObjectError + 3437, "FEM.SetElementData", "要素の材料番号が範囲外です。要素=" & CStr(ii + 1)
        End If
        Elem(ii).IsJoint = P3MaterialIsJoint(Elem(ii).MatNo)
        If Elem(ii).IsJoint Then P3HasJointElements = True
        For j = 0 To 7
            If P6InputElementNode(i, j + 1) <= 0 Then
                Elem(ii).node(j) = -1
                Elem(ii).x(j) = 0#
                Elem(ii).y(j) = 0#
                Elem(ii).ElNode(2 * j) = -1
                Elem(ii).ElNode(2 * j + 1) = -1
            Else
                originalNode = P6InputElementNode(i, j + 1) - 1
                Elem(ii).node(j) = FEMGetFreeNode(originalNode)
                Elem(ii).x(j) = x(originalNode): Elem(ii).y(j) = y(originalNode)
                Elem(ii).ElNode(2 * j) = 2 * Elem(ii).node(j)
                Elem(ii).ElNode(2 * j + 1) = 2 * Elem(ii).node(j) + 1
            End If
        Next j
        If Not Elem(ii).IsJoint Then P3EnsureElementCounterClockwise Elem(ii)
    Next i
    If SuppressUserMessages Then SaveP0Progress "element_topology_complete"

    For i = 1 To NumberOfNode
        ax = P6InputNodeId(i) - 1
        If ax < 0 Or ax > NAB Then Err.Raise vbObjectError + 3438, "FEM.SetElementData", "節点番号が範囲外です。"
        ii = FEMFreeNodeMap(ax)
        If ii < 0 Then Err.Raise vbObjectError + 3439, "FEM.SetElementData", "要素に接続していない節点があります。"
        ix = 2 * ii
        XXX(ix) = P6InputNodeX(i): XXX(ix + 1) = P6InputNodeY(i)
        NodeCond(ix) = P6InputNodeCondX(i): NodeCond(ix + 1) = P6InputNodeCondY(i)
        Disp(ix) = P6InputNodeDispX(i): Disp(ix + 1) = P6InputNodeDispY(i)
        Force(ix) = P6InputNodeForceX(i): Force(ix + 1) = P6InputNodeForceY(i)
    Next i
    P3BumpConstraintGeneration
    If SuppressUserMessages Then SaveP0Progress "element_node_load_complete"
    If SuppressUserMessages Then SaveP0Progress "element_q8_node_data_complete"
    P6InvalidateCSRConstraintCache
    P6PrepareMeshCaches
End Sub

Private Function P6ComputeTopologySignature() As String
    ' 入力読込時に確定した整数世代を使い、解析開始時の要素走査を行わない。
    P6ComputeTopologySignature = CStr(P6MeshRevision)
End Function

Public Sub P6PrepareMeshCaches()
    Dim topologySignature As String

    topologySignature = P6ComputeTopologySignature()
    If topologySignature <> P6RCMCacheSignature Then
        P6RCMCacheSignature = topologySignature
        P6RCMCacheReady = False
        P6RCMMapReady = False
        P6RCMCacheOriginalBandwidth = 0
        P6RCMCacheCandidateBandwidth = 0
        Erase P6RCMOldToNew
        Erase P6RCMNewToOld
        P6MeshInvariantReady = False
        P6MeshInvariantSignature = vbNullString
        P6MeshInvariantOriginalBandwidth = 0
        P6MeshInvariantCandidateBandwidth = 0
        P6MeshInvariantRCMApplied = False
        P6MeshInvariantSymmetryError = 0#
        P6TangentSignature = vbNullString
        P6CSRNumericGeneration = -1
        P6CSRMaterialGeneration = -1
        P6CSRNumericAssemblyStatus = vbNullString
        P6CSRRebuildCount = 0
        P6CSRReuseCount = 0
        P6CSRLastAssemblyPath = vbNullString
        P6SolverSelectionReady = False
        P6SolverEvaluationStatus = "mesh_changed"

        P6CSRPatternReady = False
        P6CSRPatternSignature = vbNullString
        Erase P6CSRRowPtr
        Erase P6CSRColumnIndex
        Erase P6CSRDiagonalPosition
        Erase P6CSRElementPosition
        Erase P6CSRBoundaryZeroPosition
        Erase P6CSRBoundaryDiagonalPosition
        Erase P6CSRBoundaryRHSPosition
        Erase P6CSRBoundaryRHSRow
        Erase P6CSRBoundaryRHSColumn
        P6CSRScatterCount = 0
        P6CSRScatterElementCount = -1
        P6CSRScatterTangentGeneration = -1
        P6CSRScatterMaterialGeneration = -1
        P6CSRScatterPatternSignature = vbNullString
        Erase P6CSRScatterStart
        Erase P6CSRScatterLocalRow
        Erase P6CSRScatterLocalColumn
        Erase P6CSRScatterPosition
        Erase P6CSRScatterValueIndex
        P6CSRNNZ = 0
        P6CSRValueCapacity = 0
        P6CSRStorageBytes = 0#
        P6CSRBoundaryConstraintVersion = -1
        P6CSRBoundaryZeroCount = 0
        P6CSRBoundaryDiagonalCount = 0
        P6CSRBoundaryRHSCount = 0
        P6CSRDiagonalInverseReady = False
        Erase P6CSRDiagonalInverse
        Erase P6CSROriginalValues
        Erase P6CSRValues
    Else
        P6RCMMapReady = False
    End If
    P6CSRReady = False
    P6CSRBoundaryApplied = False
    P6CSRDiagonalInverseReady = False
End Sub

Sub SetBmat(ByRef elementData As Element_Data, Optional ByVal elementId As Long = -1)
Dim im As Long: Dim jm As Long
Dim i As Long, j As Long, k As Long
  'Bマトリックスの計算
  Dim Jmat(0 To 1, 0 To 1, 0 To 3) As Double
  Dim JmatT(0 To 1, 0 To 1, 0 To 3) As Double
  Dim JTF(0 To 15, 0 To 3) As Double
  Dim coordX(0 To 7) As Double, coordY(0 To 7) As Double
  Dim localDof(0 To 15) As Long
  Dim minX As Double, maxX As Double, minY As Double, maxY As Double
  Dim repLen As Double, jTol As Double, detJ As Double

  P6EnsureGaussCache
  minX = elementData.x(0): maxX = elementData.x(0)
  minY = elementData.y(0): maxY = elementData.y(0)
  For i = 0 To 7
    coordX(i) = elementData.x(i)
    coordY(i) = elementData.y(i)
    If coordX(i) < minX Then minX = coordX(i)
    If coordX(i) > maxX Then maxX = coordX(i)
    If coordY(i) < minY Then minY = coordY(i)
    If coordY(i) > maxY Then maxY = coordY(i)
    localDof(2 * i) = 2 * elementData.node(i)
    localDof(2 * i + 1) = localDof(2 * i) + 1
    elementData.ElNode(2 * i) = localDof(2 * i)
    elementData.ElNode(2 * i + 1) = localDof(2 * i + 1)
  Next i
  repLen = maxX - minX
  If maxY - minY > repLen Then repLen = maxY - minY
  If repLen < 0.000001 Then repLen = 0.000001
  jTol = 0.00000001 * repLen * repLen
  If jTol < 1E-16 Then jTol = 1E-16
  Erase Jmat
  For i = 0 To 3
        For im = 0 To 7 'im X
            jm = im + 8 'jm Y
            'N/∂ξ 0 to 7
            'N/∂η 8 to 15
          Jmat(0, 0, i) = Jmat(0, 0, i) + P6GaussRN(i, im) * coordX(im) ' ∂x/∂ξ=Σ∂N/∂ξ X
          Jmat(0, 1, i) = Jmat(0, 1, i) + P6GaussRN(i, im) * coordY(im) ' ∂y/∂ξ= Σ∂N/∂ξ Y
          Jmat(1, 0, i) = Jmat(1, 0, i) + P6GaussRN(i, jm) * coordX(im) ' ∂x/∂η= Σ∂N/∂η X
          Jmat(1, 1, i) = Jmat(1, 1, i) + P6GaussRN(i, jm) * coordY(im) ' ∂y/∂η= Σ∂N/∂η Y
        Next im
        detJ = Jmat(0, 0, i) * Jmat(1, 1, i) - Jmat(0, 1, i) * Jmat(1, 0, i)
        elementData.dj(i) = detJ
        If detJ <= jTol Then
          SetAnalysisFailure RESULT_INPUT_ERROR, "符号付きJacobianがGauss点で正でないか小さすぎます。要素=" & CStr(elementId + 1) & " 積分点=" & CStr(i) & " detJ=" & Format$(detJ, "0.000E+00"), vbObjectError + 3024, elementId + 1, i, CurrentIncrement, CurrentIteration
          Err.Raise vbObjectError + 3024, "FEM.SetBmat", AnalysisMessage
        End If
        JmatT(0, 0, i) = Jmat(1, 1, i) / detJ '∂ξ/∂x
        JmatT(0, 1, i) = -Jmat(0, 1, i) / detJ '∂η/∂x
        JmatT(1, 0, i) = -Jmat(1, 0, i) / detJ '∂ξ/∂y
        JmatT(1, 1, i) = Jmat(0, 0, i) / detJ '∂η/∂y
        For im = 0 To 7
           jm = im + 8
             JTF(im, i) = JmatT(0, 0, i) * P6GaussRN(i, im) + JmatT(0, 1, i) * P6GaussRN(i, jm) '∂N/∂x
             JTF(jm, i) = JmatT(1, 0, i) * P6GaussRN(i, im) + JmatT(1, 1, i) * P6GaussRN(i, jm) '∂N/∂y
       Next im
       For j = 0 To 7
          im = 2 * j
          
            elementData.Bmat(0, im, i) = -JTF(j, i) '∂N/∂x
            elementData.Bmat(1, im + 1, i) = -JTF(j + 8, i) '∂N/∂y
            elementData.Bmat(2, im, i) = -JTF(j + 8, i) '∂N/∂y
            elementData.Bmat(2, im + 1, i) = -JTF(j, i) '∂N/∂x
       
       Next j
  Next i
End Sub
Sub Gauss_integral(RN() As Double, n() As Double)
  P6EnsureGaussCache
  P6CopyGaussCache RN, n
End Sub

Public Sub P6EnsureGaussCache()
  Dim rawRN(0 To 3, 0 To 15) As Double, rawN(0 To 3, 0 To 7) As Double
  Dim i As Long, j As Long
  If P6GaussCacheReady Then Exit Sub
  P6ComputeGaussIntegral rawRN, rawN
  For i = 0 To 3
    For j = 0 To 15
      P6GaussRN(i, j) = rawRN(i, j)
    Next j
    For j = 0 To 7
      P6GaussN(i, j) = rawN(i, j)
    Next j
  Next i
  P6GaussCacheReady = True
End Sub

Private Sub P6CopyGaussCache(ByRef RN() As Double, ByRef n() As Double)
  Dim i As Long, j As Long
  For i = 0 To 3
    For j = 0 To 15
      RN(i, j) = P6GaussRN(i, j)
    Next j
    For j = 0 To 7
      n(i, j) = P6GaussN(i, j)
    Next j
  Next i
End Sub

Public Sub FEMGaussPhysicalXY(ByVal elementId As Long, ByVal gaussId As Long, ByRef xValue As Double, ByRef yValue As Double)
  Dim nodeNo As Long
  xValue = 0#
  yValue = 0#
  If elementId < 0 Or elementId >= NumberOfElement Then Exit Sub
  If gaussId < 0 Or gaussId > 3 Then Exit Sub
  P6EnsureGaussCache
  For nodeNo = 0 To 7
    xValue = xValue + P6GaussN(gaussId, nodeNo) * Elem(elementId).x(nodeNo)
    yValue = yValue + P6GaussN(gaussId, nodeNo) * Elem(elementId).y(nodeNo)
  Next nodeNo
End Sub

Private Sub P6ComputeGaussIntegral(ByRef RN() As Double, ByRef n() As Double)
 Dim pxi(0 To 3) As Double, pea(0 To 3) As Double
 Dim i As Long
 Erase pxi, pea
 pxi(0) = -0.5773502692
 pxi(1) = 0.5773502692
 pea(0) = pxi(0): pea(1) = pxi(0)
 pxi(2) = pxi(1): pea(2) = pxi(1): pxi(3) = pxi(0): pea(3) = pxi(1)
 'pxi ξ
 'pea η
 
 For i = 0 To 3
 
   n(i, 0) = 0.25 * (1# - pxi(i)) * (1# - pea(i)) * (-1# - pxi(i) - pea(i))  'N1=1/4(1-ξ)(1-η)(-1-ξ-η)
   n(i, 1) = 0.25 * (1# + pxi(i)) * (1# - pea(i)) * (-1# + pxi(i) - pea(i)) 'N2=1/4(1+ξ)(1-η)(-1+ξ-η)
   n(i, 2) = 0.25 * (1# + pxi(i)) * (1# + pea(i)) * (-1# + pxi(i) + pea(i)) 'N3=1/4(1+ξ)(1+η)(-1+ξ+η)
   n(i, 3) = 0.25 * (1# - pxi(i)) * (1# + pea(i)) * (-1# - pxi(i) + pea(i)) 'N4=1/4(1-ξ)(1+η)(-1-ξ+η)
   n(i, 4) = 0.5 * (1# - pxi(i) ^ 2) * (1# - pea(i)) 'N5=1/2(1-ξ*ξ)(1-η)
   n(i, 5) = 0.5 * (1# + pxi(i)) * (1# - pea(i) ^ 2) 'N6=1/2(1+ξ)(1-η*η)
   n(i, 6) = 0.5 * (1# - pxi(i) ^ 2) * (1# + pea(i)) 'N7=1/2(1-ξ*ξ)(1+η)
   n(i, 7) = 0.5 * (1# - pxi(i)) * (1# - pea(i) ^ 2) 'N8=1/2(1-ξ)(1-η*η)
   
   RN(i, 0) = 0.25 * (1# - pea(i)) * (2 * pxi(i) + pea(i))  '∂N1/∂ξ=1/4(1-η)(2ξ+η)
   RN(i, 1) = 0.25 * (1# - pea(i)) * (2 * pxi(i) - pea(i))  '∂N2/∂ξ=1/4(1-η)(2ξ-η)
   RN(i, 2) = 0.25 * (1# + pea(i)) * (2 * pxi(i) + pea(i))  '∂N3/∂ξ=1/4(1+η)(2ξ+η)
   RN(i, 3) = 0.25 * (1# + pea(i)) * (2 * pxi(i) - pea(i))  '∂N4/∂ξ=1/4(1+η)(2ξ-η)
   RN(i, 4) = -pxi(i) * (1# - pea(i))                       '∂N5/∂ξ=-ξ(1-η)
   RN(i, 5) = 0.5 * (1# - pea(i) ^ 2)                       '∂N6/∂ξ=1/2(1-η*η)
   RN(i, 6) = -pxi(i) * (1# + pea(i))                       '∂N7/∂ξ=-ξ(1+η)
   RN(i, 7) = -0.5 * (1# - pea(i) ^ 2)                      '∂N8/∂ξ=-1/2(1-η*η)
   
   RN(i, 8) = 0.25 * (1# - pxi(i)) * (2 * pea(i) + pxi(i))  '∂N1/∂η=1/4(1-η)(2ξ+η)
   RN(i, 9) = 0.25 * (1# + pxi(i)) * (2 * pea(i) - pxi(i))  '∂N2/∂η=1/4(1+η)(2ξ-η)
   RN(i, 10) = 0.25 * (1# + pxi(i)) * (2 * pea(i) + pxi(i)) '∂N3/∂η=1/4(1+η)(2ξ+η)
   RN(i, 11) = 0.25 * (1# - pxi(i)) * (2 * pea(i) - pxi(i)) '∂N4/∂η=1/4(1-η)(2ξ-η)
   RN(i, 12) = -0.5 * (1# - pxi(i) ^ 2)                     '∂N5/∂η=-1/2(1-ξ*ξ)
   RN(i, 13) = -pea(i) * (1# + pxi(i))                      '∂N6/∂η=-η(1+ξ)
   RN(i, 14) = 0.5 * (1# - pxi(i) ^ 2)                      '∂N7/∂η=1/2(1-ξ*ξ)
   RN(i, 15) = -pea(i) * (1# - pxi(i))                      '∂N8/∂η=-η(1-ξ)
Next i
 
End Sub
Sub SetDmat(young As Double, Poisson As Double, Dmat() As Double)
Dim a As Double: Dim b As Double: Dim c As Double
'弾性定数行列。地盤用途の既定は平面ひずみ。
  If young <= 0# Then
    Err.Raise vbObjectError + 3025, "FEM.SetDmat", "ヤング率は正で指定してください。"
  End If
  If Poisson <= -1# Or Poisson >= 0.5 Then
    Err.Raise vbObjectError + 3026, "FEM.SetDmat", "ポアソン比は-1より大きく0.5未満で指定してください。値=" & Format$(Poisson, "0.##########")
  End If
  P1Formulation = P1_FORMULATION_PLANE_STRAIN
  a = young * (1# - Poisson) / ((1# + Poisson) * (1# - 2# * Poisson))
  b = young * Poisson / ((1# + Poisson) * (1# - 2# * Poisson))
  c = 0.5 * young / (1# + Poisson)
  Dmat(0, 0) = a
  Dmat(0, 1) = b
  Dmat(0, 2) = 0#
  Dmat(1, 0) = b
  Dmat(1, 1) = a
  Dmat(1, 2) = 0#
  Dmat(2, 0) = 0#
  Dmat(2, 1) = 0#
  Dmat(2, 2) = c
End Sub
Sub SetSmat(Dmat() As Double, Bmat() As Double, Smat() As Double)
Dim im As Long
Dim i As Long, j As Long, k As Long
Dim s As Double
Dim l As Long
'S行列設定
   For im = 0 To 3
     For i = 0 To 2
        For j = 0 To 15
           s = 0#
         If i = 2 And j = 14 Then
          l = 0
          
         End If
         
           For k = 0 To 2
             s = s + Dmat(i, k) * Bmat(k, j, im) 'D行列とB行列の積　［S］=［D］［B］
           Next k
           Smat(i, j, im) = s
        Next j
     Next i
   Next im
End Sub
Sub SetElmStiffness(thickness As Double, weight As Double, Bmat() As Double, Smat() As Double, kmat() As Double, ElmWeight As Double, dj() As Double, n() As Double)
Dim im As Long
Dim i As Long, j As Long, k As Long
Dim s As Double
'要素剛性行列作成
Erase kmat
   ElmWeight = 0#
   For im = 0 To 3
     For i = 0 To 15
        For j = 0 To 15
          
         
          s = 0#
          For k = 0 To 2
             s = s + Bmat(k, i, im) * Smat(k, j, im) 'B行列転地とS行列の積　［K］＝［B］T［S］
          Next k
          kmat(i, j) = kmat(i, j) + s * dj(im) * thickness '［K］=［K］・V
        Next j
        If i <= 7 Then
          ElmWeight = ElmWeight + weight * thickness * n(im, i) * dj(im)
        End If


     Next i
   Next im
End Sub
Private Function P6ReadHourglassFactor() As Double
  Dim factorValue As Double
  factorValue = P6ReadSetting("Q8_HOURGLASS_FACTOR", P6_HOURGLASS_FACTOR_DEFAULT)
  If factorValue < 0# Then factorValue = 0#
  If factorValue > 1# Then factorValue = 1#
  P6HourglassFactor = factorValue
  P6ReadHourglassFactor = factorValue
End Function

Public Function P6InvertDense(ByRef a() As Double, ByVal n As Long) As Boolean
  Dim i As Long, j As Long, k As Long, pivotRow As Long
  Dim pivotValue As Double, bestValue As Double, swapValue As Double, scaleValue As Double
  Dim ipiv() As Long, inv() As Double
  P6InvertDense = False
  If n < 1 Then Exit Function
  ReDim ipiv(0 To n - 1)
  ReDim inv(0 To n - 1, 0 To n - 1)
  For i = 0 To n - 1
    ipiv(i) = i
    inv(i, i) = 1#
  Next i
  For k = 0 To n - 1
    pivotRow = k
    bestValue = Abs(a(k, k))
    For i = k + 1 To n - 1
      If Abs(a(i, k)) > bestValue Then
        bestValue = Abs(a(i, k))
        pivotRow = i
      End If
    Next i
    If bestValue <= 1E-30 Then Exit Function
    If pivotRow <> k Then
      For j = 0 To n - 1
        swapValue = a(k, j): a(k, j) = a(pivotRow, j): a(pivotRow, j) = swapValue
        swapValue = inv(k, j): inv(k, j) = inv(pivotRow, j): inv(pivotRow, j) = swapValue
      Next j
    End If
    pivotValue = a(k, k)
    For j = 0 To n - 1
      a(k, j) = a(k, j) / pivotValue
      inv(k, j) = inv(k, j) / pivotValue
    Next j
    For i = 0 To n - 1
      If i <> k Then
        scaleValue = a(i, k)
        If scaleValue <> 0# Then
          For j = 0 To n - 1
            a(i, j) = a(i, j) - scaleValue * a(k, j)
            inv(i, j) = inv(i, j) - scaleValue * inv(k, j)
          Next j
        End If
      End If
    Next i
  Next k
  For i = 0 To n - 1
    For j = 0 To n - 1
      a(i, j) = inv(i, j)
    Next j
  Next i
  P6InvertDense = True
End Function

Public Sub P6AddQ8HourglassStabilization(ByRef elementData As Element_Data, ByVal thickness As Double, ByVal addToStiffness As Boolean)
  Dim gaussId As Long, strainId As Long, dofId As Long, seedId As Long, modeId As Long
  Dim nodeId As Long, otherId As Long, linId As Long, linCount As Long
  Dim areaValue As Double, shearModulus As Double, scaleValue As Double
  Dim gRow As Long, nrmValue As Double, projValue As Double, xc As Double, yc As Double
  Dim gMat(0 To 11, 0 To 15) As Double
  Dim ggt(0 To 11, 0 To 11) As Double
  Dim seedVec(0 To 15) As Double, modeVec(0 To 15) As Double, workVec(0 To 15) As Double
  Dim linVec(0 To 5, 0 To 15) As Double
  Dim yVec(0 To 11) As Double
  Dim reuseModes As Boolean
  Dim t0 As Double
  t0 = Timer

  If elementData.IsJoint Then
    elementData.HourglassModeCount = 0
    elementData.HourglassScale = 0#
    elementData.HourglassKReady = False
    elementData.HourglassKScaledReady = False
    Exit Sub
  End If
  If P6HourglassFactor <= 0# Or thickness <= 0# Then
    elementData.HourglassModeCount = 0
    elementData.HourglassScale = 0#
    elementData.HourglassKReady = False
    elementData.HourglassKScaledReady = False
    P6ProfHgMs = P6ProfHgMs + P6ElapsedMs(t0)
    Exit Sub
  End If
  reuseModes = (elementData.HourglassModeCount > 0 And elementData.HourglassShapeGen = P6HgShapeGeneration)
  If reuseModes Then
    P6HourglassReuseCount = P6HourglassReuseCount + 1
    GoTo ApplyHgScale
  End If
  elementData.HourglassModeCount = 0
  elementData.HourglassScale = 0#
  elementData.HourglassKReady = False
  elementData.HourglassKScaledReady = False
  For modeId = 0 To 3
    For dofId = 0 To 15
      elementData.HourglassMode(modeId, dofId) = 0#
    Next dofId
  Next modeId
  areaValue = 0#
  For gaussId = 0 To 3
    areaValue = areaValue + elementData.dj(gaussId)
  Next gaussId
  If Abs(areaValue) <= 1E-30 Then
    P6ProfHgMs = P6ProfHgMs + P6ElapsedMs(t0)
    Exit Sub
  End If

  For gaussId = 0 To 3
    For strainId = 0 To 2
      gRow = gaussId * 3 + strainId
      For dofId = 0 To 15
        gMat(gRow, dofId) = elementData.Bmat(strainId, dofId, gaussId) * elementData.dj(gaussId)
      Next dofId
    Next strainId
  Next gaussId
  For gRow = 0 To 11
    For strainId = 0 To 11
      projValue = 0#
      For dofId = 0 To 15
        projValue = projValue + gMat(gRow, dofId) * gMat(strainId, dofId)
      Next dofId
      ggt(gRow, strainId) = projValue
    Next strainId
  Next gRow
  If Not P6InvertDense(ggt, 12) Then
    P6ProfHgMs = P6ProfHgMs + P6ElapsedMs(t0)
    Exit Sub
  End If

  xc = 0#: yc = 0#
  For nodeId = 0 To 7
    xc = xc + elementData.x(nodeId)
    yc = yc + elementData.y(nodeId)
  Next nodeId
  xc = xc / 8#: yc = yc / 8#
  For nodeId = 0 To 7
    linVec(0, 2 * nodeId) = 1#: linVec(0, 2 * nodeId + 1) = 0#
    linVec(1, 2 * nodeId) = 0#: linVec(1, 2 * nodeId + 1) = 1#
    linVec(2, 2 * nodeId) = -(elementData.y(nodeId) - yc)
    linVec(2, 2 * nodeId + 1) = elementData.x(nodeId) - xc
    linVec(3, 2 * nodeId) = elementData.x(nodeId) - xc
    linVec(3, 2 * nodeId + 1) = 0#
    linVec(4, 2 * nodeId) = 0#
    linVec(4, 2 * nodeId + 1) = elementData.y(nodeId) - yc
    linVec(5, 2 * nodeId) = elementData.y(nodeId) - yc
    linVec(5, 2 * nodeId + 1) = elementData.x(nodeId) - xc
  Next nodeId
  linCount = 0
  For linId = 0 To 5
    For modeId = 0 To linCount - 1
      projValue = 0#
      For dofId = 0 To 15
        projValue = projValue + linVec(linId, dofId) * linVec(modeId, dofId)
      Next dofId
      For dofId = 0 To 15
        linVec(linId, dofId) = linVec(linId, dofId) - projValue * linVec(modeId, dofId)
      Next dofId
    Next modeId
    nrmValue = 0#
    For dofId = 0 To 15
      nrmValue = nrmValue + linVec(linId, dofId) * linVec(linId, dofId)
    Next dofId
    If nrmValue > 1E-20 Then
      nrmValue = Sqr(nrmValue)
      If linCount <> linId Then
        For dofId = 0 To 15
          linVec(linCount, dofId) = linVec(linId, dofId) / nrmValue
        Next dofId
      Else
        For dofId = 0 To 15
          linVec(linCount, dofId) = linVec(linCount, dofId) / nrmValue
        Next dofId
      End If
      linCount = linCount + 1
    End If
  Next linId
  If linCount < 3 Then
    P6ProfHgMs = P6ProfHgMs + P6ElapsedMs(t0)
    Exit Sub
  End If

  For seedId = 0 To 15
    For dofId = 0 To 15
      seedVec(dofId) = 0#
    Next dofId
    seedVec(seedId) = 1#
    For gRow = 0 To 11
      projValue = 0#
      For dofId = 0 To 15
        projValue = projValue + gMat(gRow, dofId) * seedVec(dofId)
      Next dofId
      yVec(gRow) = projValue
    Next gRow
    For gRow = 0 To 11
      workVec(gRow) = 0#
      For strainId = 0 To 11
        workVec(gRow) = workVec(gRow) + ggt(gRow, strainId) * yVec(strainId)
      Next strainId
    Next gRow
    For dofId = 0 To 15
      projValue = 0#
      For gRow = 0 To 11
        projValue = projValue + gMat(gRow, dofId) * workVec(gRow)
      Next gRow
      modeVec(dofId) = seedVec(dofId) - projValue
    Next dofId
    For linId = 0 To linCount - 1
      projValue = 0#
      For dofId = 0 To 15
        projValue = projValue + modeVec(dofId) * linVec(linId, dofId)
      Next dofId
      For dofId = 0 To 15
        modeVec(dofId) = modeVec(dofId) - projValue * linVec(linId, dofId)
      Next dofId
    Next linId
    For modeId = 0 To elementData.HourglassModeCount - 1
      projValue = 0#
      For dofId = 0 To 15
        projValue = projValue + modeVec(dofId) * elementData.HourglassMode(modeId, dofId)
      Next dofId
      For dofId = 0 To 15
        modeVec(dofId) = modeVec(dofId) - projValue * elementData.HourglassMode(modeId, dofId)
      Next dofId
    Next modeId
    nrmValue = 0#
    For dofId = 0 To 15
      nrmValue = nrmValue + modeVec(dofId) * modeVec(dofId)
    Next dofId
    If nrmValue > 0.00000001 Then
      nrmValue = Sqr(nrmValue)
      If elementData.HourglassModeCount > 3 Then Exit For
      For dofId = 0 To 15
        elementData.HourglassMode(elementData.HourglassModeCount, dofId) = modeVec(dofId) / nrmValue
      Next dofId
      elementData.HourglassModeCount = elementData.HourglassModeCount + 1
      P6HourglassModeCount = P6HourglassModeCount + 1
    End If
  Next seedId
  If elementData.HourglassModeCount < 1 Then
    P6ProfHgMs = P6ProfHgMs + P6ElapsedMs(t0)
    Exit Sub
  End If
  elementData.HourglassShapeGen = P6HgShapeGeneration

ApplyHgScale:
  shearModulus = Abs(elementData.Dmat(2, 2))
  If shearModulus <= 1E-30 Then
    P6ProfHgMs = P6ProfHgMs + P6ElapsedMs(t0)
    Exit Sub
  End If
  ' モードは無次元正規化。人工剛性は物理剛性と同じ E t 次元（面積を掛けない）
  scaleValue = P6HourglassFactor * shearModulus * thickness
  elementData.HourglassScale = scaleValue
  If Not addToStiffness Then
    P6ProfHgMs = P6ProfHgMs + P6ElapsedMs(t0)
    Exit Sub
  End If
  If Not elementData.HourglassKReady Then
    For dofId = 0 To 15
      For otherId = 0 To 15
        elementData.HourglassK(dofId, otherId) = 0#
      Next otherId
    Next dofId
    For modeId = 0 To elementData.HourglassModeCount - 1
      For dofId = 0 To 15
        For otherId = 0 To 15
          elementData.HourglassK(dofId, otherId) = elementData.HourglassK(dofId, otherId) + elementData.HourglassMode(modeId, dofId) * elementData.HourglassMode(modeId, otherId)
        Next otherId
      Next dofId
    Next modeId
    elementData.HourglassKReady = True
    elementData.HourglassKScaledReady = False
  End If
  If (Not elementData.HourglassKScaledReady) Or Abs(elementData.HourglassCachedScale - scaleValue) > 0# Then
    For dofId = 0 To 15
      For otherId = 0 To 15
        elementData.HourglassKScaled(dofId, otherId) = scaleValue * elementData.HourglassK(dofId, otherId)
      Next otherId
    Next dofId
    elementData.HourglassCachedScale = scaleValue
    elementData.HourglassKScaledReady = True
  End If
  For dofId = 0 To 15
    For otherId = 0 To 15
      projValue = elementData.HourglassKScaled(dofId, otherId)
      If Abs(projValue) > 0# Then
        elementData.kmat(dofId, otherId) = elementData.kmat(dofId, otherId) + projValue
      End If
    Next otherId
  Next dofId
  P6ProfHgMs = P6ProfHgMs + P6ElapsedMs(t0)
End Sub
Private Function CornerJacobian(ByRef elementData As Element_Data, ByVal xi As Double, ByVal eta As Double) As Double
  Dim dNxi(0 To 7) As Double, dNeta(0 To 7) As Double
  Dim dXdxi As Double, dYdxi As Double, dXdeta As Double, dYdeta As Double
  Dim i As Long
  dNxi(0) = 0.25 * (1# - eta) * (2# * xi + eta)
  dNeta(0) = 0.25 * (1# - xi) * (xi + 2# * eta)
  dNxi(1) = 0.25 * (1# - eta) * (2# * xi - eta)
  dNeta(1) = 0.25 * (1# + xi) * (-xi + 2# * eta)
  dNxi(2) = 0.25 * (1# + eta) * (2# * xi + eta)
  dNeta(2) = 0.25 * (1# + xi) * (xi + 2# * eta)
  dNxi(3) = 0.25 * (1# + eta) * (2# * xi - eta)
  dNeta(3) = 0.25 * (1# - xi) * (-xi + 2# * eta)
  dNxi(4) = -xi * (1# - eta)
  dNeta(4) = -0.5 * (1# - xi * xi)
  dNxi(5) = 0.5 * (1# - eta * eta)
  dNeta(5) = -eta * (1# + xi)
  dNxi(6) = -xi * (1# + eta)
  dNeta(6) = 0.5 * (1# - xi * xi)
  dNxi(7) = -0.5 * (1# - eta * eta)
  dNeta(7) = -eta * (1# - xi)
  For i = 0 To 7
    dXdxi = dXdxi + dNxi(i) * elementData.x(i)
    dYdxi = dYdxi + dNxi(i) * elementData.y(i)
    dXdeta = dXdeta + dNeta(i) * elementData.x(i)
    dYdeta = dYdeta + dNeta(i) * elementData.y(i)
  Next i
  CornerJacobian = dXdxi * dYdeta - dYdxi * dXdeta
End Function

Private Function CheckElementJacobian(ByRef elementData As Element_Data, ByVal elementId As Long, Optional ByVal checkGauss As Boolean = True) As Boolean
  Dim cornerXi(0 To 3) As Double, cornerEta(0 To 3) As Double
  Dim cornerId As Long, gaussId As Long
  Dim determinant As Double
  cornerXi(0) = -1#: cornerXi(1) = 1#: cornerXi(2) = 1#: cornerXi(3) = -1#
  cornerEta(0) = -1#: cornerEta(1) = -1#: cornerEta(2) = 1#: cornerEta(3) = 1#
  CheckElementJacobian = False
  For cornerId = 0 To 3
    determinant = CornerJacobian(elementData, cornerXi(cornerId), cornerEta(cornerId))
    If determinant < MinDetJCorner Then MinDetJCorner = determinant
    If determinant > MaxDetJCorner Then MaxDetJCorner = determinant
    If determinant <= 0.000000000001 Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "符号付きJacobianが隅点で正ではありません。要素=" & CStr(elementId + 1), vbObjectError + 3023, elementId + 1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
  Next cornerId
  If Not checkGauss Then
    CheckElementJacobian = True
    Exit Function
  End If
  For gaussId = 0 To 3
    determinant = elementData.dj(gaussId)
    If determinant < MinDetJGauss Then MinDetJGauss = determinant
    If determinant > MaxDetJGauss Then MaxDetJGauss = determinant
    If determinant <= 0.000000000001 Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "符号付きJacobianがGauss点で正ではありません。要素=" & CStr(elementId + 1) & "、積分点=" & CStr(gaussId), vbObjectError + 3024, elementId + 1, gaussId, CurrentIncrement, CurrentIteration
      Exit Function
    End If
  Next gaussId
  CheckElementJacobian = True
End Function

Function SetElmMat() As Boolean
  Dim i As Long, j As Long, k As Long, im As Long, materialIndex As Long, cacheBase As Long
  Dim tangentSignature As String
  Dim reuseElasticK As Boolean
  Dim reuseElasticOperator As Boolean
  SetElmMat = False
  P6EnsureGaussCache
  P6ReadHourglassFactor
  tangentSignature = P6ComputeTangentSignature()
  reuseElasticK = P6ElasticKCacheReady And P6ElasticKCacheElementCount = NumberOfElement And P6ElasticKCacheSignature = tangentSignature
  If Not P6ElasticKCacheReady Then
    P6CacheMissReason = "not_ready"
  ElseIf P6ElasticKCacheElementCount <> NumberOfElement Then
    P6CacheMissReason = "count " & CStr(P6ElasticKCacheElementCount) & "->" & CStr(NumberOfElement)
  ElseIf P6ElasticKCacheSignature <> tangentSignature Then
    P6CacheMissReason = "signature " & P6ElasticKCacheSignature & " -> " & tangentSignature
  ElseIf Not P6ElasticWeightCacheUsable() Then
    P6CacheMissReason = "missing_weight"
  Else
    P6CacheMissReason = "reuse"
  End If
  If reuseElasticK Then
    If P6ElasticWeightCacheUsable() Then
      P6ElasticKCacheStatus = "reuse"
    Else
      reuseElasticK = False
      P6ElasticKCacheStatus = "rebuild_missing_weight"
    End If
  Else
    P6ElasticKCacheStatus = "build"
  End If
  reuseElasticOperator = reuseElasticK And P6ElasticOperatorCacheReady
  If Not reuseElasticOperator Then
    P6HgShapeGeneration = P6HgShapeGeneration + 1
    If P6HgShapeGeneration <= 0 Then P6HgShapeGeneration = 1
    P6HourglassModeCount = 0
  End If
  For i = 0 To NumberOfElement - 1
    If Elem(i).IsJoint Then
      P3JointBuild i
    Else
    With Elem(i)
      materialIndex = .MatNo
      If Not reuseElasticOperator Then
        If Material(materialIndex).MaterialCacheReady Then
          .Dmat(0, 0) = Material(materialIndex).ElasticD00
          .Dmat(0, 1) = Material(materialIndex).ElasticD01
          .Dmat(0, 2) = 0#
          .Dmat(1, 0) = Material(materialIndex).ElasticD01
          .Dmat(1, 1) = Material(materialIndex).ElasticD00
          .Dmat(1, 2) = 0#
          .Dmat(2, 0) = 0#
          .Dmat(2, 1) = 0#
          .Dmat(2, 2) = Material(materialIndex).ElasticD22
        Else
          SetDmat Material(materialIndex).young, Material(materialIndex).Poisson, .Dmat
        End If
        If Not CheckElementJacobian(Elem(i), i, False) Then Exit Function
        Call SetBmat(Elem(i), i)
        If Not CheckElementJacobian(Elem(i), i, True) Then Exit Function
        SetSmat .Dmat, .Bmat, .Smat
        For im = 0 To 3
          For j = 0 To 2
            For k = 0 To 2
              .Dpmat(j, k, im) = .Dmat(j, k)
            Next k
            For k = 0 To 15
              .Spmat(j, k, im) = .Smat(j, k, im)
            Next k
          Next j
        Next im
      End If
      If reuseElasticK Then
        cacheBase = i * 256
        For j = 0 To 15
          For k = 0 To 15
            .kmat(j, k) = P6ElasticKCache(cacheBase + j * 16 + k)
          Next k
        Next j
        .ElmWeight = P6ElasticElementWeightCache(i)
        Call SetBmat(Elem(i), i)
        P6AddQ8HourglassStabilization Elem(i), Material(materialIndex).thickness, False
      Else
        SetElmStiffness Material(materialIndex).thickness, Material(materialIndex).weight, .Bmat, .Smat, .kmat, .ElmWeight, .dj, P6GaussN
        P6AddQ8HourglassStabilization Elem(i), Material(materialIndex).thickness, True
      End If
      .TangentDirty = False
    End With
    End If
  Next i
  If Not reuseElasticK Then
    ReDim P6ElasticKCache(0 To NumberOfElement * 256 - 1)
    ReDim P6ElasticElementWeightCache(0 To NumberOfElement - 1)
    For i = 0 To NumberOfElement - 1
      cacheBase = i * 256
      P6ElasticElementWeightCache(i) = Elem(i).ElmWeight
      For j = 0 To 15
        For k = 0 To 15
          P6ElasticKCache(cacheBase + j * 16 + k) = Elem(i).kmat(j, k)
        Next k
      Next j
    Next i
    P6ElasticKCacheElementCount = NumberOfElement
    P6ElasticKCacheSignature = tangentSignature
    P6ElasticWeightCacheReady = True
  End If
  P6ElasticKCacheReady = True
  P6ElasticOperatorCacheReady = True
  P6QuadratureArea = 0#
  For i = 0 To NumberOfElement - 1
    For j = 0 To 3
      P6QuadratureArea = P6QuadratureArea + Elem(i).dj(j)
    Next j
  Next i
  If tangentSignature <> P6TangentSignature Then
    P6TangentGeneration = P6TangentGeneration + 1
    If P6TangentGeneration <= 0 Then P6TangentGeneration = 1
    P6TangentSignature = tangentSignature
  End If
  SetElmMat = True
End Function

Private Function P6ComputeTangentSignature() As String
  ' 入力指紋そのものを接線キーにする。世代カウンタは診断用。
  P6ComputeTangentSignature = Format$(P6InputTopologyFingerprint, "0") & "/" & Format$(P6InputGeometryFingerprint, "0") & "/" & Format$(P6InputMaterialFingerprint, "0") & "/hg" & Format$(P6HourglassFactor, "0.0000")
End Function

Public Function P6ComputeGeometrySignature() As String
  P6ComputeGeometrySignature = CStr(P6GeometryRevision)
End Function

Public Sub P6InvalidateStrengthGeneration()
  P6StrengthGeneration = P6StrengthGeneration + 1
  If P6StrengthGeneration <= 0 Then P6StrengthGeneration = 1
  P6FactorReady = False
  P6FactorNodeCondN = -1
  P6CSRReady = False
  P6CSRNumericGeneration = -1
  P6CSRMaterialGeneration = -1
  P6CSRBoundaryApplied = False
  P6CSRDiagonalInverseReady = False
  P3TrialStateValid = False
  P3ForceTangentRebuild = True
End Sub

Public Sub P6InvalidateTangentGeneration()
  P6MaterialGeneration = P6MaterialGeneration + 1
  If P6MaterialGeneration <= 0 Then P6MaterialGeneration = 1
  P6MaterialSignature = CStr(P6MaterialInputRevision)
  P6TangentGeneration = P6TangentGeneration + 1
  If P6TangentGeneration <= 0 Then P6TangentGeneration = 1
  P6TangentSignature = vbNullString
  P6ElasticKCacheReady = False
  P6ElasticOperatorCacheReady = False
  P6ElasticWeightCacheReady = False
  Erase P6ElasticElementWeightCache
  P6ElasticKCacheStatus = "invalidated_tangent"
  P6HgShapeGeneration = P6HgShapeGeneration + 1
  If P6HgShapeGeneration <= 0 Then P6HgShapeGeneration = 1
  P6InvalidateSpeedCaches
  P3TrialStateValid = False
  P3MarkAllTangentDirty
  P6CSRMaterialGeneration = -1
  P6CSRNumericGeneration = -1
  P6CSRReady = False
  P6CSRBoundaryApplied = False
  P6CSRDiagonalInverseReady = False
  P6CSRScatterCount = 0
  P6CSRScatterElementCount = -1
  P6CSRScatterTangentGeneration = -1
  P6CSRScatterMaterialGeneration = -1
  P6CSRScatterPatternSignature = vbNullString
    Erase P6CSRScatterStart
    Erase P6CSRScatterLocalRow
    Erase P6CSRScatterLocalColumn
    Erase P6CSRScatterPosition
    Erase P6CSRScatterValueIndex
End Sub
Sub SaveTotalMat()
Dim i As Long, j As Long
'境界条件設定後の全体剛性行列の保存
  With ThisWorkbook.Worksheets("全体剛性行列2")
    For i = 0 To lastDof
       For j = 0 To BandWidth
       If P6BandArrayReady(P6FactoredBand) Then
         .Cells(i + 2, j + 2) = P6FactoredBand(i, j)
       ElseIf P6BandArrayReady(OriginalMat) Then
         .Cells(i + 2, j + 2) = OriginalMat(i, j)
       End If
       Next j
     Next i
  End With
End Sub

Sub SaveStress()
'応力の保存。先頭列はステージ番号。ALLなら既存の下へ追記。
  Dim ws As Worksheet
  Dim rows() As Variant
  Dim elementId As Long, gaussId As Long, resultId As Long
  Dim startRow As Long, stageNo As Long, outputMode As String, lastUsed As Long
  Const stressItemCount As Long = 17
  Const lastDataColumn As Long = 4 + 4 * stressItemCount
  FEMPracticalSaveJointResults
  Set ws = ThisWorkbook.Worksheets("結果要素")
  P6WriteElementResultHeaders ws
  If NumberOfElement <= 0 Then Exit Sub
  stageNo = P3ResultStageNumber()
  outputMode = UCase$(Trim$(P6ReadTextSetting("OUTPUT_STAGE_MODE", "FINAL")))
  If outputMode = "ALL" Then
    startRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If startRow < 2 Then startRow = 2 Else startRow = startRow + 1
  Else
    startRow = 2
    lastUsed = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If P6LastStressOutRow > lastUsed Then lastUsed = P6LastStressOutRow
    If lastUsed >= 2 Then
      ws.range(ws.Cells(2, 1), ws.Cells(lastUsed, lastDataColumn)).ClearContents
    End If
  End If
  Dim skipInactive As Boolean, writeCount As Long
  skipInactive = (UCase$(Trim$(P6ReadTextSetting("OUTPUT_INACTIVE_ELEMENTS", "SKIP"))) <> "INCLUDE")
  ReDim rows(1 To NumberOfElement, 1 To lastDataColumn)
  writeCount = 0
  For elementId = 0 To NumberOfElement - 1
    If skipInactive And Not P3IsElementActive(elementId) Then GoTo NextStressElement
    If Elem(elementId).IsJoint Then
      writeCount = writeCount + 1
      rows(writeCount, 1) = stageNo
      rows(writeCount, 2) = elementId + 1
      rows(writeCount, 3) = Elem(elementId).x(2)
      rows(writeCount, 4) = Elem(elementId).y(2)
      If P3IsElementActive(elementId) Then rows(writeCount, 5) = "接合結果参照"
      GoTo NextStressElement
    End If
    CalcStrain elementId, Elem(elementId).Bmat, TDisp, Elem(elementId).Stmat
    CalcElePoint elementId, Elem(elementId).x, Elem(elementId).y, Elem(elementId).p
    writeCount = writeCount + 1
    rows(writeCount, 1) = stageNo
    rows(writeCount, 2) = elementId + 1
    rows(writeCount, 3) = Elem(elementId).p(0)
    rows(writeCount, 4) = Elem(elementId).p(1)
    For gaussId = 0 To 3
      For resultId = 0 To stressItemCount - 1
        rows(writeCount, 5 + resultId + stressItemCount * gaussId) = Elem(elementId).Stmat(resultId, gaussId)
      Next resultId
    Next gaussId
NextStressElement:
  Next elementId
  If writeCount > 0 Then
    ws.range(ws.Cells(startRow, 1), ws.Cells(startRow + writeCount - 1, lastDataColumn)).value2 = rows
    P6LastStressOutRow = startRow + writeCount - 1
  End If
End Sub

Private Sub P6WriteElementResultHeaders(ByVal ws As Worksheet)
  Dim gaussId As Long, resultId As Long, columnId As Long
  ws.Cells(1, 1).value2 = "ステージ番号"
  ws.Cells(1, 2).value2 = "要素番号"
  ws.Cells(1, 3).value2 = "要素X"
  ws.Cells(1, 4).value2 = "要素Y"
  For gaussId = 0 To 3
    For resultId = 0 To 16
      columnId = 5 + resultId + 17 * gaussId
      ws.Cells(1, columnId).value2 = "G" & CStr(gaussId + 1) & " " & P6StressResultName(resultId)
    Next resultId
  Next gaussId
  On Error Resume Next
  If ws.AutoFilterMode Then ws.AutoFilterMode = False
  ws.range(ws.Cells(1, 1), ws.Cells(1, 72)).AutoFilter
  On Error GoTo 0
End Sub

Private Function P6StressResultName(ByVal resultId As Long) As String
  Select Case resultId
    Case 0: P6StressResultName = "X応力"
    Case 1: P6StressResultName = "Y応力"
    Case 2: P6StressResultName = "XY応力"
    Case 3: P6StressResultName = "最大主応力"
    Case 4: P6StressResultName = "最小主応力"
    Case 5: P6StressResultName = "最大せん断応力"
    Case 6: P6StressResultName = "主応力角"
    Case 7: P6StressResultName = "相当応力"
    Case 8: P6StressResultName = "偏差応力"
    Case 9: P6StressResultName = "降伏関数"
    Case 10: P6StressResultName = "Xひずみ"
    Case 11: P6StressResultName = "Yひずみ"
    Case 12: P6StressResultName = "XYせん断ひずみ"
    Case 13: P6StressResultName = "8面せん断ひずみ"
    Case 14: P6StressResultName = "体積ひずみ"
    Case 15: P6StressResultName = "最大せん断ひずみ"
    Case 16: P6StressResultName = "Z応力"
    Case Else: P6StressResultName = "項目" & CStr(resultId)
  End Select
End Function
Sub GenNode(k, xx, yy)
'節点データ書込
  With ThisWorkbook.Worksheets("節点データ")
     k = k + 1
     .Cells(k + 1, 1) = k
     .Cells(k + 1, 2) = xx
     .Cells(k + 1, 3) = yy
  End With
End Sub
Sub GenElm(k, n1, n2, n3, N4, mat)
  Err.Raise vbObjectError + 3480, "FEM.GenElm", "Q8統一後は4節点要素を書き込めません。"
End Sub
Sub SetStress()
Dim ii As Long
  Dim i As Long, j As Long, k As Long
  For i = 0 To NumberOfElement - 1
     If Elem(i).IsJoint Then GoTo NextSetStress
     With Elem(i)
        ii = .MatNo
        For j = 0 To 15 '
           k = .ElNode(j)
           If k >= 0 Then .u(j) = TDisp(k)
        Next j
        CalcStress .Smat, .Stmat, .u, Material(ii).fai, Material(ii).cohesion, Material(ii).psai, Material(ii).Poisson
     End With
NextSetStress:
  Next i
End Sub
Sub UpdateDisp()
  Dim i As Long
  For i = 0 To lastDof
    TDisp(i) = Disp(i)
    UDisp(i) = Disp(i)
  Next i
End Sub

Sub SetForce()
Dim ii As Long: Dim kk As Long
  Dim i As Long, j As Long
  For i = 0 To NumberOfElement - 1
     If Elem(i).IsJoint Then
       P3JointEval i, True
       GoTo NextSetForce
     End If
     With Elem(i)
        ii = .MatNo
        .ElNo = i
        For j = 0 To 15 '15
           kk = .ElNode(j)
           If kk >= 0 Then .u(j) = Disp(kk)
        Next j
        CalcStress .Smat, .Stmat, .u, Material(ii).fai, Material(ii).cohesion, Material(ii).psai, Material(ii).Poisson
        CalcForce .ElNo, .Bmat, .dj, .Stmat
     End With
NextSetForce:
  Next i
End Sub
Sub CalcStress(Smat() As Double, Stmat() As Double, UE() As Double, fai As Double, cohesion As Double, psai As Double, Poisson As Double)
  Dim im As Long, i As Long, j As Long
  Dim s As Double, centerValue As Double, radiusValue As Double
  Dim planeSigma1 As Double, planeSigma2 As Double, sigmaZ As Double
  Dim sigmaMax As Double, sigmaMid As Double, sigmaMin As Double
  Dim frictionSin As Double, frictionCos As Double
  Dim principalAngle As Double, piValue As Double
  piValue = 3.14159265358979
  frictionSin = Sin(fai * piValue / 180#)
  frictionCos = Cos(fai * piValue / 180#)
  For im = 0 To 3
    For i = 0 To 2
      s = 0#
      For j = 0 To 15
        s = s + Smat(i, j, im) * UE(j)
      Next j
      Stmat(i, im) = s
    Next i
    If P1Formulation <> P1_FORMULATION_PLANE_STRAIN Then
      P1Formulation = P1_FORMULATION_PLANE_STRAIN
    End If
    sigmaZ = Poisson * (Stmat(0, im) + Stmat(1, im))
    Stmat(16, im) = sigmaZ
    centerValue = 0.5 * (Stmat(0, im) + Stmat(1, im))
    radiusValue = Sqr((0.5 * (Stmat(0, im) - Stmat(1, im))) ^ 2 + Stmat(2, im) ^ 2)
    planeSigma1 = centerValue + radiusValue
    planeSigma2 = centerValue - radiusValue
    principalAngle = 0.5 * P1Atan2(2# * Stmat(2, im), Stmat(0, im) - Stmat(1, im))
    P1SortPrincipal3 planeSigma1, planeSigma2, sigmaZ, sigmaMax, sigmaMid, sigmaMin
    Stmat(3, im) = sigmaMax
    Stmat(4, im) = sigmaMin
    Stmat(5, im) = 0.5 * (sigmaMax - sigmaMin)
    Stmat(6, im) = principalAngle * 180# / piValue
    Stmat(7, im) = P1VonMises3D(Stmat(0, im), Stmat(1, im), sigmaZ, Stmat(2, im))
    Stmat(8, im) = Stmat(5, im)
    Stmat(9, im) = sigmaMax - sigmaMin - 2# * cohesion * frictionCos - (sigmaMax + sigmaMin) * frictionSin
  Next im
End Sub
Sub SaveDmat()
'全体剛性行列の保存
  Dim ws1 As Worksheet
  Dim i As Long, j As Long, k As Long
  Set ws1 = ThisWorkbook.Worksheets("Sheet1")
  With ws1
    For i = 0 To 0 'NumberOfElement
      With Elem(i)
      For j = 0 To 15
        For k = 0 To 15
          ws1.Cells(j + 2, k + 2) = .kmat(j, k)
        Next k
      Next j
      End With
    Next i
  End With
End Sub
Sub WeightUpdate()
Dim nfreedom As Long
Dim i As Long, j As Long, k As Long
Dim a As Long, b As Long, materialIndex As Long
Dim thickness As Double, gammaValue As Double, detJ As Double, dVolume As Double
Dim dXdxi As Double, dYdxi As Double, dXdeta As Double, dYdeta As Double
Dim xiValue As Double, etaValue As Double, shapeValue As Double
Dim gaussXi(0 To 1) As Double, gaussW(0 To 1) As Double
Dim shapeN(0 To 7) As Double, dNxi(0 To 7) As Double, dNeta(0 To 7) As Double
  P3EnsurePolicyCache
  If P3WeightMode <> 0 Then
    Err.Raise vbObjectError + 3481, "FEM.WeightUpdate", "WEIGHT_MODEはTOTALのみです。浮力・水中重量は未実装。"
  End If
  P6WeightMode = "TOTAL"
  P6WeightNote = "全応力γ。Q8 2×2（剛性・内力と同じ積分）。水位・浮力なし"
  gaussXi(0) = -0.5773502692: gaussXi(1) = 0.5773502692
  gaussW(0) = 1#: gaussW(1) = 1#
  TotalSelfWeight = 0#
  ReDim P3AppliedForce(lastDof)
  ReDim OriginalForce(lastDof)
  For i = 0 To lastDof
    P3AppliedForce(i) = Force(i)
  Next i
  For i = 0 To NumberOfElement - 1
    If Not P3IsElementActive(i) Then GoTo NextWeightElement
    If Elem(i).IsJoint Then GoTo NextWeightElement
    materialIndex = Elem(i).MatNo
    If materialIndex < 0 Or materialIndex > NumberOfMaterial - 1 Then GoTo NextWeightElement
    thickness = Material(materialIndex).thickness
    gammaValue = Material(materialIndex).weight
    For a = 0 To 1
      For b = 0 To 1
        xiValue = gaussXi(a)
        etaValue = gaussXi(b)
        P6Q8Shape xiValue, etaValue, shapeN, dNxi, dNeta
        dXdxi = 0#: dYdxi = 0#: dXdeta = 0#: dYdeta = 0#
        For j = 0 To 7
          dXdxi = dXdxi + dNxi(j) * Elem(i).x(j)
          dYdxi = dYdxi + dNxi(j) * Elem(i).y(j)
          dXdeta = dXdeta + dNeta(j) * Elem(i).x(j)
          dYdeta = dYdeta + dNeta(j) * Elem(i).y(j)
        Next j
        detJ = dXdxi * dYdeta - dYdxi * dXdeta
        dVolume = detJ * gaussW(a) * gaussW(b) * thickness
        TotalSelfWeight = TotalSelfWeight + gammaValue * dVolume
        For j = 0 To 7
          shapeValue = gammaValue * shapeN(j) * dVolume
          k = 2 * Elem(i).node(j)
          If k >= 0 And k < nDof Then
            Force(k) = Force(k) + shapeValue * nh
            Force(k + 1) = Force(k + 1) + shapeValue * nv
          End If
        Next j
      Next b
    Next a
NextWeightElement:
  Next i
  ReDim P3SelfWeightForce(lastDof)
  TotalAppliedLoadX = 0#
  TotalAppliedLoadY = 0#
  For i = 0 To lastDof
    NForce(i) = Force(i)
    OriginalForce(i) = Force(i)
    P3SelfWeightForce(i) = Force(i) - P3AppliedForce(i)
    If i Mod 2 = 0 Then
      TotalAppliedLoadX = TotalAppliedLoadX + Force(i)
    Else
      TotalAppliedLoadY = TotalAppliedLoadY + Force(i)
    End If
  Next i
End Sub

Public Sub P6Q8Shape(ByVal xi As Double, ByVal eta As Double, ByRef shapeN() As Double, ByRef dNxi() As Double, ByRef dNeta() As Double)
  shapeN(0) = 0.25 * (1# - xi) * (1# - eta) * (-1# - xi - eta)
  shapeN(1) = 0.25 * (1# + xi) * (1# - eta) * (-1# + xi - eta)
  shapeN(2) = 0.25 * (1# + xi) * (1# + eta) * (-1# + xi + eta)
  shapeN(3) = 0.25 * (1# - xi) * (1# + eta) * (-1# - xi + eta)
  shapeN(4) = 0.5 * (1# - xi * xi) * (1# - eta)
  shapeN(5) = 0.5 * (1# + xi) * (1# - eta * eta)
  shapeN(6) = 0.5 * (1# - xi * xi) * (1# + eta)
  shapeN(7) = 0.5 * (1# - xi) * (1# - eta * eta)
  dNxi(0) = 0.25 * (1# - eta) * (2# * xi + eta)
  dNxi(1) = 0.25 * (1# - eta) * (2# * xi - eta)
  dNxi(2) = 0.25 * (1# + eta) * (2# * xi + eta)
  dNxi(3) = 0.25 * (1# + eta) * (2# * xi - eta)
  dNxi(4) = -xi * (1# - eta)
  dNxi(5) = 0.5 * (1# - eta * eta)
  dNxi(6) = -xi * (1# + eta)
  dNxi(7) = -0.5 * (1# - eta * eta)
  dNeta(0) = 0.25 * (1# - xi) * (2# * eta + xi)
  dNeta(1) = 0.25 * (1# + xi) * (2# * eta - xi)
  dNeta(2) = 0.25 * (1# + xi) * (2# * eta + xi)
  dNeta(3) = 0.25 * (1# - xi) * (2# * eta - xi)
  dNeta(4) = -0.5 * (1# - xi * xi)
  dNeta(5) = -eta * (1# + xi)
  dNeta(6) = 0.5 * (1# - xi * xi)
  dNeta(7) = -eta * (1# - xi)
End Sub

Sub SaveData0()
Dim ii As Long: Dim nfreedom As Long
Dim i As Long, rows() As Variant
  nfreedom = NumberOfFreeNode * 2 - 1
  If NumberOfFreeNode <= 0 Then Exit Sub
  ReDim rows(1 To NumberOfFreeNode * 2, 1 To 1)
  For ii = 0 To NumberOfFreeNode - 1
    rows(ii + 1, 1) = NForce(ii * 2)
    rows(ii + NumberOfFreeNode + 1, 1) = NForce(ii * 2 + 1)
  Next ii
  ThisWorkbook.Worksheets("力").range(ThisWorkbook.Worksheets("力").Cells(2, 2), ThisWorkbook.Worksheets("力").Cells(NumberOfFreeNode * 2 + 1, 2)).value2 = rows
End Sub
Sub SaveData1()
Dim ii As Long: Dim nfreedom As Long
Dim i As Long, rows() As Variant
  nfreedom = NumberOfFreeNode * 2 - 1
  If NumberOfFreeNode <= 0 Then Exit Sub
  ReDim rows(1 To NumberOfFreeNode * 2, 1 To 1)
  For ii = 0 To NumberOfFreeNode - 1
    rows(ii + 1, 1) = iNForce(ii * 2)
    rows(ii + NumberOfFreeNode + 1, 1) = iNForce(ii * 2 + 1)
  Next ii
  ThisWorkbook.Worksheets("力").range(ThisWorkbook.Worksheets("力").Cells(2, 2 + nmn), ThisWorkbook.Worksheets("力").Cells(NumberOfFreeNode * 2 + 1, 2 + nmn)).value2 = rows
End Sub
Sub SaveData2()
Dim ii As Long: Dim nfreedom As Long
Dim i As Long, rows() As Variant
  nfreedom = NumberOfFreeNode * 2 - 1
  If NumberOfFreeNode <= 0 Then Exit Sub
  ReDim rows(1 To NumberOfFreeNode * 2, 1 To 1)
  For ii = 0 To NumberOfFreeNode - 1
    rows(ii + 1, 1) = TDisp(ii * 2)
    rows(ii + NumberOfFreeNode + 1, 1) = TDisp(ii * 2 + 1)
  Next ii
  ThisWorkbook.Worksheets("変位").range(ThisWorkbook.Worksheets("変位").Cells(2, 2 + nmn), ThisWorkbook.Worksheets("変位").Cells(NumberOfFreeNode * 2 + 1, 2 + nmn)).value2 = rows
End Sub

Sub SaveDisp()
Dim ii As Long, nfreedom As Long
Dim i As Long, outputNode As Long, internalNode As Long
Dim rows() As Variant, ws As Worksheet
Dim startRow As Long, stageNo As Long, outputMode As String, lastUsed As Long
  nfreedom = NumberOfFreeNode * 2 - 1
  If NumberOfFreeNode <= 0 Then Exit Sub
  stageNo = P3ResultStageNumber()
  outputMode = UCase$(Trim$(P6ReadTextSetting("OUTPUT_STAGE_MODE", "FINAL")))
  Set ws = ThisWorkbook.Worksheets("結果節点")
  ws.Cells(1, 1).value2 = "ステージ番号"
  ws.Cells(1, 2).value2 = "節点番号"
  ws.Cells(1, 3).value2 = "X"
  ws.Cells(1, 4).value2 = "Y"
  ws.Cells(1, 5).value2 = "X変位"
  ws.Cells(1, 6).value2 = "Y変位"
  ws.Cells(1, 7).value2 = "X力"
  ws.Cells(1, 8).value2 = "Y力"
  ws.Cells(1, 9).value2 = "反力X"
  ws.Cells(1, 10).value2 = "反力Y"
  If outputMode = "ALL" Then
    startRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If startRow < 2 Then startRow = 2 Else startRow = startRow + 1
  Else
    startRow = 2
    lastUsed = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If P6LastDispOutRow > lastUsed Then lastUsed = P6LastDispOutRow
    If lastUsed >= 2 Then
      ws.range(ws.Cells(2, 1), ws.Cells(lastUsed, 10)).ClearContents
    End If
  End If
  ReDim rows(1 To NumberOfFreeNode, 1 To 10)
  For outputNode = 0 To NumberOfFreeNode - 1
    internalNode = P6GetInternalFreeNode(outputNode)
    ii = outputNode
    i = 2 * internalNode
    rows(ii + 1, 1) = stageNo
    rows(ii + 1, 2) = outputNode + 1
    rows(ii + 1, 3) = XXX(i)
    rows(ii + 1, 4) = XXX(i + 1)
    rows(ii + 1, 5) = TDisp(i)
    rows(ii + 1, 6) = TDisp(i + 1)
    rows(ii + 1, 7) = OriginalForce(i)
    rows(ii + 1, 8) = OriginalForce(i + 1)
    rows(ii + 1, 9) = Reaction(i)
    rows(ii + 1, 10) = Reaction(i + 1)
  Next outputNode
  ws.range(ws.Cells(startRow, 1), ws.Cells(startRow + NumberOfFreeNode - 1, 10)).value2 = rows
  P6LastDispOutRow = startRow + NumberOfFreeNode - 1
  On Error Resume Next
  If ws.AutoFilterMode Then ws.AutoFilterMode = False
  ws.range(ws.Cells(1, 1), ws.Cells(1, 10)).AutoFilter
  On Error GoTo 0
End Sub

Public Function Nrf(Smat() As Double, Spmat() As Double, UE() As Double, Stmat() As Double, mStmat() As Double, fai As Double, cohesion As Double, psai As Double, young As Double, Poisson As Double, Dmat() As Double, Dpmat() As Double, Bmat() As Double, Optional ByVal materialIndex As Long = -1) As Boolean
  Dim inputState As P2_MaterialPointInput, outputState As P2_MaterialPointOutput
  Dim im As Long, i As Long, j As Long, k As Long
  Dim workStmat(0 To 16, 0 To 3) As Double
  Dim workSpmat(0 To 2, 0 To 15, 0 To 3) As Double
  Dim workDpmat(0 To 2, 0 To 2, 0 To 3) As Double
  Dim s As Double, avg As Double
  Dim errNumber As Long, errDescription As String
  Nrf = False
  On Error GoTo ErrLabel
  If P3CurrentElementZeroIncrement And CurrentIncrement >= 0 Then
    ' 増分開始時は試行ひずみがゼロなので、確定応力を再利用する。
    ' 構成則の降伏判定・リターンマップは非ゼロ増分時に必ず実行する。
    For im = 0 To 3
      For i = 0 To 16
        Stmat(i, im) = mStmat(i, im)
      Next i
    Next im
    Nrf = True
    Exit Function
  End If
  inputState.young = young
  inputState.Poisson = Poisson
  inputState.frictionAngle = fai
  inputState.cohesion = cohesion
  inputState.dilationAngle = psai
  inputState.tolerance = P2_DEFAULT_TOLERANCE
  inputState.maxIterations = P2_DEFAULT_MAX_ITERATIONS
  inputState.maxSubsteps = P2_DEFAULT_MAX_SUBSTEPS
  inputState.EnableSubstepping = True
  If materialIndex >= 0 And materialIndex < NumberOfMaterial Then
    If Material(materialIndex).MaterialCacheReady Then
      inputState.UseCachedConstants = True
      inputState.CachedElasticD00 = Material(materialIndex).ElasticD00
      inputState.CachedElasticD01 = Material(materialIndex).ElasticD01
      inputState.CachedElasticD22 = Material(materialIndex).ElasticD22
      inputState.CachedSinFriction = Material(materialIndex).SinFriction
      inputState.CachedCosFriction = Material(materialIndex).CosFriction
      inputState.CachedSinDilation = Material(materialIndex).SinDilation
      inputState.CachedCosDilation = Material(materialIndex).CosDilation
    End If
  End If
  For im = 0 To 3
    For i = 0 To 16
      workStmat(i, im) = Stmat(i, im)
    Next i
    inputState.PreviousStress(0) = mStmat(0, im)
    inputState.PreviousStress(1) = mStmat(1, im)
    inputState.PreviousStress(2) = mStmat(2, im)
    inputState.PreviousStress(3) = mStmat(16, im)
    For i = 0 To 2
      inputState.StrainIncrement(i) = 0#
      For j = 0 To 15
        inputState.StrainIncrement(i) = inputState.StrainIncrement(i) + Bmat(i, j, im) * UE(j)
      Next j
    Next i
    inputState.PreferredSubsteps = 1
    If P3TrialStateInitialized And FailureElement >= 0 And FailureElement < NumberOfElement Then
      If Elem(FailureElement).P2LastSuccessfulSubsteps(im) > 1 Then
        inputState.PreferredSubsteps = Elem(FailureElement).P2LastSuccessfulSubsteps(im)
      End If
    End If
    P2ResetOutput outputState
    If Not P2MaterialPointUpdate(inputState, outputState) Then
      If outputState.failureCode = 0 Then
        errNumber = vbObjectError + P2_FAILURE_NOT_CONVERGED
      Else
        errNumber = vbObjectError + outputState.failureCode
      End If
      errDescription = outputState.failureMessage
      If Len(errDescription) = 0 Then errDescription = "平面ひずみ材料点の塑性更新に失敗しました。"
      SetAnalysisFailure RESULT_MATERIAL_ERROR, errDescription, errNumber, FailureElement, im, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    ' Preserve the complete algorithmic tangent from the constitutive update.
    workStmat(0, im) = outputState.Stress(0)
    workStmat(1, im) = outputState.Stress(1)
    workStmat(2, im) = outputState.Stress(2)
    workStmat(3, im) = outputState.PrincipalStress(0)
    workStmat(4, im) = outputState.PrincipalStress(2)
    workStmat(5, im) = 0.5 * (outputState.PrincipalStress(0) - outputState.PrincipalStress(2))
    workStmat(6, im) = outputState.principalAngle
    workStmat(7, im) = P2VonMises3D(outputState.Stress(0), outputState.Stress(1), outputState.Stress(3), outputState.Stress(2))
    If outputState.PrincipalStress(2) > 0# Then
      workStmat(8, im) = outputState.PrincipalStress(0)
    ElseIf outputState.PrincipalStress(0) > 0# Then
      workStmat(8, im) = outputState.PrincipalStress(0) - outputState.PrincipalStress(2)
    Else
      workStmat(8, im) = -outputState.PrincipalStress(2)
    End If
    workStmat(9, im) = outputState.YieldFunction
    workStmat(16, im) = outputState.Stress(3)
    If P3TrialStateInitialized And FailureElement >= 0 And FailureElement < NumberOfElement Then
      For i = 0 To 3
        Elem(FailureElement).P2PlasticStrain(i, im) = outputState.PlasticStrain(i)
      Next i
      If outputState.yielded Xor Elem(FailureElement).P2Yielded(im) Then
        P3PlasticStateChangeCount = P3PlasticStateChangeCount + 1
      End If
      If outputState.yielded Or Elem(FailureElement).P2Yielded(im) Then
        Elem(FailureElement).TangentDirty = True
      End If
      If outputState.PlasticOccurred Then P3PlasticPointUpdateCount = P3PlasticPointUpdateCount + 1
      If outputState.yielded And Not P3ActivePlasticPoint(im, FailureElement) Then P3ActivePlasticPointCount = P3ActivePlasticPointCount + 1
      If Not outputState.yielded And P3ActivePlasticPoint(im, FailureElement) Then P3ActivePlasticPointCount = P3ActivePlasticPointCount - 1
      P3ActivePlasticPoint(im, FailureElement) = outputState.yielded
      Elem(FailureElement).P2PlasticMultiplier(im) = outputState.PlasticMultiplier
      Elem(FailureElement).P2YieldFunction(im) = outputState.YieldFunction
      Elem(FailureElement).P2Yielded(im) = outputState.yielded
      For i = 0 To 2
        Elem(FailureElement).P2PrincipalStress(i, im) = outputState.PrincipalStress(i)
      Next i
      If outputState.Elastic Then
        P3DecaySubstepHint FailureElement, im
      ElseIf outputState.PlasticOccurred Or outputState.UsedSubsteps > 1 Then
        P3RememberSubstepHint FailureElement, im, outputState.UsedSubsteps
      End If
    End If
    For i = 0 To 2
      For j = 0 To 2
        workDpmat(i, j, im) = outputState.tangent(i, j)
      Next j
      For j = 0 To 15
        s = 0#
        For k = 0 To 2
          s = s + outputState.tangent(i, k) * Bmat(k, j, im)
        Next k
        workSpmat(i, j, im) = s
      Next j
    Next i
  Next im
  For im = 0 To 3
    For i = 0 To 16
      Stmat(i, im) = workStmat(i, im)
    Next i
    For i = 0 To 2
      For j = 0 To 15
        Spmat(i, j, im) = workSpmat(i, j, im)
      Next j
    Next i
    If FailureElement >= 0 And FailureElement <= NumberOfElement - 1 Then
      Elem(FailureElement).SpmatValid(im) = True
    End If
  Next im
  For im = 0 To 3
    For i = 0 To 2
      For j = 0 To 2
        Dpmat(i, j, im) = workDpmat(i, j, im)
      Next j
    Next i
  Next im
  Nrf = True
  Exit Function
ErrLabel:
  errNumber = Err.Number
  errDescription = Err.Description
  If errNumber = 18 Then
    SetAnalysisFailure RESULT_NONCONVERGED, "ユーザーが解析を中断しました。", 18, FailureElement, -1, CurrentIncrement, CurrentIteration
    Nrf = False
    Exit Function
  End If
  If errNumber = 0 Then errNumber = vbObjectError + P2_FAILURE_NONFINITE
  If Len(errDescription) = 0 Then errDescription = "平面ひずみ材料点更新で不明なエラーが発生しました。"
  SetAnalysisFailure RESULT_MATERIAL_ERROR, errDescription, errNumber, FailureElement, -1, CurrentIncrement, CurrentIteration
  Nrf = False
End Function

Private Function NrfLegacyP0(Smat() As Double, Spmat() As Double, UE() As Double, Stmat() As Double, mStmat() As Double, fai As Double, cohesion As Double, psai As Double, Dmat() As Double, Dpmat() As Double, Bmat() As Double) As Boolean
Dim ai As Variant: Dim beta As Variant: Dim cf1 As Variant: Dim cf2 As Variant: Dim cr As Variant
Dim fa As Variant: Dim fa1 As Variant: Dim fa2 As Variant: Dim im As Variant: Dim msg As Variant
Dim rs As Variant: Dim sf1 As Variant: Dim sf2 As Variant: Dim txy As Variant: Dim wk As Variant
Dim i As Long, j As Long, k As Long
Dim s As Double


  Dim Dft(2, 2) As Double
  Dim CC As Double
  Dim SS As Double
  Dim AA As Double
  Dim d1 As Double
  Dim d2 As Double
  Dim Si1 As Double
  Dim Si3 As Double
  Dim TXX As Double
  Dim Cf As Double
  Dim Sf As Double
  Dim dStmat(2, 3) As Double
  Dim betaDenom As Double
  Dim errNumber As Long, errDescription As String
  NrfLegacyP0 = False
  On Error GoTo ErrLabel
  sf1 = Sin(psai * 3.1415926535 / 180)
  cf1 = Cos(psai * 3.1415926535 / 180)
  
  sf2 = Sin(fai * 3.1415926535 / 180)
  cf2 = Cos(fai * 3.1415926535 / 180)
  
  Sf = sf2
  Cf = cf2
 
  For im = 0 To 3 'ガウス点
    '.Smat（弾性仮定）変位増分の時の増分応力値計算
    '.mSmat 前回応力値計算

    
   '今の応力状態
    CC = (mStmat(0, im) + mStmat(1, im)) / 2
    SS = (mStmat(0, im) - mStmat(1, im)) / 2
    txy = mStmat(2, im)
    Si1 = CC + Sqr(SS * SS + txy * txy)
    Si3 = CC - Sqr(SS * SS + txy * txy)
    fa1 = Si1 - Si3 - (Si1 + Si3) * Sf - 2# * cohesion * Cf 'f(σ)の降伏関数
    
    d1 = Dmat(0, 0)
    d2 = Dmat(0, 1)
    Sf = sf1
    Cf = cf1
    Dft(0, 0) = (d1 - d2 - (d1 + d2) * Sf) * (d1 - d2 - (d1 + d2) * Sf)
    Dft(1, 0) = (d1 + d2) * (d1 + d2) * Sf * Sf - (d1 - d2) * (d1 - d2)
    Dft(0, 1) = Dft(1, 0)
    Dft(1, 1) = (d2 - d1 - (d1 + d2) * Sf) * (d2 - d1 - (d1 + d2) * Sf)
    ai = 2 * d1 * (1 + Sf * Sf) - 2 * d2 * (1 - Sf * Sf)
  
    If Abs(ai) <= 0.000000000001 Then
      Err.Raise vbObjectError + 3101, "FEM.Nrf", "塑性接線の分母が0または小さすぎます。"
    End If
    For i = 0 To 2
      For j = 0 To 2
        Dpmat(i, j) = Dmat(i, j) - Dft(i, j) / ai
      Next j
    Next i

    Sf = sf2
    Cf = cf2
  
'----------------------------------------------------------
'変位増分を考慮した応力増分　ΔUに対するΔσ(弾性)
    For i = 0 To 2
      dStmat(i, im) = 0#
      For j = 0 To 15
        dStmat(i, im) = dStmat(i, im) + Smat(i, j, im) * UE(j) 'dDisp
      Next j
      Stmat(i, im) = mStmat(i, im) + dStmat(i, im)  'σ+Δσ
    Next i
    'σ+Δσ(弾性)での応力状態計算
    CC = (Stmat(0, im) + Stmat(1, im)) / 2
    SS = (Stmat(0, im) - Stmat(1, im)) / 2
    txy = Stmat(2, im)
    Si1 = CC + Sqr(SS * SS + txy * txy)
    Si3 = CC - Sqr(SS * SS + txy * txy)
    fa2 = Si1 - Si3 - (Si1 + Si3) * Sf - 2# * cohesion * Cf  'f(σ+Δσ)の降伏関数
    fa = fa2 - fa1
    If fa2 > 0# Then 'f(σ+Δσ)塑性
      If Abs(fa1) < 0.0001 Or Abs(fa) < 0.0001 Then
        rs = 0#  '降伏関数上
      Else
        rs = -fa1 / fa
      End If
    Else
      rs = 1# '弾性
    End If
    rs = Abs(rs)
    '要素塑性マトリックス作成
    For i = 0 To 2
      For j = 0 To 2
        Dpmat(i, j) = Dmat(i, j) - (1 - rs) * Dpmat(i, j)
      Next j
    Next i
    For i = 0 To 2
      For j = 0 To 15
        s = 0#
        For k = 0 To 2
          s = s + Dpmat(i, k) * Bmat(k, j, im) 'D行列とB行列の積　［Sp］=［Dep］［B］
        Next k
        Spmat(i, j, im) = s
      Next j
    Next i
  
    For i = 0 To 2
      s = 0#
      For j = 0 To 15
        s = s + Spmat(i, j, im) * UE(j)
      Next j
      Stmat(i, im) = mStmat(i, im) + s
    Next i
  
  '応力補正して降伏関数上にのせる
  
    CC = (Stmat(0, im) + Stmat(1, im)) / 2
    SS = (Stmat(0, im) - Stmat(1, im)) / 2
    txy = Stmat(2, im)
    Si1 = CC + Sqr(SS * SS + txy * txy)
    Si3 = CC - Sqr(SS * SS + txy * txy)
    

      
    fa2 = Si1 - Si3 - (Si1 + Si3) * Sf - 2# * cohesion * Cf
    betaDenom = 2# * cohesion * Cf + (Si1 + Si3) * Sf
    If fa2 > 0# Then
      If Abs(betaDenom) <= 0.000000000001 Then
        Err.Raise vbObjectError + 3102, "FEM.Nrf", "降伏面補正の分母が0または小さすぎます。"
      End If
      beta = -1# + (Si1 - Si3) / betaDenom
    Else
      beta = 0#
    End If
      
    'beta = -1 + ((Si1 + Si3) * Sf + 2 * cohesion * Cf) / (Si1 - Si3)
    
    If fa2 > 0# Then
      Stmat(0, im) = CC + 1 / (1 + beta) * (Stmat(0, im) - CC)
      Stmat(1, im) = CC + 1 / (1 + beta) * (Stmat(1, im) - CC)
      Stmat(2, im) = 1 / (1 + beta) * Stmat(2, im)
    End If
    
    CC = (Stmat(0, im) + Stmat(1, im)) / 2
    SS = (Stmat(0, im) - Stmat(1, im)) / 2
    txy = Stmat(2, im)
    cr = Sqr(SS * SS + txy * txy)
    Si1 = CC + cr
    Si3 = CC - cr
    fa2 = Si1 - Si3 - (Si1 + Si3) * Sf - 2# * cohesion * Cf
    Stmat(3, im) = Si1
    Stmat(4, im) = Si3
    Stmat(5, im) = (Si1 - Si3) / 2
    If cr <= SS Then
      Stmat(6, im) = 0#
    Else
      wk = Abs((SS + cr) / (SS - cr))
      Stmat(6, im) = 57.29577951 * Atn(Sqr(wk))
    If Stmat(6, im) < 0# Then Stmat(6, im) = -Stmat(6, im)
      Stmat(6, im) = 90 - Stmat(6, im)
    End If
    Stmat(7, im) = Sqr(Stmat(3, im) * Stmat(3, im) - Stmat(3, im) * Stmat(4, im) + Stmat(4, im) * Stmat(4, im))
    If Stmat(4, im) > 0# Then
      Stmat(8, im) = Stmat(3, im)
    ElseIf Stmat(3, im) > 0# Then
      Stmat(8, im) = Stmat(3, im) - Stmat(4, im)
    Else
      Stmat(8, im) = -Stmat(4, im)
    End If
    Stmat(9, im) = fa2

  Next im

  
  
  
  


  


  NrfLegacyP0 = True
  Exit Function

ErrLabel:
  errNumber = Err.Number
  errDescription = Err.Description
  If errNumber = 0 Then errNumber = vbObjectError + 3100
  If Len(errDescription) = 0 Then errDescription = "材料点計算で不明なエラーが発生しました。"
  SetAnalysisFailure RESULT_MATERIAL_ERROR, errDescription, errNumber, FailureElement, im, CurrentIncrement, CurrentIteration
  NrfLegacyP0 = False
End Function
Sub CalcForce(ElMNo As Long, Bmat() As Double, dj() As Double, Stmat() As Double)
Dim im As Long: Dim jm As Long: Dim kj As Long
Dim j As Long, modeId As Long
Dim s As Double, projValue As Double
  For j = 0 To 15   'j　要素各点番号
    kj = Elem(ElMNo).ElNode(j)  'kj 力番号
    If kj < 0 Then GoTo NextCalcForceDof
    s = 0#
    For im = 0 To 3  'imガウス点　4個
      For jm = 0 To 2 '応力
        s = s + Bmat(jm, j, im) * Stmat(jm, im) * dj(im) * Material(Elem(ElMNo).MatNo).thickness '［B］T×J×t
      Next jm
    Next im
    iNForce(kj) = iNForce(kj) + s
NextCalcForceDof:
  Next j
  ' For INCONSISTENT SRM, hourglass stiffness is an iterative regularizer.
  ' Its unbounded elastic force must not add artificial soil bearing strength.
  ' Use the same material-only force during gravity replay and final recovery.
  If P3SrmEnabled And P3FlowPolicyIsInconsistent() Then Exit Sub
  If Elem(ElMNo).HourglassModeCount > 0 And Elem(ElMNo).HourglassScale > 0# Then
    For modeId = 0 To Elem(ElMNo).HourglassModeCount - 1
      projValue = 0#
      For j = 0 To 15
        If Elem(ElMNo).ElNode(j) >= 0 Then
          projValue = projValue + Elem(ElMNo).HourglassMode(modeId, j) * TDisp(Elem(ElMNo).ElNode(j))
        End If
      Next j
      projValue = Elem(ElMNo).HourglassScale * projValue
      For j = 0 To 15
        kj = Elem(ElMNo).ElNode(j)
        If kj >= 0 Then iNForce(kj) = iNForce(kj) + projValue * Elem(ElMNo).HourglassMode(modeId, j)
      Next j
    Next modeId
  End If
End Sub
Sub CalcElePoint(ElMNo As Long, x() As Double, y() As Double, p() As Double)
Dim b As Double: Dim im As Long
Dim j As Long
Dim s As Double


  For im = 0 To 3  'imガウス点　4個
     s = 0#
     b = 0#
     For j = 0 To 7   'j　要素各点番号
        s = s + x(j)
        b = b + y(j)
     Next j
  p(0) = s / 8#
  p(1) = b / 8#
     
     
  Next im
  


End Sub

Sub CalcStrain(ElMNo As Long, Bmat() As Double, TDisp() As Double, Stmat() As Double)
Dim im As Long: Dim jm As Long: Dim kj As Long
Dim j As Long
Dim s As Double
  For im = 0 To 3  'imガウス点　4個
     For jm = 0 To 2 '応力
        s = 0#
        For j = 0 To 15   'j　要素各点番号
           kj = Elem(ElMNo).ElNode(j)  'kj 力番号
           s = s + Bmat(jm, j, im) * TDisp(kj)   '［B］×U
        Next j
        Stmat(10 + jm, im) = s
     Next jm
     Stmat(13, im) = 1 / 3 * Sqr((Stmat(10, im) - Stmat(11, im)) ^ 2 + (Stmat(10, im)) ^ 2 + (Stmat(11, im)) ^ 2 + 1.5 * (Stmat(12, im)) ^ 2)
     Stmat(14, im) = Stmat(10, im) + Stmat(11, im)
     Stmat(15, im) = 2# * Sqr(((Stmat(10, im) - Stmat(11, im)) / 2) ^ 2 + (Stmat(12, im) / 2) ^ 2)
  Next im
       
End Sub



Private Sub UpdateTotalDisplacement()
  Dim dofId As Long
  For dofId = 0 To lastDof
    If NodeCond(dofId) = 0 Then TDisp(dofId) = TDisp(dofId) + Disp(dofId)
  Next dofId
End Sub

Private Function ComputeIncrementError() As Double
  Dim dofId As Long
  Dim incrementNorm As Double, totalNorm As Double
  incrementNorm = 0#
  totalNorm = 0#
  For dofId = 0 To lastDof
    If NodeCond(dofId) = 0 Then
      If Abs(Disp(dofId)) > incrementNorm Then incrementNorm = Abs(Disp(dofId))
      If Abs(TDisp(dofId)) > totalNorm Then totalNorm = Abs(TDisp(dofId))
    End If
  Next dofId
  If totalNorm <= 1E-30 Then
    If incrementNorm <= 1E-30 Then
      ComputeIncrementError = 0#
    Else
      ComputeIncrementError = 1#
    End If
  Else
    ComputeIncrementError = incrementNorm / totalNorm
  End If
End Function

Sub UpdateConvergenceError()
  UpdateTotalDisplacement
  ESRR = ComputeIncrementError()
End Sub

Public Sub ComputeResidual()
  Dim i As Long, j As Long, lower As Long, upper As Long
  Dim ku As Double, residualValue As Double, displacementValue As Double
  Dim workK As Double, workApplied As Double, workSupport As Double, energyScale As Double
  ResidualNormFull = 0#
  ResidualNormFree = 0#
  RelativeResidualFree = 0#
  ForceNormFree = 0#
  MaxAbsResidualFree = 0#
  EnergyError = 0#
  P6ReactionSumX = 0#
  P6ReactionSumY = 0#
  If nDof <= 0 Then Exit Sub
  P6EnsureReactionWorkspace
  If P6UseCSR And Not P3UseNonlinearResidual Then
    P6EnsureResidualVectorWorkspace
    P6CSRMatVec TDisp, P6CSRResidualProduct, P6CSROriginalValues
  End If
  For i = 0 To lastDof
    ku = 0#
    If P3UseNonlinearResidual Then
      ku = P3FinalInternalForce(i)
    ElseIf P6UseCSR Then
       ku = P6CSRResidualProduct(i)
    Else
      lower = i - BandWidth
      If lower < 0 Then lower = 0
      upper = i + BandWidth
      If upper > lastDof Then upper = lastDof
      For j = lower To upper
        If j >= i Then
          ku = ku + OriginalMat(i, j - i) * TDisp(j)
        Else
          ku = ku + OriginalMat(j, i - j) * TDisp(j)
        End If
      Next j
    End If
    residualValue = ku - OriginalForce(i)
    Reaction(i) = residualValue
    If Abs(residualValue) > ResidualNormFull Then ResidualNormFull = Abs(residualValue)
    If NodeCond(i) = 0 Then
      ResidualNormFree = ResidualNormFree + residualValue * residualValue
      If Abs(residualValue) > MaxAbsResidualFree Then MaxAbsResidualFree = Abs(residualValue)
      ForceNormFree = ForceNormFree + OriginalForce(i) * OriginalForce(i)
    Else
      If i Mod 2 = 0 Then
        P6ReactionSumX = P6ReactionSumX + residualValue
      Else
        P6ReactionSumY = P6ReactionSumY + residualValue
      End If
    End If
    displacementValue = TDisp(i)
    workK = workK + displacementValue * ku
    workApplied = workApplied + displacementValue * OriginalForce(i)
    If NodeCond(i) <> 0 Then workSupport = workSupport + displacementValue * residualValue
  Next i
  ResidualNormFree = Sqr(ResidualNormFree)
  ForceNormFree = Sqr(ForceNormFree)
  If ForceNormFree <= 1E-30 Then
    If ResidualNormFree <= 1E-30 Then
      RelativeResidualFree = 0#
    Else
      RelativeResidualFree = ResidualNormFree
    End If
  Else
    RelativeResidualFree = ResidualNormFree / ForceNormFree
  End If
  energyScale = Abs(workK)
  If Abs(workApplied) > energyScale Then energyScale = Abs(workApplied)
  If Abs(workSupport) > energyScale Then energyScale = Abs(workSupport)
  If energyScale <= 1E-30 Then energyScale = 1#
  EnergyError = Abs(workK - workApplied - workSupport) / energyScale
End Sub

Private Sub P6CaptureDiagnostics()
  Dim dofCount As Double
  P6BandStorageEntries = 0#
  P6BandStorageBytes = 0#
  P6FactorStorageBytes = 0#
  P6DenseStorageBytes = 0#
  P6EstimatedWorkMemoryBytes = 0#
  If nDof > 0 And BandWidth >= 0 Then
    dofCount = CDbl(nDof)
    P6BandStorageEntries = dofCount * (CDbl(BandWidth) + 1#)
    P6BandStorageBytes = P6BandStorageEntries * 8# ' Double 1要素の見積
    If P6FactorReady Then P6FactorStorageBytes = P6BandStorageBytes
    P6DenseStorageBytes = dofCount * dofCount * 8# ' 比較用の全行列見積
    If P6UseCSR Then
      P6EstimatedWorkMemoryBytes = P6CSRStorageBytes
      If P6SolverMode = "CSR_GMRES" Then
        P6EstimatedWorkMemoryBytes = P6EstimatedWorkMemoryBytes + dofCount * CDbl(P6GMRESRestart + 6) * 8#
      Else
        P6EstimatedWorkMemoryBytes = P6EstimatedWorkMemoryBytes + dofCount * 10# * 8#
      End If
    Else
      P6EstimatedWorkMemoryBytes = P6BandStorageBytes * 2#
    End If
  End If
  P6MeasuredStageTotalSeconds = TimeInput + TimeElementStiffness + TimeAssembly + TimeSolve + TimePlastic + TimeOutput
  P6ModelSignature = CStr(NumberOfNode) & "/" & CStr(NumberOfElement) & "/" & CStr(NumberOfMaterial) & "/" & CStr(nDof) & "/" & CStr(BandWidth)
  If Len(P6MatrixStoragePolicy) = 0 Then P6MatrixStoragePolicy = "帯域格納＋分解済み帯域因子"
  If Len(P6NodeOrderingPolicy) = 0 Then P6NodeOrderingPolicy = "原節点番号; RCM_POLICY=AUTO"
  If P6UseCSR Then
    P6OptimizationPolicy = "solver=" & P6SolverMode & "; policy=" & P6SolverPolicy & "; CSR_NNZ=" & CStr(P6CSRNNZ) & "; CSR_bytes=" & Format$(P6CSRStorageBytes, "0") _
      & "; csr_numeric=" & P6CSRNumericAssemblyStatus & "; csr_path=" & P6CSRLastAssemblyPath & "; csr_rebuilds=" & CStr(P6CSRRebuildCount) & "; csr_reuse=" & CStr(P6CSRReuseCount) _
       & "; csr_scatter_nnz=" & CStr(P6CSRScatterCount) & "; csr_diag_zero=" & CStr(P6CSRDiagonalZeroCount) & "; csr_diag_min=" & Format$(P6CSRMinimumDiagonal, "0.000E+00") & "; csr_direct_fallback=" & CStr(P6CSRDirectFallbackUsed) _
      & "; elastic_k=" & P6ElasticKCacheStatus & "; mesh_revision=" & CStr(P6MeshRevision) & "; geometry_revision=" & CStr(P6GeometryRevision) & "; material_input_revision=" & CStr(P6MaterialInputRevision) & "; material_generation=" & CStr(P6CSRMaterialGeneration) & "; tangent_generation=" & CStr(P6TangentGeneration) & "; csr_numeric_generation=" & CStr(P6CSRNumericGeneration) _
      & "; p3_active_points=" & CStr(P3ActivePlasticPointCount) & "; p3_plastic_updates=" & CStr(P3PlasticPointUpdateCount) & "; solver_eval=" & P6SolverEvaluationStatus & "; solver_eval_count=" & CStr(P6SolverEvaluationCount) & "; solver_eval_threshold=" & CStr(P6SolverReevaluationThreshold) & "; GMRES_restart=" & CStr(P6GMRESRestartRequested) & "->" & CStr(P6GMRESRestart) & "; iter=" & CStr(P6IterativeLastIterations) & "; iter_relres=" & Format$(P6IterativeLastResidual, "0.000E+00") & "; fallback=" & P6IterativeFallbackStatus & "; hg=" & Format$(P6HourglassFactor, "0.00") & "; mixed_up=" & CStr(P6MixedUP) & "; p_count=" & CStr(P6PressureCount) & "; uzawa=" & CStr(P6UzawaIterations)
  Else
    P6OptimizationPolicy = "solver=BAND_LU; factor_bytes=" & Format$(P6FactorStorageBytes, "0") _
      & "; elastic_k=" & P6ElasticKCacheStatus & "; mesh_revision=" & CStr(P6MeshRevision) & "; geometry_revision=" & CStr(P6GeometryRevision) & "; material_input_revision=" & CStr(P6MaterialInputRevision) & "; solver_eval=" & P6SolverEvaluationStatus & "; solver_eval_count=" & CStr(P6SolverEvaluationCount) & "; solver_eval_threshold=" & CStr(P6SolverReevaluationThreshold) _
      & "; p3_active_points=" & CStr(P3ActivePlasticPointCount) & "; p3_plastic_updates=" & CStr(P3PlasticPointUpdateCount) & "; material_generation=" & CStr(P6MaterialGeneration) & "; strength_generation=" & CStr(P6StrengthGeneration) & "; tangent_generation=" & CStr(P6TangentGeneration) & "; factorization/reuse=" & CStr(P6FactorizationCount) & "/" & CStr(P6FactorizationReuseCount) & "; modnewton=" & CStr(P3ModifiedNewtonSkips) & "; linesearch=" & CStr(P3LineSearchCuts) & "; joint_asm=" & CStr(P3JointAssembleCount) & "; eval_skip=" & CStr(P3EvalSkipCount) & "; dirty_k=" & CStr(P3TangentDirtyRebuildCount) & "; yield_xor=" & CStr(P3PlasticStateChangeCount) & "; elastic_fast=" & CStr(P3ElasticFastCount) _
      & "; hg_build/reuse=" & CStr(P6HourglassModeCount) & "/" & CStr(P6HourglassReuseCount) & "; scatter=" & CStr(P6ScatterN) _
      & "; t_eval/hg/asm/fac/sol/load/tan_ms=" & Format$(P6ProfEvalMs, "0") & "/" & Format$(P6ProfHgMs, "0") & "/" & Format$(P6ProfAssembleMs, "0") & "/" & Format$(P6ProfFactorMs, "0") & "/" & Format$(P6ProfSolveMs, "0") & "/" & Format$(P6ProfLoadMs, "0") & "/" & Format$(P6ProfTangentMs, "0") _
      & "; band_after_iterative=" & CStr(P6BandAfterIterative) & "; fallback=" & P6IterativeFallbackStatus & "; hg=" & Format$(P6HourglassFactor, "0.00") & "; mixed_up=" & CStr(P6MixedUP) & "; p_count=" & CStr(P6PressureCount) & "; consol=" & CStr(P6ConsolActive) & "; t=" & Format$(P6ConsolTime, "0.000E+00") & "; pmax=" & Format$(P6MaxPressure, "0.000E+00")
  End If
  If AnalysisOK Then
    P6IntegrationStatus = "解析・P1結果・P2材料結果の保存完了"
  Else
    P6IntegrationStatus = "解析失敗または未実行"
  End If
End Sub

Public Sub SaveAnalysisReport()
  Dim ws As Worksheet
  If FEM_DIAGNOSTIC_LAST_ROW >= FEM_INTEGRATED_P1_START_ROW Then
    Err.Raise vbObjectError + 3090, "FEM.SaveAnalysisReport", "診断最終行がP1開始行と重なります。FEM_DIAGNOSTIC_LAST_ROW=" & CStr(FEM_DIAGNOSTIC_LAST_ROW) & " FEM_INTEGRATED_P1_START_ROW=" & CStr(FEM_INTEGRATED_P1_START_ROW)
  End If
  Set ws = GetDiagnosticSheet()
   P6CaptureDiagnostics
   ws.range("A1:B" & CStr(FEM_DIAGNOSTIC_LAST_ROW)).ClearContents
  ws.Cells(1, 1).value2 = "2DFEM_ P0診断"
  ws.Cells(2, 1).value2 = "ResultStatus": ws.Cells(2, 2).value2 = ResultStatus
  ws.Cells(3, 1).value2 = "AnalysisOK": ws.Cells(3, 2).value2 = AnalysisOK
  ws.Cells(4, 1).value2 = "AnalysisMessage": ws.Cells(4, 2).value2 = AnalysisMessage
  ws.Cells(5, 1).value2 = "AnalysisErrorNumber": ws.Cells(5, 2).value2 = AnalysisErrorNumber
  ws.Cells(6, 1).value2 = "FailureElement": ws.Cells(6, 2).value2 = FailureElement
  ws.Cells(7, 1).value2 = "FailureGaussPoint": ws.Cells(7, 2).value2 = FailureGaussPoint
  ws.Cells(8, 1).value2 = "FailureIncrement": ws.Cells(8, 2).value2 = FailureIncrement
  ws.Cells(9, 1).value2 = "FailureIteration": ws.Cells(9, 2).value2 = FailureIteration
  ws.Cells(10, 1).value2 = "NumberOfNode": ws.Cells(10, 2).value2 = NumberOfNode
  ws.Cells(11, 1).value2 = "NumberOfElement": ws.Cells(11, 2).value2 = NumberOfElement
  ws.Cells(12, 1).value2 = "NumberOfMaterial": ws.Cells(12, 2).value2 = NumberOfMaterial
  ws.Cells(13, 1).value2 = "NumberOfFreeNode": ws.Cells(13, 2).value2 = NumberOfFreeNode
  ws.Cells(14, 1).value2 = "nDof": ws.Cells(14, 2).value2 = nDof
  ws.Cells(15, 1).value2 = "lastDof": ws.Cells(15, 2).value2 = lastDof
  ws.Cells(16, 1).value2 = "BandWidth": ws.Cells(16, 2).value2 = BandWidth
  ws.Cells(17, 1).value2 = "RelativeResidualFree": ws.Cells(17, 2).value2 = RelativeResidualFree
  ws.Cells(18, 1).value2 = "ResidualNormFull": ws.Cells(18, 2).value2 = ResidualNormFull
  ws.Cells(19, 1).value2 = "MaxAbsResidualFree": ws.Cells(19, 2).value2 = MaxAbsResidualFree
  ws.Cells(20, 1).value2 = "EnergyError": ws.Cells(20, 2).value2 = EnergyError  ' 無次元 |u?R| / (|F||u|)
  ws.Cells(21, 1).value2 = "MatrixSymmetryError": ws.Cells(21, 2).value2 = MatrixSymmetryError
  ws.Cells(22, 1).value2 = "TotalAppliedLoadX": ws.Cells(22, 2).value2 = TotalAppliedLoadX
   ws.Cells(23, 1).value2 = "TotalAppliedLoadY": ws.Cells(23, 2).value2 = TotalAppliedLoadY
   ws.Cells(24, 1).value2 = "TotalSelfWeight": ws.Cells(24, 2).value2 = TotalSelfWeight
   ws.Cells(25, 1).value2 = "MinDetJCorner": ws.Cells(25, 2).value2 = MinDetJCorner
   ws.Cells(26, 1).value2 = "MaxDetJCorner": ws.Cells(26, 2).value2 = MaxDetJCorner
   ws.Cells(27, 1).value2 = "MinDetJGauss": ws.Cells(27, 2).value2 = MinDetJGauss
   ws.Cells(28, 1).value2 = "MaxDetJGauss": ws.Cells(28, 2).value2 = MaxDetJGauss
   ws.Cells(29, 1).value2 = "ModelRevision": ws.Cells(29, 2).value2 = ModelRevision
   ws.Cells(30, 1).value2 = "ResultRevision": ws.Cells(30, 2).value2 = ResultRevision
   ws.Cells(31, 1).value2 = "TimeInput": ws.Cells(31, 2).value2 = TimeInput
   ws.Cells(32, 1).value2 = "TimeElementStiffness": ws.Cells(32, 2).value2 = TimeElementStiffness
   ws.Cells(33, 1).value2 = "TimeAssembly": ws.Cells(33, 2).value2 = TimeAssembly
   ws.Cells(34, 1).value2 = "TimeSolve": ws.Cells(34, 2).value2 = TimeSolve
   ws.Cells(35, 1).value2 = "TimePlastic": ws.Cells(35, 2).value2 = TimePlastic
   ws.Cells(36, 1).value2 = "TimeOutput": ws.Cells(36, 2).value2 = TimeOutput
   ws.Cells(37, 1).value2 = "P1Formulation": ws.Cells(37, 2).value2 = P1Formulation
   ws.Cells(38, 1).value2 = "P1ResultReady": ws.Cells(38, 2).value2 = P1ResultReady
   ws.Cells(39, 1).value2 = "P1ResultModelRevision": ws.Cells(39, 2).value2 = P1ResultModelRevision
   ws.Cells(40, 1).value2 = "P1ModelNote": ws.Cells(40, 2).value2 = "2D平面ひずみ。σz=ν(σx+σy)、3D主応力・3D相当応力を保存。曲げ・横せん断力は現行DOFの対象外。"
   ws.Cells(41, 1).value2 = "P3SuccessfulIncrementCount": ws.Cells(41, 2).value2 = P3SuccessfulIncrementCount
   ws.Cells(42, 1).value2 = "P3GlobalIterationCount": ws.Cells(42, 2).value2 = P3GlobalIterationCount
   ws.Cells(43, 1).value2 = "P3RetryCount": ws.Cells(43, 2).value2 = P3RetryCount
   ws.Cells(44, 1).value2 = "P3LastConvergedLoadFactor": ws.Cells(44, 2).value2 = P3LastConvergedLoadFactor
   ws.Cells(45, 1).value2 = "P3CurrentStrengthFactor": ws.Cells(45, 2).value2 = P3CurrentStrengthFactor
   ws.Cells(46, 1).value2 = "P3UseNonlinearResidual": ws.Cells(46, 2).value2 = P3UseNonlinearResidual
   ws.Cells(47, 1).value2 = "P3DefaultIncrementCount": ws.Cells(47, 2).value2 = P3DefaultIncrementCount
   ws.Cells(48, 1).value2 = "P3StateNote": ws.Cells(48, 2).value2 = "mStmat/CommittedStateは確定値、Stmat/TDispは試行値。収束後のみコミットし、失敗時は直前確定状態へ復元。"
   ws.Cells(49, 1).value2 = "P3MaxTrialDisp": ws.Cells(49, 2).value2 = P3MaxTrialDisp
   ws.Cells(50, 1).value2 = "P3MaxCorrection": ws.Cells(50, 2).value2 = P3MaxCorrection
   ws.Cells(51, 1).value2 = "P3LastRelativeResidual": ws.Cells(51, 2).value2 = P3LastRelativeResidual
   ws.Cells(52, 1).value2 = "P5ModelScale": ws.Cells(52, 2).value2 = P5ModelScale
   ws.Cells(53, 1).value2 = "P5GeometryTolerance": ws.Cells(53, 2).value2 = P5GeometryTolerance
   ws.Cells(54, 1).value2 = "P5AreaTolerance": ws.Cells(54, 2).value2 = P5AreaTolerance
   ws.Cells(55, 1).value2 = "P5MinElementArea": ws.Cells(55, 2).value2 = P5MinElementArea
   ws.Cells(56, 1).value2 = "P5MinJacobian": ws.Cells(56, 2).value2 = P5MinJacobian
   ws.Cells(57, 1).value2 = "P5MaxAspectRatio": ws.Cells(57, 2).value2 = P5MaxAspectRatio
   ws.Cells(58, 1).value2 = "P5MinAngleDeg": ws.Cells(58, 2).value2 = P5MinAngleDeg
   ws.Cells(59, 1).value2 = "P5MaxAngleDeg": ws.Cells(59, 2).value2 = P5MaxAngleDeg
   ws.Cells(60, 1).value2 = "P5BoundaryEdgeCount": ws.Cells(60, 2).value2 = P5BoundaryEdgeCount
   ws.Cells(61, 1).value2 = "P5SharedEdgeCount": ws.Cells(61, 2).value2 = P5SharedEdgeCount
   ws.Cells(62, 1).value2 = "P5MaterialInterfaceCount": ws.Cells(62, 2).value2 = P5MaterialInterfaceCount
   ws.Cells(63, 1).value2 = "P5IsolatedNodeCount": ws.Cells(63, 2).value2 = P5IsolatedNodeCount
   ws.Cells(64, 1).value2 = "P5ConnectedComponentCount": ws.Cells(64, 2).value2 = P5ConnectedComponentCount
   ws.Cells(65, 1).value2 = "P5BoundarySelfIntersectionCount": ws.Cells(65, 2).value2 = P5BoundarySelfIntersectionCount
   ws.Cells(66, 1).value2 = "P5LocateCallCount": ws.Cells(66, 2).value2 = P5LocateCallCount
   ws.Cells(67, 1).value2 = "P5LocateTotalIterations": ws.Cells(67, 2).value2 = P5LocateTotalIterations
   ws.Cells(68, 1).value2 = "P5LocateMaxIterations": ws.Cells(68, 2).value2 = P5LocateMaxIterations
   ws.Cells(69, 1).value2 = "P5LocateLastIterations": ws.Cells(69, 2).value2 = P5LocateLastIterations
   ws.Cells(70, 1).value2 = "P5LocateLastElement": ws.Cells(70, 2).value2 = P5LocateLastElement
   ws.Cells(71, 1).value2 = "P5LocateLastFailure": ws.Cells(71, 2).value2 = P5LocateLastFailure
   ws.Cells(72, 1).value2 = "P6MatrixStoragePolicy": ws.Cells(72, 2).value2 = P6MatrixStoragePolicy
   ws.Cells(73, 1).value2 = "P6BandStorageEntries": ws.Cells(73, 2).value2 = P6BandStorageEntries
   ws.Cells(74, 1).value2 = "P6BandStorageBytes": ws.Cells(74, 2).value2 = P6BandStorageBytes
   ws.Cells(75, 1).value2 = "P6DenseStorageBytes": ws.Cells(75, 2).value2 = P6DenseStorageBytes
   ws.Cells(76, 1).value2 = "P6EstimatedWorkMemoryBytes": ws.Cells(76, 2).value2 = P6EstimatedWorkMemoryBytes
   ws.Cells(77, 1).value2 = "P6BandSolveCallCount": ws.Cells(77, 2).value2 = P6BandSolveCallCount
   ws.Cells(78, 1).value2 = "P6FactorizationCount": ws.Cells(78, 2).value2 = P6FactorizationCount
   ws.Cells(79, 1).value2 = "P6FactorizationReuseCount": ws.Cells(79, 2).value2 = P6FactorizationReuseCount
   ws.Cells(80, 1).value2 = "P6RenumberingApplied": ws.Cells(80, 2).value2 = P6RenumberingApplied
   ws.Cells(81, 1).value2 = "P6NodeOrderingPolicy": ws.Cells(81, 2).value2 = P6NodeOrderingPolicy
   ws.Cells(82, 1).value2 = "P6MeasuredStageTotalSeconds": ws.Cells(82, 2).value2 = P6MeasuredStageTotalSeconds
   ws.Cells(83, 1).value2 = "P6TimeInput": ws.Cells(83, 2).value2 = TimeInput
   ws.Cells(84, 1).value2 = "P6TimeElementStiffness": ws.Cells(84, 2).value2 = TimeElementStiffness
   ws.Cells(85, 1).value2 = "P6TimeAssembly": ws.Cells(85, 2).value2 = TimeAssembly
   ws.Cells(86, 1).value2 = "P6TimeSolve": ws.Cells(86, 2).value2 = TimeSolve
   ws.Cells(87, 1).value2 = "P6TimePlastic": ws.Cells(87, 2).value2 = TimePlastic
   ws.Cells(88, 1).value2 = "P6TimeOutput": ws.Cells(88, 2).value2 = TimeOutput
   ws.Cells(89, 1).value2 = "P6ModelSignature": ws.Cells(89, 2).value2 = P6ModelSignature
   ws.Cells(90, 1).value2 = "P6IntegrationStatus": ws.Cells(90, 2).value2 = P6IntegrationStatus
   ws.Cells(91, 1).value2 = "P6RCMOriginalBandwidth": ws.Cells(91, 2).value2 = P6RCMOriginalBandwidth
   ws.Cells(92, 1).value2 = "P6RCMCandidateBandwidth": ws.Cells(92, 2).value2 = P6RCMCandidateBandwidth
   ws.Cells(93, 1).value2 = "P6RCMStatus": ws.Cells(93, 2).value2 = P6RCMStatus
   ws.Cells(94, 1).value2 = "P6FactorizationReady": ws.Cells(94, 2).value2 = P6FactorReady
   ws.Cells(95, 1).value2 = "P6SolverMode": ws.Cells(95, 2).value2 = P6SolverMode
   ws.Cells(96, 1).value2 = "P6OptimizationPolicy": ws.Cells(96, 2).value2 = P6OptimizationPolicy
   ws.Cells(97, 1).value2 = "P6ElasticKCacheStatus": ws.Cells(97, 2).value2 = P6ElasticKCacheStatus
   ws.Cells(98, 1).value2 = "P6SolverMemoryLimitBytes": ws.Cells(98, 2).value2 = P6SolverMemoryLimitBytes
   ws.Cells(99, 1).value2 = "P6SolverPolicy": ws.Cells(99, 2).value2 = P6SolverPolicy
   ws.Cells(100, 1).value2 = "P6SolverEvaluationStatus": ws.Cells(100, 2).value2 = P6SolverEvaluationStatus
   ws.Cells(101, 1).value2 = "P6CacheMissReason": ws.Cells(101, 2).value2 = P6CacheMissReason
   ws.Cells(102, 1).value2 = "P6QuadratureArea": ws.Cells(102, 2).value2 = P6QuadratureArea
   ws.Cells(103, 1).value2 = "P6MeshRevision": ws.Cells(103, 2).value2 = P6MeshRevision
   ws.Cells(104, 1).value2 = "P6GeometryRevision": ws.Cells(104, 2).value2 = P6GeometryRevision
   ws.Cells(105, 1).value2 = "P6MaterialGeneration": ws.Cells(105, 2).value2 = P6MaterialGeneration
   ws.Cells(105, 3).value2 = "P6StrengthGeneration": ws.Cells(105, 4).value2 = P6StrengthGeneration
   ws.Cells(106, 1).value2 = "P6MaterialInputRevision": ws.Cells(106, 2).value2 = P6MaterialInputRevision
   ws.Cells(107, 1).value2 = "P6IterativeFallbackStatus": ws.Cells(107, 2).value2 = P6IterativeFallbackStatus
   ws.Cells(108, 1).value2 = "FEM_BUILD_STAMP": ws.Cells(108, 2).value2 = FEM_BUILD_STAMP
   ws.Cells(109, 1).value2 = "Q8_HOURGLASS_FACTOR": ws.Cells(109, 2).value2 = P6HourglassFactor
   ws.Cells(110, 1).value2 = "P6HourglassModeCount": ws.Cells(110, 2).value2 = P6HourglassModeCount
   ws.Cells(111, 1).value2 = "MIXED_UP": ws.Cells(111, 2).value2 = P6MixedUP
   ws.Cells(112, 1).value2 = "P6PressureCount": ws.Cells(112, 2).value2 = P6PressureCount
   ws.Cells(113, 1).value2 = "P6UzawaIterations": ws.Cells(113, 2).value2 = P6UzawaIterations
   ws.Cells(114, 1).value2 = "P6ConsolActive": ws.Cells(114, 2).value2 = P6ConsolActive
   ws.Cells(115, 1).value2 = "P6ConsolTime": ws.Cells(115, 2).value2 = P6ConsolTime
   ws.Cells(116, 1).value2 = "P6ConsolStep": ws.Cells(116, 2).value2 = P6ConsolStep
   ws.Cells(117, 1).value2 = "P6MaxPressure": ws.Cells(117, 2).value2 = P6MaxPressure
   ws.Cells(118, 1).value2 = "P6DrainCount": ws.Cells(118, 2).value2 = P6DrainCount
   ws.Cells(119, 1).value2 = "P6WeightMode": ws.Cells(119, 2).value2 = P6WeightMode
   ws.Cells(120, 1).value2 = "P6WeightNote": ws.Cells(120, 2).value2 = P6WeightNote
   ws.Cells(121, 1).value2 = "P3MaxBoundaryDispError": ws.Cells(121, 2).value2 = P3MaxBoundaryDispError
   ws.Cells(122, 1).value2 = "P3RelativeBoundaryDispError": ws.Cells(122, 2).value2 = P3RelativeBoundaryDispError
   ws.Cells(123, 1).value2 = "P6ReactionSumX": ws.Cells(123, 2).value2 = P6ReactionSumX
   ws.Cells(124, 1).value2 = "P6ReactionSumY": ws.Cells(124, 2).value2 = P6ReactionSumY
   ws.Cells(125, 1).value2 = "P3SrmEnabled": ws.Cells(125, 2).value2 = P3SrmEnabled
   ws.Cells(126, 1).value2 = "P3SrmTrialCount": ws.Cells(126, 2).value2 = P3SrmTrialCount
   ws.Cells(127, 1).value2 = "P3SrmLower": ws.Cells(127, 2).value2 = P3SrmLower
   ws.Cells(128, 1).value2 = "P3SrmUpper": ws.Cells(128, 2).value2 = P3SrmUpper
   ws.Cells(129, 1).value2 = "P3SrmNote": ws.Cells(129, 2).value2 = P3SrmNote
   ws.Cells(130, 1).value2 = "P3StagePlanText": ws.Cells(130, 2).value2 = P3StagePlanText
   ws.Cells(131, 1).value2 = "P6HourglassReuseCount": ws.Cells(131, 2).value2 = P6HourglassReuseCount
   ws.Cells(132, 1).value2 = "P6ScatterN": ws.Cells(132, 2).value2 = P6ScatterN
   ws.Cells(133, 1).value2 = "P6ProfHgMs": ws.Cells(133, 2).value2 = P6ProfHgMs
   ws.Cells(134, 1).value2 = "P6ProfAssembleMs": ws.Cells(134, 2).value2 = P6ProfAssembleMs
   ws.Cells(135, 1).value2 = "P6ProfFactorMs": ws.Cells(135, 2).value2 = P6ProfFactorMs
   ws.Cells(136, 1).value2 = "P6ProfSolveMs": ws.Cells(136, 2).value2 = P6ProfSolveMs
   ws.Cells(137, 1).value2 = "P6ProfLoadMs": ws.Cells(137, 2).value2 = P6ProfLoadMs
   ws.Cells(138, 1).value2 = "P6ProfTangentMs": ws.Cells(138, 2).value2 = P6ProfTangentMs
   ws.Cells(139, 1).value2 = "P6ProfEvalCount": ws.Cells(139, 2).value2 = P6ProfEvalCount
   ws.Cells(140, 1).value2 = "P6ProfEvalMs": ws.Cells(140, 2).value2 = P6ProfEvalMs
   ws.Cells(141, 1).value2 = "FLOW_POLICY": ws.Cells(141, 2).value2 = P3FlowPolicyText()
   ws.Cells(142, 1).value2 = "FOS_INTERPRETATION": ws.Cells(142, 2).value2 = P3FosInterpretation
   ws.Cells(143, 1).value2 = "SRM_TRIAL_CLASS": ws.Cells(143, 2).value2 = P3SrmTrialClassification
   ws.Cells(144, 1).value2 = "FAILURE_KIND": ws.Cells(144, 2).value2 = P3FailureKind
   ws.Cells(145, 1).value2 = "FAILED_TRIAL_LAMBDA": ws.Cells(145, 2).value2 = P3FailureLambda
   ws.Cells(146, 1).value2 = "FAILED_TRIAL_RELRES": ws.Cells(146, 2).value2 = P3FailureResidual
   ws.Cells(147, 1).value2 = "FAILED_TRIAL_LINEAR_RELRES": ws.Cells(147, 2).value2 = P3FailureLinearResidual
   ws.Cells(148, 1).value2 = "FAILED_TRIAL_CORRECTION": ws.Cells(148, 2).value2 = P3FailureCorrection
   ws.Cells(149, 1).value2 = "FAILED_TRIAL_PLASTIC_POINTS": ws.Cells(149, 2).value2 = P3FailurePlasticPoints
   ws.Cells(150, 1).value2 = "FAILED_TRIAL_UMAX": ws.Cells(150, 2).value2 = P3FailureMaxDisp
   If P6ReadSetting("OUTPUT_AUTOFIT", 0#) <> 0# Then ws.columns("A:B").AutoFit
End Sub

Public Function P1Atan2(ByVal yValue As Double, ByVal xValue As Double) As Double
  If Abs(xValue) <= 1E-30 Then
    If yValue > 0# Then
      P1Atan2 = 1.5707963267949
    ElseIf yValue < 0# Then
      P1Atan2 = -1.5707963267949
    Else
      P1Atan2 = 0#
    End If
  Else
    P1Atan2 = Atn(yValue / xValue)
    If xValue < 0# Then
      If yValue >= 0# Then
        P1Atan2 = P1Atan2 + 3.14159265358979
      Else
        P1Atan2 = P1Atan2 - 3.14159265358979
      End If
    End If
  End If
End Function

Public Sub P1CornerShape(ByVal xi As Double, ByVal eta As Double, ByRef shapeValues() As Double)
  shapeValues(0) = 0.25 * (1# - xi) * (1# - eta)
  shapeValues(1) = 0.25 * (1# + xi) * (1# - eta)
  shapeValues(2) = 0.25 * (1# + xi) * (1# + eta)
  shapeValues(3) = 0.25 * (1# - xi) * (1# + eta)
End Sub

Public Sub BuildGaussToNodeExtrapolationMatrix(ByRef extrapolation() As Double)
  Dim augmented(0 To 3, 0 To 7) As Double
  Dim shapeValues(0 To 3) As Double
  Dim rowNo As Long, colNo As Long, pivotRow As Long, candidateRow As Long
  Dim swapValue As Double, pivotValue As Double, factorValue As Double
  Dim xiValues(0 To 3) As Double, etaValues(0 To 3) As Double
  xiValues(0) = -0.577350269189626: etaValues(0) = -0.577350269189626
  xiValues(1) = 0.577350269189626: etaValues(1) = -0.577350269189626
  xiValues(2) = 0.577350269189626: etaValues(2) = 0.577350269189626
  xiValues(3) = -0.577350269189626: etaValues(3) = 0.577350269189626
  For rowNo = 0 To 3
    P1CornerShape xiValues(rowNo), etaValues(rowNo), shapeValues
    For colNo = 0 To 3
      augmented(rowNo, colNo) = shapeValues(colNo)
    Next colNo
    For colNo = 0 To 3
      If colNo = rowNo Then
        augmented(rowNo, colNo + 4) = 1#
      Else
        augmented(rowNo, colNo + 4) = 0#
      End If
    Next colNo
  Next rowNo
  For rowNo = 0 To 3
    pivotRow = rowNo
    For candidateRow = rowNo + 1 To 3
      If Abs(augmented(candidateRow, rowNo)) > Abs(augmented(pivotRow, rowNo)) Then pivotRow = candidateRow
    Next candidateRow
    If Abs(augmented(pivotRow, rowNo)) <= 0.000000000001 Then Err.Raise vbObjectError + 3030, "FEM.BuildGaussToNodeExtrapolationMatrix", "Gauss点から節点への外挿行列が特異です。"
    If pivotRow <> rowNo Then
      For colNo = 0 To 7
        swapValue = augmented(rowNo, colNo)
        augmented(rowNo, colNo) = augmented(pivotRow, colNo)
        augmented(pivotRow, colNo) = swapValue
      Next colNo
    End If
    pivotValue = augmented(rowNo, rowNo)
    For colNo = 0 To 7
      augmented(rowNo, colNo) = augmented(rowNo, colNo) / pivotValue
    Next colNo
    For candidateRow = 0 To 3
      If candidateRow <> rowNo Then
        factorValue = augmented(candidateRow, rowNo)
        If Abs(factorValue) > 1E-30 Then
          For colNo = 0 To 7
            augmented(candidateRow, colNo) = augmented(candidateRow, colNo) - factorValue * augmented(rowNo, colNo)
          Next colNo
        End If
      End If
    Next candidateRow
  Next rowNo
  For rowNo = 0 To 3
    For colNo = 0 To 3
      extrapolation(rowNo, colNo) = augmented(rowNo, colNo + 4)
    Next colNo
  Next rowNo
End Sub

Public Sub ExtrapolateElementResult(ByRef gaussValues() As Double, ByRef nodeValues() As Double, ByRef extrapolation() As Double)
  Dim nodeNo As Long, gaussNo As Long
  For nodeNo = 0 To 3
    nodeValues(nodeNo) = 0#
    For gaussNo = 0 To P1_GAUSS_COUNT - 1
      nodeValues(nodeNo) = nodeValues(nodeNo) + extrapolation(nodeNo, gaussNo) * gaussValues(gaussNo)
    Next gaussNo
  Next nodeNo
End Sub

Private Function P6EnsureVariantBuffer(ByRef buffer() As Variant, ByRef storedRows As Long, ByRef storedColumns As Long, _
                                       ByVal requiredRows As Long, ByVal requiredColumns As Long) As Boolean
  If storedRows = requiredRows And storedColumns = requiredColumns And storedRows > 0 And storedColumns > 0 Then
    P6EnsureVariantBuffer = True
    Exit Function
  End If
  ReDim buffer(1 To requiredRows, 1 To requiredColumns)
  storedRows = requiredRows
  storedColumns = requiredColumns
  P6EnsureVariantBuffer = False
End Function

Private Sub P6ClearVariantBuffer(ByRef buffer() As Variant, ByVal rowCount As Long, ByVal columnCount As Long)
  Dim rowNo As Long, columnNo As Long
  For rowNo = 1 To rowCount
    For columnNo = 1 To columnCount
      buffer(rowNo, columnNo) = Empty
    Next columnNo
  Next rowNo
End Sub

Private Function P1ElementArea(ByRef elementData As Element_Data) As Double
  Dim signedArea As Double
  signedArea = 0.5 * (elementData.x(0) * elementData.y(1) + elementData.x(1) * elementData.y(2) + elementData.x(2) * elementData.y(3) + elementData.x(3) * elementData.y(0) - elementData.y(0) * elementData.x(1) - elementData.y(1) * elementData.x(2) - elementData.y(2) * elementData.x(3) - elementData.y(3) * elementData.x(0))
  P1ElementArea = Abs(signedArea)
End Function

Private Function P1ElementAngle(ByRef elementData As Element_Data) As Double
  Dim dxValue As Double, dyValue As Double
  dxValue = elementData.x(1) - elementData.x(0)
  dyValue = elementData.y(1) - elementData.y(0)
  If Abs(dxValue) <= 1E-30 Then
    If dyValue >= 0# Then
      P1ElementAngle = 1.5707963267949
    Else
      P1ElementAngle = -1.5707963267949
    End If
  Else
    P1ElementAngle = P1Atan2(dyValue, dxValue)
  End If
End Function

Private Sub P1TransformStress(ByVal stressX As Double, ByVal stressY As Double, ByVal shearXY As Double, ByVal elementAngle As Double, ByRef localX As Double, ByRef localY As Double, ByRef localShear As Double)
  Dim cosineValue As Double, sineValue As Double
  cosineValue = Cos(elementAngle)
  sineValue = Sin(elementAngle)
  localX = cosineValue * cosineValue * stressX + sineValue * sineValue * stressY + 2# * sineValue * cosineValue * shearXY
  localY = sineValue * sineValue * stressX + cosineValue * cosineValue * stressY - 2# * sineValue * cosineValue * shearXY
  localShear = -sineValue * cosineValue * stressX + sineValue * cosineValue * stressY + (cosineValue * cosineValue - sineValue * sineValue) * shearXY
End Sub

Public Sub P1SortPrincipal3(ByVal value1 As Double, ByVal value2 As Double, ByVal value3 As Double, ByRef maximumValue As Double, ByRef middleValue As Double, ByRef minimumValue As Double)
  Dim tempValue As Double
  maximumValue = value1
  middleValue = value2
  minimumValue = value3
  If maximumValue < middleValue Then
    tempValue = maximumValue: maximumValue = middleValue: middleValue = tempValue
  End If
  If middleValue < minimumValue Then
    tempValue = middleValue: middleValue = minimumValue: minimumValue = tempValue
  End If
  If maximumValue < middleValue Then
    tempValue = maximumValue: maximumValue = middleValue: middleValue = tempValue
  End If
End Sub

Public Function P1VonMises3D(ByVal stressX As Double, ByVal stressY As Double, ByVal stressZ As Double, ByVal shearXY As Double) As Double
  P1VonMises3D = Sqr(0.5 * ((stressX - stressY) ^ 2 + (stressY - stressZ) ^ 2 + (stressZ - stressX) ^ 2) + 3# * shearXY ^ 2)
End Function

Private Function P1VonMises(ByVal stressX As Double, ByVal stressY As Double, ByVal shearXY As Double) As Double
  P1VonMises = Sqr(stressX * stressX - stressX * stressY + stressY * stressY + 3# * shearXY * shearXY)
End Function

Public Sub RecoverGaussResult(ByVal elementId As Long)
  Dim gaussNo As Long, resultId As Long
  Dim thicknessValue As Double, angleValue As Double
  Dim globalX As Double, globalY As Double, globalShear As Double, sigmaZValue As Double
  Dim localX As Double, localY As Double, localShear As Double
  Dim centerValue As Double, radiusValue As Double, sigma1Value As Double, sigma2Value As Double
  Dim sigmaMax3D As Double, sigmaMid3D As Double, sigmaMin3D As Double
  Dim principalAngle As Double, vonMisesValue As Double
  If elementId < 0 Or elementId >= NumberOfElement Then Err.Raise vbObjectError + 3031, "FEM.RecoverGaussResult", "要素番号が範囲外です。"
  If Not P3IsElementActive(elementId) Then Exit Sub
  If Elem(elementId).IsJoint Then Exit Sub
  If Elem(elementId).MatNo < 0 Or Elem(elementId).MatNo > NumberOfMaterial - 1 Then Err.Raise vbObjectError + 3031, "FEM.RecoverGaussResult", "要素の材料番号が範囲外です。要素=" & CStr(elementId + 1)
  thicknessValue = Material(Elem(elementId).MatNo).thickness
  angleValue = P1ElementAngle(Elem(elementId))
  For gaussNo = 0 To P1_GAUSS_COUNT - 1
    globalX = Elem(elementId).Stmat(0, gaussNo)
    globalY = Elem(elementId).Stmat(1, gaussNo)
    globalShear = Elem(elementId).Stmat(2, gaussNo)
    sigmaZValue = Elem(elementId).Stmat(16, gaussNo)
    GaussGlobalResult(elementId, gaussNo, 0) = globalX
    GaussGlobalResult(elementId, gaussNo, 1) = globalY
    GaussGlobalResult(elementId, gaussNo, 2) = globalShear
    P1TransformStress globalX, globalY, globalShear, angleValue, localX, localY, localShear
    centerValue = 0.5 * (localX + localY)
    radiusValue = Sqr((0.5 * (localX - localY)) ^ 2 + localShear ^ 2)
    sigma1Value = centerValue + radiusValue
    sigma2Value = centerValue - radiusValue
    P1SortPrincipal3 sigma1Value, sigma2Value, sigmaZValue, sigmaMax3D, sigmaMid3D, sigmaMin3D
    principalAngle = 0.5 * P1Atan2(2# * localShear, localX - localY) * 180# / 3.14159265358979
    vonMisesValue = P1VonMises3D(localX, localY, sigmaZValue, localShear)
    GaussResult(elementId, gaussNo, P1_RESULT_N1_LOCAL) = localX * thicknessValue
    GaussResult(elementId, gaussNo, P1_RESULT_N2_LOCAL) = localY * thicknessValue
    GaussResult(elementId, gaussNo, P1_RESULT_N12_LOCAL) = localShear * thicknessValue
    GaussResult(elementId, gaussNo, P1_RESULT_SIGMA1_TOP) = sigma1Value
    GaussResult(elementId, gaussNo, P1_RESULT_SIGMA2_TOP) = sigma2Value
    GaussResult(elementId, gaussNo, P1_RESULT_TAU12_TOP) = localShear
    GaussResult(elementId, gaussNo, P1_RESULT_SIGMA1_BOTTOM) = sigma1Value
    GaussResult(elementId, gaussNo, P1_RESULT_SIGMA2_BOTTOM) = sigma2Value
    GaussResult(elementId, gaussNo, P1_RESULT_TAU12_BOTTOM) = localShear
    GaussResult(elementId, gaussNo, P1_RESULT_PRINCIPAL_ANGLE_TOP) = principalAngle
    GaussResult(elementId, gaussNo, P1_RESULT_VON_MISES_TOP) = vonMisesValue
    GaussResult(elementId, gaussNo, P1_RESULT_PRINCIPAL_ANGLE_BOTTOM) = principalAngle
    GaussResult(elementId, gaussNo, P1_RESULT_VON_MISES_BOTTOM) = vonMisesValue
    GaussResult(elementId, gaussNo, P1_RESULT_SIGMA_Z) = sigmaZValue
    GaussResult(elementId, gaussNo, P1_RESULT_SIGMA_MAX_3D) = sigmaMax3D
    GaussResult(elementId, gaussNo, P1_RESULT_SIGMA_MID_3D) = sigmaMid3D
    GaussResult(elementId, gaussNo, P1_RESULT_SIGMA_MIN_3D) = sigmaMin3D
    GaussWeight(elementId, gaussNo) = Abs(Elem(elementId).dj(gaussNo))
  Next gaussNo
  For resultId = 0 To P1_RESULT_COUNT - 1
    ElementAverageResult(elementId, resultId) = 0#
    For gaussNo = 0 To P1_GAUSS_COUNT - 1
      ElementAverageResult(elementId, resultId) = ElementAverageResult(elementId, resultId) + GaussResult(elementId, gaussNo, resultId)
    Next gaussNo
    ElementAverageResult(elementId, resultId) = ElementAverageResult(elementId, resultId) / P1_GAUSS_COUNT
  Next resultId
End Sub
Public Sub RecoverElementAverage(ByVal elementId As Long)
  Dim resultId As Long, gaussNo As Long
  For resultId = 0 To P1_RESULT_COUNT - 1
    ElementAverageResult(elementId, resultId) = 0#
    For gaussNo = 0 To P1_GAUSS_COUNT - 1
      ElementAverageResult(elementId, resultId) = ElementAverageResult(elementId, resultId) + GaussResult(elementId, gaussNo, resultId)
    Next gaussNo
    ElementAverageResult(elementId, resultId) = ElementAverageResult(elementId, resultId) / P1_GAUSS_COUNT
  Next resultId
End Sub

Public Sub AccumulateNodalResult(ByVal nodeId As Long, ByVal resultId As Long, ByVal resultValue As Double, ByVal elementWeight As Double)
  If nodeId < 0 Or nodeId >= NumberOfFreeNode Then Err.Raise vbObjectError + 3032, "FEM.AccumulateNodalResult", "回復先節点番号が範囲外です。"
  NodalResult(nodeId, resultId) = NodalResult(nodeId, resultId) + elementWeight * resultValue
  If resultId = 0 Then NodalResultWeight(nodeId) = NodalResultWeight(nodeId) + elementWeight
End Sub

Public Sub FinalizeNodalResult()
  Dim nodeId As Long, resultId As Long
  For nodeId = 0 To NumberOfFreeNode - 1
    If NodalResultWeight(nodeId) > 0# Then
      NodalX(nodeId) = NodalX(nodeId) / NodalResultWeight(nodeId)
      NodalY(nodeId) = NodalY(nodeId) / NodalResultWeight(nodeId)
      For resultId = 0 To P1_RESULT_COUNT - 1
        NodalResult(nodeId, resultId) = NodalResult(nodeId, resultId) / NodalResultWeight(nodeId)
      Next resultId
      For resultId = 0 To 2
        NodalGlobalResult(nodeId, resultId) = NodalGlobalResult(nodeId, resultId) / NodalResultWeight(nodeId)
      Next resultId
    End If
  Next nodeId
End Sub

Public Sub BuildP1Results()
  Dim elementId As Long, gaussNo As Long, resultId As Long, localNode As Long
  Dim nodeId As Long, globalId As Long
  Dim extrapolation(0 To 3, 0 To 3) As Double
  Dim gaussValues() As Double, cornerValues() As Double
  Dim elementWeight As Double, midValue As Double
  FEMLastProc = "BuildP1Results"
  If Not AnalysisOK Then Err.Raise vbObjectError + 3033, "FEM.BuildP1Results", "解析が成功していないためP1結果を作成できません。"
  If NumberOfElement <= 0 Or NumberOfFreeNode <= 0 Then Err.Raise vbObjectError + 3034, "FEM.BuildP1Results", "P1結果の対象要素または節点がありません。"
  ReDim GaussResult(0 To NumberOfElement - 1, 0 To P1_GAUSS_COUNT - 1, 0 To P1_RESULT_COUNT - 1)
  ReDim GaussGlobalResult(0 To NumberOfElement - 1, 0 To P1_GAUSS_COUNT - 1, 0 To 2)
  ReDim GaussWeight(0 To NumberOfElement - 1, 0 To P1_GAUSS_COUNT - 1)
  ReDim ElementAverageResult(0 To NumberOfElement - 1, 0 To P1_RESULT_COUNT - 1)
  ReDim ElementNodeResult(0 To NumberOfElement - 1, 0 To 7, 0 To P1_RESULT_COUNT - 1)
  ReDim NodalResult(0 To NumberOfFreeNode - 1, 0 To P1_RESULT_COUNT - 1)
  ReDim NodalGlobalResult(0 To NumberOfFreeNode - 1, 0 To 2)
  ReDim NodalResultWeight(0 To NumberOfFreeNode - 1)
  ReDim NodalX(0 To NumberOfFreeNode - 1)
  ReDim NodalY(0 To NumberOfFreeNode - 1)
  ReDim DisplacementMagnitude(0 To NumberOfFreeNode - 1)
  ReDim gaussValues(0 To P1_GAUSS_COUNT - 1)
  ReDim cornerValues(0 To 3)
  For resultId = 0 To P1_RESULT_COUNT - 1
    P1ResultAvailable(resultId) = False
  Next resultId
  P1ResultAvailable(P1_RESULT_N1_LOCAL) = True
  P1ResultAvailable(P1_RESULT_N2_LOCAL) = True
  P1ResultAvailable(P1_RESULT_N12_LOCAL) = True
  P1ResultAvailable(P1_RESULT_SIGMA1_TOP) = True
  P1ResultAvailable(P1_RESULT_SIGMA2_TOP) = True
  P1ResultAvailable(P1_RESULT_TAU12_TOP) = True
  P1ResultAvailable(P1_RESULT_SIGMA1_BOTTOM) = True
  P1ResultAvailable(P1_RESULT_SIGMA2_BOTTOM) = True
  P1ResultAvailable(P1_RESULT_TAU12_BOTTOM) = True
  P1ResultAvailable(P1_RESULT_PRINCIPAL_ANGLE_TOP) = True
  P1ResultAvailable(P1_RESULT_VON_MISES_TOP) = True
  P1ResultAvailable(P1_RESULT_PRINCIPAL_ANGLE_BOTTOM) = True
  P1ResultAvailable(P1_RESULT_VON_MISES_BOTTOM) = True
  P1ResultAvailable(P1_RESULT_SIGMA_Z) = True
  P1ResultAvailable(P1_RESULT_SIGMA_MAX_3D) = True
  P1ResultAvailable(P1_RESULT_SIGMA_MID_3D) = True
  P1ResultAvailable(P1_RESULT_SIGMA_MIN_3D) = True
  BuildGaussToNodeExtrapolationMatrix extrapolation
  For elementId = 0 To NumberOfElement - 1
    If Not P3IsElementActive(elementId) Or Elem(elementId).IsJoint Then GoTo NextRecoveryElement
    RecoverGaussResult elementId
    RecoverElementAverage elementId
    elementWeight = P1ElementArea(Elem(elementId))
    If elementWeight <= 1E-30 Then Err.Raise vbObjectError + 3035, "FEM.BuildP1Results", "要素面積が0または小さすぎます。要素=" & CStr(elementId + 1)
    For resultId = 0 To P1_RESULT_COUNT - 1
      If P1ResultAvailable(resultId) Then
        For gaussNo = 0 To P1_GAUSS_COUNT - 1
          gaussValues(gaussNo) = GaussResult(elementId, gaussNo, resultId)
        Next gaussNo
        ExtrapolateElementResult gaussValues, cornerValues, extrapolation
        For localNode = 0 To 3
          ElementNodeResult(elementId, localNode, resultId) = cornerValues(localNode)
        Next localNode
        ElementNodeResult(elementId, 4, resultId) = 0.5 * (cornerValues(0) + cornerValues(1))
        ElementNodeResult(elementId, 5, resultId) = 0.5 * (cornerValues(1) + cornerValues(2))
        ElementNodeResult(elementId, 6, resultId) = 0.5 * (cornerValues(2) + cornerValues(3))
        ElementNodeResult(elementId, 7, resultId) = 0.5 * (cornerValues(3) + cornerValues(0))
        For localNode = 0 To 7
          nodeId = Elem(elementId).node(localNode)
          AccumulateNodalResult nodeId, resultId, ElementNodeResult(elementId, localNode, resultId), elementWeight
          If resultId = 0 Then
            NodalX(nodeId) = NodalX(nodeId) + elementWeight * Elem(elementId).x(localNode)
            NodalY(nodeId) = NodalY(nodeId) + elementWeight * Elem(elementId).y(localNode)
          End If
        Next localNode
      End If
    Next resultId
    For globalId = 0 To 2
      For gaussNo = 0 To P1_GAUSS_COUNT - 1
        gaussValues(gaussNo) = GaussGlobalResult(elementId, gaussNo, globalId)
      Next gaussNo
      ExtrapolateElementResult gaussValues, cornerValues, extrapolation
      For localNode = 0 To 3
        nodeId = Elem(elementId).node(localNode)
        NodalGlobalResult(nodeId, globalId) = NodalGlobalResult(nodeId, globalId) + elementWeight * cornerValues(localNode)
      Next localNode
      For localNode = 4 To 7
        If localNode = 4 Then midValue = 0.5 * (cornerValues(0) + cornerValues(1))
        If localNode = 5 Then midValue = 0.5 * (cornerValues(1) + cornerValues(2))
        If localNode = 6 Then midValue = 0.5 * (cornerValues(2) + cornerValues(3))
        If localNode = 7 Then midValue = 0.5 * (cornerValues(3) + cornerValues(0))
        nodeId = Elem(elementId).node(localNode)
        NodalGlobalResult(nodeId, globalId) = NodalGlobalResult(nodeId, globalId) + elementWeight * midValue
      Next localNode
    Next globalId
NextRecoveryElement:
  Next elementId
  FinalizeNodalResult
  ' Keep coordinates for nodes currently supported only by inactive or interface elements.
  For elementId = 0 To NumberOfElement - 1
    For localNode = 0 To 7
      nodeId = Elem(elementId).node(localNode)
      If nodeId >= 0 And nodeId < NumberOfFreeNode Then
        If NodalResultWeight(nodeId) = 0# Then
          NodalX(nodeId) = Elem(elementId).x(localNode)
          NodalY(nodeId) = Elem(elementId).y(localNode)
        End If
      End If
    Next localNode
  Next elementId
  For nodeId = 0 To NumberOfFreeNode - 1
    DisplacementMagnitude(nodeId) = Sqr(TDisp(2 * nodeId) * TDisp(2 * nodeId) + TDisp(2 * nodeId + 1) * TDisp(2 * nodeId + 1))
  Next nodeId
  P1ResultReady = True
  P1ResultModelRevision = ModelRevision
End Sub

Public Function P1ResultLabel(ByVal resultId As Long) As String
  Select Case resultId
    Case P1_RESULT_N1_LOCAL: P1ResultLabel = "N1_Local"
    Case P1_RESULT_N2_LOCAL: P1ResultLabel = "N2_Local"
    Case P1_RESULT_N12_LOCAL: P1ResultLabel = "N12_Local"
    Case P1_RESULT_M1_LOCAL: P1ResultLabel = "M1_Local (unsupported)"
    Case P1_RESULT_M2_LOCAL: P1ResultLabel = "M2_Local (unsupported)"
    Case P1_RESULT_M12_LOCAL: P1ResultLabel = "M12_Local (unsupported)"
    Case P1_RESULT_Q1_LOCAL: P1ResultLabel = "Q1_Local (unsupported)"
    Case P1_RESULT_Q2_LOCAL: P1ResultLabel = "Q2_Local (unsupported)"
    Case P1_RESULT_SIGMA1_TOP: P1ResultLabel = "sigma1_Top_Local"
    Case P1_RESULT_SIGMA2_TOP: P1ResultLabel = "sigma2_Top_Local"
    Case P1_RESULT_TAU12_TOP: P1ResultLabel = "tau12_Top_Local"
    Case P1_RESULT_SIGMA1_BOTTOM: P1ResultLabel = "sigma1_Bottom_Local"
    Case P1_RESULT_SIGMA2_BOTTOM: P1ResultLabel = "sigma2_Bottom_Local"
    Case P1_RESULT_TAU12_BOTTOM: P1ResultLabel = "tau12_Bottom_Local"
    Case P1_RESULT_PRINCIPAL_ANGLE_TOP: P1ResultLabel = "PrincipalAngle_Top_Local_deg"
    Case P1_RESULT_VON_MISES_TOP: P1ResultLabel = "vonMises_Top_Local"
    Case P1_RESULT_PRINCIPAL_ANGLE_BOTTOM: P1ResultLabel = "PrincipalAngle_Bottom_Local_deg"
    Case P1_RESULT_VON_MISES_BOTTOM: P1ResultLabel = "vonMises_Bottom_Local"
    Case P1_RESULT_SIGMA_Z: P1ResultLabel = "sigmaZ_OutOfPlane"
    Case P1_RESULT_SIGMA_MAX_3D: P1ResultLabel = "sigmaMax_3D"
    Case P1_RESULT_SIGMA_MID_3D: P1ResultLabel = "sigmaMid_3D"
    Case P1_RESULT_SIGMA_MIN_3D: P1ResultLabel = "sigmaMin_3D"
    Case Else: P1ResultLabel = "Unknown"
  End Select
End Function

Public Function P1ResultUnit(ByVal resultId As Long) As String
  Select Case resultId
    Case P1_RESULT_N1_LOCAL, P1_RESULT_N2_LOCAL, P1_RESULT_N12_LOCAL: P1ResultUnit = "force/length"
    Case P1_RESULT_PRINCIPAL_ANGLE_TOP, P1_RESULT_PRINCIPAL_ANGLE_BOTTOM: P1ResultUnit = "deg"
    Case P1_RESULT_M1_LOCAL, P1_RESULT_M2_LOCAL, P1_RESULT_M12_LOCAL: P1ResultUnit = "unsupported"
    Case P1_RESULT_Q1_LOCAL, P1_RESULT_Q2_LOCAL: P1ResultUnit = "unsupported"
    Case Else: P1ResultUnit = "force/area"
  End Select
End Function

Private Function GetP1ResultSheet() As Worksheet
  Set GetP1ResultSheet = GetIntegratedResultSheet()
End Function

Private Sub ClearP1ResultSheet(ByRef ws As Worksheet)
  Dim lastRow As Long, lastColumn As Long, startRow As Long
  startRow = FEM_INTEGRATED_P1_START_ROW
  If P1LastOutputLastRow > 0 Then
    lastRow = P1LastOutputLastRow
  Else
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
  End If
  lastColumn = P1_RESULT_COUNT + 14
  If lastRow < startRow Then lastRow = startRow
  If lastColumn > P1_RESULT_COUNT + 14 Then lastColumn = P1_RESULT_COUNT + 14
  ws.range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastColumn)).ClearContents
  P1LastOutputLastRow = 0
End Sub

Public Sub InvalidateP1Output()
  Dim ws As Worksheet
  Set ws = GetIntegratedResultSheet()
  ClearP1ResultSheet ws
  P1ResultReady = False
  P1ResultModelRevision = 0
End Sub

Public Sub SaveP1Results()
  Dim ws As Worksheet
  Dim sectionStart As Long
  If Not P1ResultReady Then Err.Raise vbObjectError + 3036, "FEM.SaveP1Results", "P1結果が未作成です。"
  If P1ResultModelRevision <> ModelRevision Then Err.Raise vbObjectError + 3037, "FEM.SaveP1Results", "P1結果のモデル版が現在のモデルと一致しません。"
  Set ws = GetP1ResultSheet()
  ClearP1ResultSheet ws
  sectionStart = FEM_INTEGRATED_P1_START_ROW
  ws.Cells(sectionStart, 1).value2 = "P1詳細は結果節点・結果要素。Gauss状態はブック横 *_out"
  ws.Cells(sectionStart + 1, 1).value2 = "Elements"
  ws.Cells(sectionStart + 1, 2).value2 = NumberOfElement
  ws.Cells(sectionStart + 2, 1).value2 = "Nodes"
  ws.Cells(sectionStart + 2, 2).value2 = NumberOfFreeNode
  ws.Cells(sectionStart + 3, 1).value2 = "P2塑性点は本シート AN列（適応メッシュ用）"
  P1LastOutputLastRow = sectionStart + 3
End Sub

Private Function GetP2MaterialResultSheet() As Worksheet
  Set GetP2MaterialResultSheet = GetIntegratedResultSheet()
End Function

Private Sub ClearP2MaterialResultSheet(ByRef ws As Worksheet)
  Dim lastRow As Long, startRow As Long, startColumn As Long
  startRow = FEM_INTEGRATED_P2_RESULT_START_ROW
  startColumn = FEM_INTEGRATED_P2_RESULT_START_COLUMN
  If P2LastOutputLastRow > 0 Then
    lastRow = P2LastOutputLastRow
  Else
    lastRow = ws.Cells(ws.rows.count, startColumn).End(xlUp).row
  End If
  If lastRow < startRow Then lastRow = startRow
  If lastRow > 1000000 Then lastRow = 1000000
  ws.range(ws.Cells(startRow, startColumn), ws.Cells(lastRow, startColumn + 13)).ClearContents
  P2LastOutputLastRow = 0
End Sub

Public Sub InvalidateP2Output()
  Dim ws As Worksheet
  Set ws = GetIntegratedResultSheet()
  ClearP2MaterialResultSheet ws
End Sub

Public Sub SaveP2MaterialResults()
  Dim ws As Worksheet
  Dim rowCount As Long, rowNo As Long, elementId As Long, gaussNo As Long
  Dim startRow As Long, startColumn As Long
  If Not AnalysisOK Then Err.Raise vbObjectError + 4120, "FEM.SaveP2MaterialResults", "解析が成功していないためP2材料点結果を作成できません。"
  If NumberOfElement <= 0 Then Err.Raise vbObjectError + 4121, "FEM.SaveP2MaterialResults", "P2材料点結果の対象要素がありません。"
  rowCount = NumberOfElement * P1_GAUSS_COUNT
  If rowCount + 8 > 1000000 Then Err.Raise vbObjectError + 4122, "FEM.SaveP2MaterialResults", "P2材料点結果の出力行数がExcelの上限を超えます。"
  Set ws = GetP2MaterialResultSheet()
  startRow = FEM_INTEGRATED_P2_RESULT_START_ROW
  startColumn = FEM_INTEGRATED_P2_RESULT_START_COLUMN
  ClearP2MaterialResultSheet ws
  ws.Cells(startRow, startColumn).value2 = "2DFEM_ P2 Material Point Results"
  ws.Cells(startRow + 1, startColumn).value2 = "Formulation"
  ws.Cells(startRow + 1, startColumn + 1).value2 = P1Formulation
  ws.Cells(startRow + 2, startColumn).value2 = "Yield function"
  ws.Cells(startRow + 2, startColumn + 1).value2 = "f = sigmaMax_3D - sigmaMin_3D - (sigmaMax_3D + sigmaMin_3D) * sin(phi) - 2 * cohesion * cos(phi)"
  ws.Cells(startRow + 3, startColumn).value2 = "Flow rule"
  ws.Cells(startRow + 3, startColumn + 1).value2 = "3D non-associated potential uses DilationAngle psi; sigmaZ is included and psi is independent from phi"
  ws.Cells(startRow + 4, startColumn).value2 = "Tension cutoff"
  ws.Cells(startRow + 4, startColumn + 1).value2 = "Not applied in P2; tensile-domain policy remains explicit for later extension"
  ws.Cells(startRow + 5, startColumn).value2 = "State meaning"
  ws.Cells(startRow + 5, startColumn + 1).value2 = "Plastic strain and multiplier are the latest local update; global commit/rollback remains a P3 responsibility"
  ws.Cells(startRow + 6, startColumn).value2 = "Element"
  ws.Cells(startRow + 6, startColumn + 1).value2 = "GaussPoint"
  ws.Cells(startRow + 6, startColumn + 2).value2 = "PlasticStrainX"
  ws.Cells(startRow + 6, startColumn + 3).value2 = "PlasticStrainY"
  ws.Cells(startRow + 6, startColumn + 4).value2 = "PlasticStrainZ"
  ws.Cells(startRow + 6, startColumn + 5).value2 = "PlasticGammaXY"
  ws.Cells(startRow + 6, startColumn + 6).value2 = "PlasticMultiplier"
  ws.Cells(startRow + 6, startColumn + 7).value2 = "YieldFunction"
  ws.Cells(startRow + 6, startColumn + 8).value2 = "Yielded"
  ws.Cells(startRow + 6, startColumn + 9).value2 = "SigmaMax_3D"
  ws.Cells(startRow + 6, startColumn + 10).value2 = "SigmaMid_3D"
  ws.Cells(startRow + 6, startColumn + 11).value2 = "SigmaMin_3D"
  ws.Cells(startRow + 6, startColumn + 12).value2 = "SigmaZ_OutOfPlane"
  ws.Cells(startRow + 6, startColumn + 13).value2 = "PrincipalAngle_deg"
  Call P6EnsureVariantBuffer(P2OutputRows, P2OutputRowCount, P2OutputColumnCount, rowCount, 14)
  rowNo = 0
  For elementId = 0 To NumberOfElement - 1
    For gaussNo = 0 To P1_GAUSS_COUNT - 1
      rowNo = rowNo + 1
      P2OutputRows(rowNo, 1) = elementId + 1
      P2OutputRows(rowNo, 2) = gaussNo
      P2OutputRows(rowNo, 3) = Elem(elementId).P2PlasticStrain(0, gaussNo)
      P2OutputRows(rowNo, 4) = Elem(elementId).P2PlasticStrain(1, gaussNo)
      P2OutputRows(rowNo, 5) = Elem(elementId).P2PlasticStrain(3, gaussNo)
      P2OutputRows(rowNo, 6) = Elem(elementId).P2PlasticStrain(2, gaussNo)
      P2OutputRows(rowNo, 7) = Elem(elementId).P2PlasticMultiplier(gaussNo)
      P2OutputRows(rowNo, 8) = Elem(elementId).P2YieldFunction(gaussNo)
      P2OutputRows(rowNo, 9) = Elem(elementId).P2Yielded(gaussNo)
      P2OutputRows(rowNo, 10) = Elem(elementId).P2PrincipalStress(0, gaussNo)
      P2OutputRows(rowNo, 11) = Elem(elementId).P2PrincipalStress(1, gaussNo)
      P2OutputRows(rowNo, 12) = Elem(elementId).P2PrincipalStress(2, gaussNo)
      P2OutputRows(rowNo, 13) = Elem(elementId).Stmat(16, gaussNo)
      P2OutputRows(rowNo, 14) = Elem(elementId).Stmat(6, gaussNo)
    Next gaussNo
  Next elementId
  If rowCount > 0 Then ws.range(ws.Cells(startRow + 7, startColumn), ws.Cells(startRow + 6 + rowCount, startColumn + 13)).value2 = P2OutputRows
  P2LastOutputLastRow = startRow + 6 + rowCount
  If P6ReadSetting("OUTPUT_AUTOFIT", 0#) <> 0# Then
    ws.range(ws.columns(startColumn), ws.columns(startColumn + 13)).EntireColumn.AutoFit
  End If
End Sub

Public Sub SaveP0Progress(ByVal stageName As String)
  Dim ws As Worksheet
  If left$(stageName, 8) = "p3_retry" Or InStr(1, stageName, "increment", vbTextCompare) > 0 Then Exit Sub
  Set ws = GetDiagnosticSheet()
  ws.Cells(1, 1).value2 = "2DFEM_P0進捗"
  ws.Cells(2, 1).value2 = "Stage"
  ws.Cells(2, 2).value2 = stageName
  ws.Cells(3, 1).value2 = "ResultStatus"
  ws.Cells(3, 2).value2 = ResultStatus
  ws.Cells(4, 1).value2 = "CurrentIncrement"
  ws.Cells(4, 2).value2 = CurrentIncrement
  ws.Cells(5, 1).value2 = "CurrentIteration"
  ws.Cells(5, 2).value2 = CurrentIteration
  ws.Cells(6, 1).value2 = "FailureElement"
  ws.Cells(6, 2).value2 = FailureElement
  If SuppressUserMessages Then
    If stageName = "start" Then ThisWorkbook.Save
  End If
End Sub
