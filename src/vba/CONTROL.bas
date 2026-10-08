Option Explicit
Public Type P1_ViewState
    ProjectionMatrix(0 To 2, 0 To 2) As Double
    originX As Double
    originY As Double
    drawScale As Double
    DeformedScale As Double
    AngleX As Double
    AngleY As Double
    AngleZ As Double
    AutoFit As Boolean
    ShowNodeLabels As Boolean
    ShowElementLabels As Boolean
    ShowOriginal As Boolean
    ShowDeformed As Boolean
    ShowContour As Boolean
    resultId As Long
    scaleMode As String
    DetailMode As String
    ResultSource As String
End Type
Private Const P1_VIEW_SHAPE_LIMIT As Long = 5000
Private Const P1_VIEW_SAFE_SHAPE_LIMIT As Long = 4500
Private Const P1_CONTOUR_BIN_COUNT As Long = 32
Private Const RESULT_NOT_RUN As String = "NOT_RUN"
Private Const RESULT_PASS As String = "PASS"
Private Const RESULT_INPUT_ERROR As String = "INPUT_ERROR"
Private Const RESULT_MATERIAL_ERROR As String = "MATERIAL_ERROR"
Private Const RESULT_NONCONVERGED As String = "NONCONVERGED"
Private Const RESULT_GLOBAL_SINGULAR As String = "GLOBAL_SINGULAR"
Private Const RESULT_CAPACITY_ERROR As String = "CAPACITY_ERROR"
Private Const RESULT_RUNTIME_ERROR As String = "RUNTIME_ERROR"
Private Const P1_RESULT_COUNT As Long = 22
Private Const P1_RESULT_VON_MISES_TOP As Long = 15
Public start_time As Double
Public now_time As Double
Public calc_time As Long
Public start_day As Long
Public now_day As Long
Public P1ViewerLastError As String
Private applicationStateSaved As Boolean
Private savedScreenUpdating As Boolean
Private savedEnableEvents As Boolean
Private savedCalculation As Variant
Private savedCalculationValid As Boolean
Private savedEnableCancelKey As Long
Private savedCancelKeyValid As Boolean

Private FEMSettingCacheReady As Boolean
Private FEMSettingCacheCount As Long
Private FEMSettingCacheKeys() As String
Private FEMSettingCacheValues() As Variant

Private Sub SaveApplicationState()
    savedScreenUpdating = Application.ScreenUpdating
    savedEnableEvents = Application.EnableEvents
    savedCalculationValid = False
    On Error Resume Next
    savedCalculation = Application.Calculation
    savedCalculationValid = (Err.Number = 0 And Not IsError(savedCalculation) And IsNumeric(savedCalculation))
    Err.Clear
    On Error Resume Next
    savedEnableCancelKey = Application.EnableCancelKey
    savedCancelKeyValid = (Err.Number = 0)
    Err.Clear
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo 0
    applicationStateSaved = True
End Sub

Private Sub RestoreApplicationState()
    If applicationStateSaved Then
        Application.ScreenUpdating = savedScreenUpdating
        Application.EnableEvents = savedEnableEvents
        If savedCalculationValid Then
            On Error Resume Next
            Application.Calculation = CLng(savedCalculation)
            On Error GoTo 0
        End If
        If savedCancelKeyValid Then
            On Error Resume Next
            Application.EnableCancelKey = savedEnableCancelKey
            On Error GoTo 0
        End If
        applicationStateSaved = False
        savedCalculationValid = False
        savedCancelKeyValid = False
    End If
End Sub

Public Sub FEMInvalidateSettingCache()
    FEMSettingCacheReady = False
    FEMSettingCacheCount = 0
    Erase FEMSettingCacheKeys
    Erase FEMSettingCacheValues
    FEMInvalidateP6SettingCache
End Sub

Private Sub FEMEnsureSettingCache()
    Dim ws As Worksheet, lastRow As Long, rowNo As Long
    Dim keyValue As Variant
    If FEMSettingCacheReady Then Exit Sub
    FEMSettingCacheReady = True
    FEMSettingCacheCount = 0
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("設定")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    FEMUiCalculateSettings
    lastRow = ws.Cells(ws.rows.count, 5).End(xlUp).row
    If lastRow < 1 Then Exit Sub
    ReDim FEMSettingCacheKeys(0 To lastRow - 1)
    ReDim FEMSettingCacheValues(0 To lastRow - 1)
    For rowNo = 1 To lastRow
        keyValue = ws.Cells(rowNo, 5).value2
        If Not IsError(keyValue) Then
            If Len(Trim$(CStr(keyValue))) > 0 Then
                FEMSettingCacheKeys(FEMSettingCacheCount) = UCase$(Trim$(CStr(keyValue)))
                FEMSettingCacheValues(FEMSettingCacheCount) = ws.Cells(rowNo, 3).value2
                FEMSettingCacheCount = FEMSettingCacheCount + 1
            End If
        End If
    Next rowNo
End Sub

Private Function FEMFindSettingValue(ByVal keyName As String, ByRef found As Boolean) As Variant
    Dim i As Long, normalizedKey As String
    found = False
    FEMEnsureSettingCache
    normalizedKey = UCase$(Trim$(keyName))
    For i = 0 To FEMSettingCacheCount - 1
        If FEMSettingCacheKeys(i) = normalizedKey Then
            FEMFindSettingValue = FEMSettingCacheValues(i)
            found = True
            Exit Function
        End If
    Next i
End Function

Public Function FEMReadSetting(ByVal keyName As String, ByVal defaultValue As Double) As Double
    Dim rawValue As Variant, found As Boolean
    rawValue = FEMFindSettingValue(keyName, found)
    If found Then
        If Not IsError(rawValue) Then
            If IsNumeric(rawValue) Then
                FEMReadSetting = CDbl(rawValue)
                Exit Function
            End If
        End If
    End If
    FEMReadSetting = defaultValue
End Function

Public Function FEMReadTextSetting(ByVal keyName As String, ByVal defaultValue As String) As String
    Dim rawValue As Variant, found As Boolean
    rawValue = FEMFindSettingValue(keyName, found)
    If found Then
        If Not IsError(rawValue) Then
            If Len(Trim$(CStr(rawValue))) > 0 Then
                FEMReadTextSetting = CStr(rawValue)
                Exit Function
            End If
        End If
    End If
    FEMReadTextSetting = defaultValue
End Function

Private Function FEMFindSettingRow(ByVal keyName As String) As Long
    Dim ws As Worksheet
    Dim lastRow As Long, rowNo As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("設定")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.rows.count, 5).End(xlUp).row
    For rowNo = 1 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(rowNo, 5).value2))) = UCase$(Trim$(keyName)) Then
            FEMFindSettingRow = rowNo
            Exit Function
        End If
    Next rowNo
End Function

Public Sub FEMWriteNumericSetting(ByVal keyName As String, ByVal value As Double)
    Dim ws As Worksheet
    Dim rowNo As Long, lastRow As Long
    Dim keyUpper As String
    keyUpper = UCase$(Trim$(keyName))
    If left$(keyUpper, 18) = "MESH_QUALITY_LAST_" Or left$(keyUpper, 20) = "MESH_QUALITY_REPAIR_" Then Exit Sub
    If FEMUiTryWriteSetting(keyName, value) Then
        FEMInvalidateSettingCache
        Exit Sub
    End If
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("設定")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    rowNo = FEMFindSettingRow(keyName)
    If rowNo = 0 Then
        lastRow = ws.Cells(ws.rows.count, 5).End(xlUp).row
        rowNo = lastRow + 1
        ws.Cells(rowNo, 1).value2 = "追加"
        ws.Cells(rowNo, 2).value2 = keyName
        ws.Cells(rowNo, 4).value2 = "自動記録"
        ws.Cells(rowNo, 5).value2 = keyName
    End If
    ws.Cells(rowNo, 3).value2 = value
    FEMInvalidateSettingCache
End Sub

Public Sub FEMWriteTextSetting(ByVal keyName As String, ByVal value As String)
    Dim ws As Worksheet
    Dim rowNo As Long, lastRow As Long
    Dim keyUpper As String
    keyUpper = UCase$(Trim$(keyName))
    If left$(keyUpper, 18) = "MESH_QUALITY_LAST_" Or left$(keyUpper, 20) = "MESH_QUALITY_REPAIR_" Then Exit Sub
    If FEMUiTryWriteSetting(keyName, value) Then
        FEMInvalidateSettingCache
        Exit Sub
    End If
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("設定")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    rowNo = FEMFindSettingRow(keyName)
    If rowNo = 0 Then
        lastRow = ws.Cells(ws.rows.count, 5).End(xlUp).row
        rowNo = lastRow + 1
        ws.Cells(rowNo, 1).value2 = "メッシュ品質"
        ws.Cells(rowNo, 2).value2 = keyName
        ws.Cells(rowNo, 4).value2 = "自動記録"
        ws.Cells(rowNo, 5).value2 = keyName
    End If
    ws.Cells(rowNo, 3).value2 = value
    FEMInvalidateSettingCache
End Sub

Private Function ErrorStatus(ByVal errorNumber As Long) As String
    Dim localNumber As Long
    localNumber = errorNumber - vbObjectError
    Select Case localNumber
        Case 3001 To 3008
            ErrorStatus = RESULT_INPUT_ERROR
        Case 3023 To 3029
            ErrorStatus = RESULT_INPUT_ERROR
        Case 3401 To 3419
            ErrorStatus = RESULT_CAPACITY_ERROR
        Case 3450 To 3499
            ErrorStatus = RESULT_INPUT_ERROR
        Case 3501 To 3519
            ErrorStatus = RESULT_CAPACITY_ERROR
        Case 3101 To 3102
            ErrorStatus = RESULT_MATERIAL_ERROR
        Case 3301 To 3306
            ErrorStatus = RESULT_GLOBAL_SINGULAR
        Case 3201 To 3202
            ErrorStatus = RESULT_NONCONVERGED
        Case Else
            ErrorStatus = RESULT_RUNTIME_ERROR
    End Select
End Function

Private Function TimerElapsed(ByVal startedAt As Double) As Double
    TimerElapsed = Timer - startedAt
    If TimerElapsed < 0# Then TimerElapsed = TimerElapsed + 86400#
End Function

Public Sub P0_RunAnalysis()
    SetP0SilentMode True
    FEMInvalidateSettingCache
    ボタン5_Click
    SetP0SilentMode False
End Sub

Private Sub CreateMeshFromGeometryDefinition()
    MeshPerfEnter "CONTROL.CreateMeshFromGeometryDefinition"
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim stateWasSaved As Boolean
    Dim currentStage As String

    On Error GoTo ErrorHandler
    currentStage = "開始"
    If AnalysisRunning Then
        If Not SuppressUserMessages Then MsgBox "解析処理中です。", vbExclamation
        MeshPerfLeave "CONTROL.CreateMeshFromGeometryDefinition": Exit Sub
    End If

    SaveApplicationState
    stateWasSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo ErrorHandler

    ' ①の座標点メッシュ生成とは別に、図形定義を入力として
    ' デローニ分割・細分化・四角形化を実行する。
    ' 図形細分化シートやMESH_POINT*はこの経路では参照しない。
    currentStage = "出力無効化"
    ControlInvalidateP1Output
    ControlInvalidateP2Output
    currentStage = "解析入力キャッシュ無効化"
    FEMInvalidateInputCache
    FEMInvalidateSettingCache
    currentStage = "デローニ初期化"
    CER
    currentStage = "図形定義作業領域確保"
    P5PrepareGeometryDefinitionWorkspace
    currentStage = "図形定義読込"
    PINPUT NEX, NIN, ibex, ibin, ibno, NOB, NIB, px, py, delx, xlok
    currentStage = "境界トポロジー出力"
    TROUTPUT NEX, NIN, ibex, ibin, ibno, NOB, NIB, px, py, delx, xlok
    currentStage = "三角形モデル生成"
    TRMODEL NEX, NIN, ibex, ibin, ibno, NOB, NIB, mindex, node, px, py, NELM, mtj, jac, idm, ifix, delx, xlok
    currentStage = "四角形化要素生成"
    QUEXELM NELM, mtj, jac, kv, idm, map, px, py
    currentStage = "要素情報生成"
    QUISOGEN node, NELM, mtj, jac, px, py, id, mmtj, ifix, idm
    currentStage = "Q8節点入力"
    QUDATASQINPUT NELM, mtj, mmtj
    currentStage = "メッシュ検査"
    QUCHECKDATA node, NNEX, px, py, mpx, mpy, mibex, mibno, ifix
    currentStage = "Q8結果出力"
    QUDATA node, NELM, mmtj, px, py, ifix, idm
    currentStage = "境界条件"
    MeshApplySheetBoundaryConditions
    currentStage = "現行シート検証"
    P5EvaluateCurrentSheet "要素作成②-図形定義-デローニ"
    currentStage = "デローニと地形格子の比較"
    P5SelectBestGeneratedMesh
    If Not P5QualityMeetsSettings() And Not SuppressUserMessages Then
        MsgBox "要素作成②はメッシュ品質の目標値を下回っています。" & vbCrLf & P5QualitySummary() & vbCrLf & MeshAdaptLastSummary & vbCrLf & "図を確認し、格子間隔を小さくするか図形定義の点配置を見直してください。", vbExclamation
    End If
    currentStage = "接合要素"
    MeshInsertJointsIfNeeded

    P5ReleaseGeometryDefinitionWorkspace
    If stateWasSaved Then RestoreApplicationState
    MeshPerfLeave "CONTROL.CreateMeshFromGeometryDefinition": Exit Sub
ErrorHandler:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    P5ReleaseGeometryDefinitionWorkspace
    On Error GoTo 0
    If stateWasSaved Then RestoreApplicationState
    If Not SuppressUserMessages Then
        MsgBox "要素作成②（図形定義→デローニ分割）に失敗しました [" & CStr(errorNumber) & "] " & currentStage & "：" & errorDescription, vbCritical
    End If
    MeshPerfLeave "CONTROL.CreateMeshFromGeometryDefinition"
End Sub

Sub ボタン19_Click()
    CreateMeshFromGeometryDefinition
End Sub
Private Function MeshPointEqual(ByVal x1 As Double, ByVal y1 As Double, _
                                ByVal x2 As Double, ByVal y2 As Double, _
                                ByVal tolerance As Double) As Boolean
    MeshPointEqual = (Abs(x1 - x2) <= tolerance And Abs(y1 - y2) <= tolerance)
End Function

Private Function MeshUniquePointCount(ByRef pointX() As Double, ByRef pointY() As Double, _
                                      ByVal pointCount As Long, ByVal tolerance As Double, _
                                      ByRef uniqueX() As Double, ByRef uniqueY() As Double) As Long
    Dim i As Long, j As Long
    Dim isNew As Boolean

    MeshUniquePointCount = 0
    For i = 1 To pointCount
        isNew = True
        For j = 1 To MeshUniquePointCount
            If MeshPointEqual(pointX(i), pointY(i), uniqueX(j), uniqueY(j), tolerance) Then
                isNew = False
                Exit For
            End If
        Next j
        If isNew Then
            MeshUniquePointCount = MeshUniquePointCount + 1
            uniqueX(MeshUniquePointCount) = pointX(i)
            uniqueY(MeshUniquePointCount) = pointY(i)
        End If
    Next i
End Function

Private Function MeshCoordinateTolerance(ByRef pointX() As Double, ByRef pointY() As Double, _
                                         ByVal pointCount As Long) As Double
    Dim i As Long
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim modelScale As Double

    minX = pointX(1): maxX = pointX(1)
    minY = pointY(1): maxY = pointY(1)
    For i = 2 To pointCount
        If pointX(i) < minX Then minX = pointX(i)
        If pointX(i) > maxX Then maxX = pointX(i)
        If pointY(i) < minY Then minY = pointY(i)
        If pointY(i) > maxY Then maxY = pointY(i)
    Next i
    modelScale = maxX - minX
    If modelScale < maxY - minY Then modelScale = maxY - minY
    If modelScale < 1# Then modelScale = 1#
    MeshCoordinateTolerance = modelScale * 0.0000000001
End Function

Private Function MeshSignedArea2(ByRef pointX() As Double, ByRef pointY() As Double, _
                                 ByVal pointCount As Long) As Double
    Dim i As Long, nextIndex As Long
    For i = 1 To pointCount
        nextIndex = i + 1
        If nextIndex > pointCount Then nextIndex = 1
        MeshSignedArea2 = MeshSignedArea2 + pointX(i) * pointY(nextIndex) - pointX(nextIndex) * pointY(i)
    Next i
End Function

Private Sub MeshNormalizeCounterClockwise(ByRef pointX() As Double, ByRef pointY() As Double, _
                                          ByVal pointCount As Long, ByVal tolerance As Double)
    Dim swapValue As Double
    Dim area2 As Double

    area2 = MeshSignedArea2(pointX, pointY, pointCount)
    If Abs(area2) <= tolerance Then
        Err.Raise vbObjectError + 3461, "CONTROL.MeshNormalizeCounterClockwise", _
                  "入力点が一直線上、または面積が小さすぎます。"
    End If
    If area2 < 0# Then
        swapValue = pointX(2): pointX(2) = pointX(pointCount): pointX(pointCount) = swapValue
        swapValue = pointY(2): pointY(2) = pointY(pointCount): pointY(pointCount) = swapValue
    End If
End Sub

Private Sub MeshWriteQ8Element(ByRef elementCount As Long, _
                               ByRef elementRows() As Variant, _
                               ByVal node1 As Long, ByVal node2 As Long, ByVal node3 As Long, ByVal node4 As Long, _
                               ByVal node5 As Long, ByVal node6 As Long, ByVal node7 As Long, ByVal node8 As Long, _
                               ByVal materialNumber As Long)
    ' 生成中はグローバルmtjがまだ確保されていないため、
    ' ローカルの出力配列容量だけを確認する。
    If elementCount + 1 > UBound(elementRows, 1) Then
        Err.Raise vbObjectError + 3469, "CONTROL.MeshWriteQ8Element", _
                  "Q8生成要素のローカル配列容量を超えました。分割数を下げてください。"
    End If
    elementCount = elementCount + 1
    elementRows(elementCount, 1) = elementCount
    elementRows(elementCount, 2) = node1
    elementRows(elementCount, 3) = node2
    elementRows(elementCount, 4) = node3
    elementRows(elementCount, 5) = node4
    elementRows(elementCount, 6) = node5
    elementRows(elementCount, 7) = node6
    elementRows(elementCount, 8) = node7
    elementRows(elementCount, 9) = node8
    elementRows(elementCount, 10) = materialNumber
End Sub

Private Sub MeshAppendNode(ByRef nodeCount As Long, ByRef nodeX() As Double, ByRef nodeY() As Double, _
                           ByVal coordinateX As Double, ByVal coordinateY As Double)
    nodeCount = nodeCount + 1
    If nodeCount > UBound(nodeX) Or nodeCount > UBound(nodeY) Then
        Err.Raise vbObjectError + 3467, "CONTROL.MeshAppendNode", "Q8メッシュ節点の固定配列容量を超えました。分割数を下げてください。"
    End If
    nodeX(nodeCount) = coordinateX
    nodeY(nodeCount) = coordinateY
End Sub

Private Sub MeshCommitGeneratedMesh(ByVal nodeCount As Long, ByVal elementCount As Long, _
                                    ByRef nodeX() As Double, ByRef nodeY() As Double, _
                                    ByRef elementRows() As Variant)
    Dim wsNode As Worksheet, wsElement As Worksheet
    Dim oldNodeCount As Long, oldElementCount As Long
    Dim nodeOutput() As Variant
    Dim condX() As Long, condY() As Long
    Dim i As Long

    If nodeCount < 1 Or elementCount < 1 Then
        Err.Raise vbObjectError + 3468, "CONTROL.MeshCommitGeneratedMesh", "生成されたQ8メッシュが空です。"
    End If
    ReDim condX(1 To nodeCount)
    ReDim condY(1 To nodeCount)
    MeshApplyBoundaryByGeometry nodeX, nodeY, nodeCount, condX, condY

    Set wsNode = ThisWorkbook.Worksheets("節点データ")
    Set wsElement = ThisWorkbook.Worksheets("要素データ")
    oldNodeCount = wsNode.Cells(wsNode.rows.count, 1).End(xlUp).row - 1
    oldElementCount = wsElement.Cells(wsElement.rows.count, 1).End(xlUp).row - 1
    If oldNodeCount > 0 Then wsNode.range(wsNode.Cells(2, 1), wsNode.Cells(oldNodeCount + 1, 14)).ClearContents
    If oldElementCount > 0 Then wsElement.range(wsElement.Cells(2, 1), wsElement.Cells(oldElementCount + 1, 10)).ClearContents

    ReDim nodeOutput(1 To nodeCount, 1 To 9)
    For i = 1 To nodeCount
        nodeOutput(i, 1) = i
        nodeOutput(i, 2) = nodeX(i)
        nodeOutput(i, 3) = nodeY(i)
        nodeOutput(i, 4) = condX(i)
        nodeOutput(i, 5) = condY(i)
        nodeOutput(i, 6) = 0
        nodeOutput(i, 7) = 0
        nodeOutput(i, 8) = 0
        nodeOutput(i, 9) = 0
    Next i
    wsNode.range(wsNode.Cells(2, 1), wsNode.Cells(nodeCount + 1, 9)).value2 = nodeOutput
    wsElement.range(wsElement.Cells(2, 1), wsElement.Cells(elementCount + 1, 10)).value2 = elementRows
    FEMInvalidateInputCache
End Sub

Private Function MeshBoundaryCode(ByVal keyName As String, ByVal defaultText As String) As String
    Dim codeText As String
    codeText = UCase$(Trim$(FEMReadTextSetting(keyName, defaultText)))
    If codeText <> "NONE" And codeText <> "ROLLER" And codeText <> "PINNED" And codeText <> "AUTO" Then
        codeText = UCase$(Trim$(defaultText))
    End If
    MeshBoundaryCode = codeText
End Function

Private Sub MeshApplySideConstraint(ByVal sideCode As String, ByVal fixX As Boolean, ByVal onSide As Boolean, _
                                    ByRef condX As Long, ByRef condY As Long)
    If Not onSide Then Exit Sub
    If sideCode = "ROLLER" Then
        If fixX Then condX = 1 Else condY = 1
    ElseIf sideCode = "PINNED" Then
        condX = 1
        condY = 1
    End If
End Sub

Private Sub MeshApplyBoundaryByGeometry(ByRef nodeX() As Double, ByRef nodeY() As Double, ByVal nodeCount As Long, _
                                        ByRef condX() As Long, ByRef condY() As Long)
    Dim i As Long, pinIndex As Long
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double, yTol As Double, xTol As Double
    Dim bottomCode As String, leftCode As String, rightCode As String, topCode As String, pinCode As String
    Dim pinX As Double

    If nodeCount < 1 Then Exit Sub
    minX = nodeX(1): maxX = nodeX(1): minY = nodeY(1): maxY = nodeY(1)
    For i = 2 To nodeCount
        If nodeX(i) < minX Then minX = nodeX(i)
        If nodeX(i) > maxX Then maxX = nodeX(i)
        If nodeY(i) < minY Then minY = nodeY(i)
        If nodeY(i) > maxY Then maxY = nodeY(i)
    Next i
    xTol = Abs(maxX - minX) * 0.000000001
    yTol = Abs(maxY - minY) * 0.000000001
    If xTol < 0.000000001 Then xTol = 0.000000001
    If yTol < 0.000000001 Then yTol = 0.000000001

    bottomCode = MeshBoundaryCode("MESH_BC_BOTTOM", "ROLLER")
    leftCode = MeshBoundaryCode("MESH_BC_LEFT", "ROLLER")
    rightCode = MeshBoundaryCode("MESH_BC_RIGHT", "NONE")
    topCode = MeshBoundaryCode("MESH_BC_TOP", "NONE")
    pinCode = MeshBoundaryCode("MESH_BC_PIN_CORNER", "NONE")

    pinIndex = 0
    pinX = 1E+30
    For i = 1 To nodeCount
        condX(i) = 0
        condY(i) = 0
        MeshApplySideConstraint bottomCode, False, Abs(nodeY(i) - minY) <= yTol, condX(i), condY(i)
        MeshApplySideConstraint topCode, False, Abs(nodeY(i) - maxY) <= yTol, condX(i), condY(i)
        MeshApplySideConstraint leftCode, True, Abs(nodeX(i) - minX) <= xTol, condX(i), condY(i)
        MeshApplySideConstraint rightCode, True, Abs(nodeX(i) - maxX) <= xTol, condX(i), condY(i)
        If pinCode = "AUTO" And Abs(nodeY(i) - minY) <= yTol Then
            If nodeX(i) < pinX - xTol Then
                pinX = nodeX(i)
                pinIndex = i
            End If
        End If
    Next i
    If pinIndex > 0 Then
        condX(pinIndex) = 1
        condY(pinIndex) = 1
    End If
End Sub

Public Sub MeshApplySheetBoundaryConditions()
    MeshPerfEnter "CONTROL.MeshApplySheetBoundaryConditions"
    Dim wsNode As Worksheet
    Dim lastRow As Long, nodeCount As Long, i As Long
    Dim nodeX() As Double, nodeY() As Double
    Dim condX() As Long, condY() As Long
    Dim nodeOutput() As Variant
    Dim rawValues As Variant

    Set wsNode = ThisWorkbook.Worksheets("節点データ")
    lastRow = wsNode.Cells(wsNode.rows.count, 1).End(xlUp).row
    nodeCount = lastRow - 1
    If nodeCount < 1 Then MeshPerfLeave "CONTROL.MeshApplySheetBoundaryConditions": Exit Sub
    rawValues = wsNode.range(wsNode.Cells(2, 1), wsNode.Cells(nodeCount + 1, 3)).value2
    ReDim nodeX(1 To nodeCount)
    ReDim nodeY(1 To nodeCount)
    ReDim condX(1 To nodeCount)
    ReDim condY(1 To nodeCount)
    For i = 1 To nodeCount
        nodeX(i) = CDbl(rawValues(i, 2))
        nodeY(i) = CDbl(rawValues(i, 3))
    Next i
    MeshApplyBoundaryByGeometry nodeX, nodeY, nodeCount, condX, condY
    ReDim nodeOutput(1 To nodeCount, 1 To 2)
    For i = 1 To nodeCount
        nodeOutput(i, 1) = condX(i)
        nodeOutput(i, 2) = condY(i)
    Next i
    wsNode.range(wsNode.Cells(2, 4), wsNode.Cells(nodeCount + 1, 5)).value2 = nodeOutput
    FEMInvalidateInputCache
    MeshPerfLeave "CONTROL.MeshApplySheetBoundaryConditions"
End Sub

Public Function MeshReadMaterialNumber() As Long
    Dim materialNumber As Long
    materialNumber = CLng(FEMReadSetting("MESH_MATERIAL", 1#))
    If materialNumber < 1 Then materialNumber = 1
    MeshReadMaterialNumber = materialNumber
End Function

Private Function MeshQuadSignedArea(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                                    ByVal x3 As Double, ByVal y3 As Double, ByVal x4 As Double, ByVal y4 As Double) As Double
    MeshQuadSignedArea = (x1 * y2 - x2 * y1) + (x2 * y3 - x3 * y2) + (x3 * y4 - x4 * y3) + (x4 * y1 - x1 * y4)
End Function

Private Function MeshSegmentsProperIntersect(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double, _
                                             ByVal cx As Double, ByVal cy As Double, ByVal dx As Double, ByVal dy As Double) As Boolean
    Dim o1 As Double, o2 As Double, o3 As Double, o4 As Double
    o1 = (bx - ax) * (cy - ay) - (by - ay) * (cx - ax)
    o2 = (bx - ax) * (dy - ay) - (by - ay) * (dx - ax)
    o3 = (dx - cx) * (ay - cy) - (dy - cy) * (ax - cx)
    o4 = (dx - cx) * (by - cy) - (dy - cy) * (bx - cx)
    MeshSegmentsProperIntersect = (o1 * o2 < 0#) And (o3 * o4 < 0#)
End Function

Private Sub MeshValidateSourceQuad(ByRef pointX() As Double, ByRef pointY() As Double)
    Dim areaValue As Double
    areaValue = MeshQuadSignedArea(pointX(1), pointY(1), pointX(2), pointY(2), pointX(3), pointY(3), pointX(4), pointY(4))
    If areaValue <= 0# Then
        Err.Raise vbObjectError + 3470, "CONTROL.MeshValidateSourceQuad", "四辺形の面積が正ではありません。点の順序を確認してください。"
    End If
    If MeshSegmentsProperIntersect(pointX(1), pointY(1), pointX(2), pointY(2), pointX(3), pointY(3), pointX(4), pointY(4)) Then
        Err.Raise vbObjectError + 3471, "CONTROL.MeshValidateSourceQuad", "四辺形の対辺が交差しています。"
    End If
    If MeshSegmentsProperIntersect(pointX(2), pointY(2), pointX(3), pointY(3), pointX(4), pointY(4), pointX(1), pointY(1)) Then
        Err.Raise vbObjectError + 3471, "CONTROL.MeshValidateSourceQuad", "四辺形の対辺が交差しています。"
    End If
End Sub

Private Sub MeshCreateQuadrilateral(ByRef pointX() As Double, ByRef pointY() As Double, _
                                    ByVal subdivisionX As Long, ByVal subdivisionY As Long)
    Dim i As Long, j As Long, nodeCount As Long, elementCount As Long
    Dim n1 As Long, n2 As Long, n3 As Long, N4 As Long, n5 As Long, n6 As Long, n7 As Long, n8 As Long
    Dim u As Double, v As Double, xx As Double, yy As Double
    Dim n00 As Double, n10 As Double, n11 As Double, n01 As Double
    Dim expectedNodeCount As Long, expectedElementCount As Long, edgeCapacity As Long
    Dim nodeX() As Double, nodeY() As Double, edgeA() As Long, edgeB() As Long, edgeMid() As Long
    Dim elementRows() As Variant
    Dim edgeCount As Long
    Dim materialNumber As Long
    Dim edgeLookup As Collection
    Dim elementArea As Double

    If subdivisionX < 1 Or subdivisionY < 1 Then
        Err.Raise vbObjectError + 3462, "CONTROL.MeshCreateQuadrilateral", "MESH_NX/MESH_NYは1以上にしてください。"
    End If
    MeshValidateSourceQuad pointX, pointY
    expectedElementCount = subdivisionX * subdivisionY
    expectedNodeCount = (subdivisionX + 1) * (subdivisionY + 1) + subdivisionX * (subdivisionY + 1) + subdivisionY * (subdivisionX + 1)
    edgeCapacity = subdivisionX * (subdivisionY + 1) + subdivisionY * (subdivisionX + 1)
    If expectedNodeCount > 60000 Or expectedElementCount > 60000 Then
        Err.Raise vbObjectError + 3462, "CONTROL.MeshCreateQuadrilateral", _
                  "指定した分割数ではQ8節点が容量を超えます。必要節点=" & CStr(expectedNodeCount) & "。"
    End If
    ReDim nodeX(1 To expectedNodeCount)
    ReDim nodeY(1 To expectedNodeCount)
    ReDim edgeA(1 To edgeCapacity)
    ReDim edgeB(1 To edgeCapacity)
    ReDim edgeMid(1 To edgeCapacity)
    ReDim elementRows(1 To expectedElementCount, 1 To 10)
    Set edgeLookup = New Collection
    materialNumber = MeshReadMaterialNumber()

    For j = 0 To subdivisionY
        v = CDbl(j) / CDbl(subdivisionY)
        For i = 0 To subdivisionX
            u = CDbl(i) / CDbl(subdivisionX)
            n00 = (1# - u) * (1# - v)
            n10 = u * (1# - v)
            n11 = u * v
            n01 = (1# - u) * v
            xx = n00 * pointX(1) + n10 * pointX(2) + n11 * pointX(3) + n01 * pointX(4)
            yy = n00 * pointY(1) + n10 * pointY(2) + n11 * pointY(3) + n01 * pointY(4)
            MeshAppendNode nodeCount, nodeX, nodeY, xx, yy
        Next i
    Next j

    For j = 0 To subdivisionY - 1
        For i = 0 To subdivisionX - 1
            n1 = i + (subdivisionX + 1) * j + 1
            n2 = n1 + 1
            n3 = n2 + (subdivisionX + 1)
            N4 = n1 + (subdivisionX + 1)
            elementArea = MeshQuadSignedArea(nodeX(n1), nodeY(n1), nodeX(n2), nodeY(n2), nodeX(n3), nodeY(n3), nodeX(N4), nodeY(N4))
            If elementArea <= 0# Then
                Err.Raise vbObjectError + 3472, "CONTROL.MeshCreateQuadrilateral", _
                          "生成要素が裏返っています。i=" & CStr(i) & " j=" & CStr(j)
            End If
            n5 = MeshGetEdgeMidpoint(n1, n2, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
            n6 = MeshGetEdgeMidpoint(n2, n3, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
            n7 = MeshGetEdgeMidpoint(n3, N4, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
            n8 = MeshGetEdgeMidpoint(N4, n1, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
            MeshWriteQ8Element elementCount, elementRows, n1, n2, n3, N4, n5, n6, n7, n8, materialNumber
        Next i
    Next j
    If nodeCount <> expectedNodeCount Then
        Err.Raise vbObjectError + 3473, "CONTROL.MeshCreateQuadrilateral", _
                  "Q8節点数が想定と違います。想定=" & CStr(expectedNodeCount) & " 実際=" & CStr(nodeCount)
    End If
    P5ValidateMeshCapacity nodeCount, elementCount
    MeshCommitGeneratedMesh nodeCount, elementCount, nodeX, nodeY, elementRows
End Sub

Private Function MeshGetEdgeMidpoint(ByVal nodeA As Long, ByVal nodeB As Long, _
                                     ByRef nodeX() As Double, ByRef nodeY() As Double, _
                                     ByRef edgeA() As Long, ByRef edgeB() As Long, _
                                     ByRef edgeMid() As Long, ByRef edgeCount As Long, _
                                     ByRef nodeCount As Long, ByVal edgeLookup As Collection) As Long
    Dim lowNode As Long, highNode As Long
    Dim midX As Double, midY As Double
    Dim keyName As String
    Dim existingIndex As Long

    If nodeA < nodeB Then
        lowNode = nodeA: highNode = nodeB
    Else
        lowNode = nodeB: highNode = nodeA
    End If
    keyName = CStr(lowNode) & "-" & CStr(highNode)
    On Error Resume Next
    existingIndex = CLng(edgeLookup.item(keyName))
    If Err.Number = 0 Then
        On Error GoTo 0
        MeshGetEdgeMidpoint = existingIndex
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    edgeCount = edgeCount + 1
    If edgeCount > UBound(edgeA) Then
        Err.Raise vbObjectError + 3463, "CONTROL.MeshGetEdgeMidpoint", _
                  "Q8辺配列容量を超えました。分割数を下げてください。"
    End If
    midX = (nodeX(nodeA) + nodeX(nodeB)) / 2#
    midY = (nodeY(nodeA) + nodeY(nodeB)) / 2#
    MeshAppendNode nodeCount, nodeX, nodeY, midX, midY
    edgeA(edgeCount) = lowNode
    edgeB(edgeCount) = highNode
    edgeMid(edgeCount) = nodeCount
    edgeLookup.Add nodeCount, keyName
    MeshGetEdgeMidpoint = nodeCount
End Function

Private Sub MeshWriteTriangleWedges(ByVal nodeA As Long, ByVal nodeB As Long, ByVal nodeC As Long, _
                                    ByRef nodeX() As Double, ByRef nodeY() As Double, _
                                    ByRef edgeA() As Long, ByRef edgeB() As Long, _
                                    ByRef edgeMid() As Long, ByRef edgeCount As Long, _
                                    ByRef nodeCount As Long, ByRef elementCount As Long, _
                                    ByRef elementRows() As Variant, ByVal edgeLookup As Collection, _
                                    ByVal materialNumber As Long)
    Dim midAB As Long, midBC As Long, midCA As Long, centerNode As Long
    Dim q8m1 As Long, q8m2 As Long, q8m3 As Long, q8m4 As Long
    Dim centerX As Double, centerY As Double

    midAB = MeshGetEdgeMidpoint(nodeA, nodeB, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    midBC = MeshGetEdgeMidpoint(nodeB, nodeC, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    midCA = MeshGetEdgeMidpoint(nodeC, nodeA, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    centerX = (nodeX(nodeA) + nodeX(nodeB) + nodeX(nodeC)) / 3#
    centerY = (nodeY(nodeA) + nodeY(nodeB) + nodeY(nodeC)) / 3#
    MeshAppendNode nodeCount, nodeX, nodeY, centerX, centerY
    centerNode = nodeCount

    q8m1 = MeshGetEdgeMidpoint(midCA, nodeA, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    q8m2 = MeshGetEdgeMidpoint(nodeA, midAB, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    q8m3 = MeshGetEdgeMidpoint(midAB, centerNode, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    q8m4 = MeshGetEdgeMidpoint(centerNode, midCA, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    MeshWriteQ8Element elementCount, elementRows, midCA, nodeA, midAB, centerNode, q8m1, q8m2, q8m3, q8m4, materialNumber

    q8m1 = MeshGetEdgeMidpoint(midAB, nodeB, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    q8m2 = MeshGetEdgeMidpoint(nodeB, midBC, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    q8m3 = MeshGetEdgeMidpoint(midBC, centerNode, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    q8m4 = MeshGetEdgeMidpoint(centerNode, midAB, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    MeshWriteQ8Element elementCount, elementRows, midAB, nodeB, midBC, centerNode, q8m1, q8m2, q8m3, q8m4, materialNumber

    q8m1 = MeshGetEdgeMidpoint(midBC, nodeC, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    q8m2 = MeshGetEdgeMidpoint(nodeC, midCA, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    q8m3 = MeshGetEdgeMidpoint(midCA, centerNode, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    q8m4 = MeshGetEdgeMidpoint(centerNode, midBC, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, edgeLookup)
    MeshWriteQ8Element elementCount, elementRows, midBC, nodeC, midCA, centerNode, q8m1, q8m2, q8m3, q8m4, materialNumber
End Sub

Private Sub MeshCreateTriangleQuadratic(ByRef pointX() As Double, ByRef pointY() As Double, _
                                        ByVal subdivisionX As Long, ByVal subdivisionY As Long)
    Dim subdivision As Long
    Dim row As Long, col As Long, nodeCount As Long, elementCount As Long, edgeCount As Long
    Dim nodeA As Long, nodeB As Long, nodeC As Long, nodeD As Long
    Dim baseNode() As Long
    Dim nodeX() As Double, nodeY() As Double
    Dim edgeA() As Long, edgeB() As Long, edgeMid() As Long
    Dim elementRows() As Variant
    Dim edgeCapacity As Long, nodeCapacity As Long
    Dim u As Double, v As Double, w As Double
    Dim xx As Double, yy As Double
    Dim expectedElementCount As Long
    Dim materialNumber As Long
    Dim edgeLookup As Collection

    subdivision = subdivisionX
    If subdivision < subdivisionY Then subdivision = subdivisionY
    If subdivision < 1 Then
        Err.Raise vbObjectError + 3464, "CONTROL.MeshCreateTriangleQuadratic", _
                  "三角形メッシュの分割数は1以上にしてください。"
    End If
    expectedElementCount = 3 * subdivision * subdivision
    If expectedElementCount > 60000 Then
        Err.Raise vbObjectError + 3465, "CONTROL.MeshCreateTriangleQuadratic", _
                  "指定した分割数では三角形Q8メッシュの要素容量を超えます。分割数を下げてください。"
    End If
    nodeCapacity = 12 * subdivision * subdivision + 32
    If nodeCapacity > 60000 Then nodeCapacity = 60000
    edgeCapacity = 8 * expectedElementCount + 8
    If edgeCapacity > 60000 Then
        Err.Raise vbObjectError + 3466, "CONTROL.MeshCreateTriangleQuadratic", _
                  "Q8辺配列容量を超えます。分割数を下げてください。"
    End If

    ReDim baseNode(0 To subdivision, 0 To subdivision)
    ReDim nodeX(1 To nodeCapacity)
    ReDim nodeY(1 To nodeCapacity)
    ReDim edgeA(1 To edgeCapacity)
    ReDim edgeB(1 To edgeCapacity)
    ReDim edgeMid(1 To edgeCapacity)
    ReDim elementRows(1 To expectedElementCount, 1 To 10)
    Set edgeLookup = New Collection
    materialNumber = MeshReadMaterialNumber()

    For row = 0 To subdivision
        For col = 0 To subdivision - row
            u = CDbl(col) / CDbl(subdivision)
            v = CDbl(row) / CDbl(subdivision)
            w = 1# - u - v
            xx = w * pointX(1) + u * pointX(2) + v * pointX(3)
            yy = w * pointY(1) + u * pointY(2) + v * pointY(3)
            MeshAppendNode nodeCount, nodeX, nodeY, xx, yy
            baseNode(row, col) = nodeCount
        Next col
    Next row

    For row = 0 To subdivision - 1
        For col = 0 To subdivision - row - 1
            nodeA = baseNode(row, col)
            nodeB = baseNode(row, col + 1)
            nodeC = baseNode(row + 1, col)
            MeshWriteTriangleWedges nodeA, nodeB, nodeC, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, elementCount, elementRows, edgeLookup, materialNumber
            If col < subdivision - row - 1 Then
                nodeD = baseNode(row + 1, col + 1)
                MeshWriteTriangleWedges nodeB, nodeD, nodeC, nodeX, nodeY, edgeA, edgeB, edgeMid, edgeCount, nodeCount, elementCount, elementRows, edgeLookup, materialNumber
            End If
        Next col
    Next row
    P5ValidateMeshCapacity nodeCount, elementCount
    MeshCommitGeneratedMesh nodeCount, elementCount, nodeX, nodeY, elementRows
End Sub

Sub ボタン1_Click()
    Dim pointX(1 To 4) As Double, pointY(1 To 4) As Double
    Dim uniqueX(1 To 4) As Double, uniqueY(1 To 4) As Double
    Dim i As Long, uniqueCount As Long
    Dim nx As Long, my As Long
    Dim tolerance As Double
    Dim errorNumber As Long, errorDescription As String

    On Error GoTo ErrorHandler
    FEMWriteBuildStamp
    FEMInvalidateSettingCache
    For i = 1 To 4
        pointX(i) = FEMReadSetting("MESH_POINT" & CStr(i) & "_X", 0#)
        pointY(i) = FEMReadSetting("MESH_POINT" & CStr(i) & "_Y", 0#)
    Next i
    nx = CLng(FEMReadSetting("MESH_NX", 1#))
    my = CLng(FEMReadSetting("MESH_NY", 1#))
    If nx <= 0 Or my <= 0 Then
        MsgBox "設定シートのMESH_NX/MESH_NYは1以上にしてください。", vbExclamation
        Exit Sub
    End If

    tolerance = MeshCoordinateTolerance(pointX, pointY, 4)
    uniqueCount = MeshUniquePointCount(pointX, pointY, 4, tolerance, uniqueX, uniqueY)
    If uniqueCount <= 2 Then
        MsgBox "要素作成①: 重複を除いた有効な座標点が2点以下です。3点または4点を入力してください。", vbCritical
        Exit Sub
    End If

    ClearElementInputData
    If uniqueCount = 3 Then
        For i = 1 To 3
            pointX(i) = uniqueX(i)
            pointY(i) = uniqueY(i)
        Next i
        MeshNormalizeCounterClockwise pointX, pointY, 3, tolerance
        MeshCreateTriangleQuadratic pointX, pointY, nx, my
    Else
        MeshNormalizeCounterClockwise pointX, pointY, 4, tolerance
        MeshCreateQuadrilateral pointX, pointY, nx, my
    End If
    If uniqueCount = 3 Then
        Call P5EvaluateCurrentSheet("要素作成①-三角形")
    Else
        Call P5EvaluateCurrentSheet("要素作成①-四辺形")
    End If
    If Not P5QualityMeetsSettings() And Not SuppressUserMessages Then
        MsgBox "要素作成①はメッシュ品質の目標値を下回っています。設定値を確認し、要素修正を実行してください。" & vbCrLf & P5QualitySummary(), vbExclamation
    End If
    Exit Sub
ErrorHandler:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MsgBox "要素作成①に失敗しました [" & CStr(errorNumber) & "] " & errorDescription, vbCritical
End Sub
Sub ボタン2_Click()
    MsgBox "旧形式の要素読込はQ8統一のため使用できません。要素作成①または要素作成②を使用してください。", vbExclamation
End Sub
Sub ボタン3_Click()
    Dim i As Long
    Dim errorNumber As Long
    Dim errorDescription As String
    P1DrawOriginalViewer
    Exit Sub
    If AnalysisRunning Then
        MsgBox "解析処理中です。", vbExclamation
        Exit Sub
    End If
    SaveApplicationState
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo ErrorHandler
    ControlInvalidateP1Output
    ControlInvalidateP2Output
    ControlInvalidateP2Output
    ValidateInputData
    SetNodeData
    SetElementData
    For i = 0 To NumberOfElement - 1
        With Elem(i)
            DrawElement i + 1, .x(0), .y(0), .x(4), .y(4), .x(1), .y(1), .x(5), .y(5), .x(2), .y(2), .x(6), .y(6), .x(3), .y(3), .x(7), .y(7)
        End With
    Next i
    For i = 0 To NumberOfNode - 1
        DrawNo i + 1, x(i), y(i)
    Next i
    RestoreApplicationState
    Exit Sub
ErrorHandler:
    errorNumber = Err.Number
    errorDescription = Err.Description
    RestoreApplicationState
    MsgBox "描画用データの読み込み失敗 [" & CStr(errorNumber) & "] " & errorDescription, vbCritical
End Sub
Sub ボタン4_Click()
    Dim wsFigure As Worksheet
    Dim wsIntegratedResult As Worksheet
    Dim shapeNo As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Set wsFigure = ThisWorkbook.Worksheets("図")

    '図形の種類を限定しない。要素線・節点番号・P1表示・凡例を一括削除する。
    For shapeNo = wsFigure.Shapes.count To 1 Step -1
        wsFigure.Shapes.item(shapeNo).Delete
    Next shapeNo

    '旧形式の結果シートと統合結果シートの計算結果を消去する。
    Cleardata
    ControlInvalidateP1Output
    ControlInvalidateP2Output
    ClearControlResultSheet "P2材料点試験"
    ClearControlResultSheet "P1結果"
    ClearControlResultSheet "P2材料点結果"
    Set wsIntegratedResult = Nothing
    On Error Resume Next
    Set wsIntegratedResult = ThisWorkbook.Worksheets("診断")
    On Error GoTo ErrorHandler
    If Not wsIntegratedResult Is Nothing Then wsIntegratedResult.Cells.ClearContents

    '解析済み状態も無効化し、クリア後の表示ボタンが古い結果を参照しないようにする。
    AnalysisOK = False
    ResultStatus = RESULT_NOT_RUN
    ResultRevision = 0
    AnalysisMessage = ""
    AnalysisErrorNumber = 0
    MatrixFactored = False
    P1ResultReady = False
    P1ResultModelRevision = 0

    Application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Application.ScreenUpdating = True
    MsgBox "クリアに失敗しました [" & CStr(errorNumber) & "] " & errorDescription, vbCritical
End Sub

Private Sub ClearControlResultSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then ws.UsedRange.ClearContents
End Sub

Private Sub ControlInvalidateP1Output()
    ClearControlResultSheet "P1結果"
End Sub

Private Sub ControlInvalidateP2Output()
    ClearControlResultSheet "P2材料点結果"
End Sub

Private Sub ClearElementInputData()
    Dim ws As Worksheet
    Dim lastRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("節点データ")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
        If lastRow >= 2 Then ws.range(ws.Cells(2, 1), ws.Cells(lastRow, 14)).ClearContents
    End If
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("要素データ")
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
        If lastRow >= 2 Then ws.range(ws.Cells(2, 1), ws.Cells(lastRow, 10)).ClearContents
    End If
    FEMInvalidateInputCache
End Sub


Sub ボタン5_Click()
    Dim primaryError As Long
    Dim primaryDescription As String
    Dim reportError As Long
    Dim reportDescription As String
    Dim stageStart As Double
    If AnalysisRunning Then
        If Not SuppressUserMessages Then MsgBox "解析処理中です。", vbExclamation
        Exit Sub
    End If
    AnalysisRunning = True
    On Error GoTo ErrorHandler
    FEMViewerMarkDirty
    P1ClearViewerLayers
    SaveApplicationState
    FEMProgressBegin
    FEMLastProc = "P0_RunAnalysis"
    ResetAnalysisState
    FEMWriteBuildStamp
    FEMInvalidateInputCache
    FEMInvalidateSettingCache
    FEMIoBegin
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo ErrorHandler
    ControlInvalidateP1Output
    start_day = CLng(Date)
    start_time = Timer
    stageStart = Timer
    If SuppressUserMessages Then SaveP0Progress "start"
    ' P5: 入力検証を出力シートの消去より前に行う。
    MeshInsertJointsIfNeeded
    ValidateInputData
    If SuppressUserMessages Then SaveP0Progress "validate_complete"
    Cleardata
    If SuppressUserMessages Then SaveP0Progress "clear_output_complete"
    SetNodeData
    If SuppressUserMessages Then SaveP0Progress "node_input_complete"
    SetMaterialData
    If SuppressUserMessages Then SaveP0Progress "material_input_complete"
    SetElementData
    If SuppressUserMessages Then SaveP0Progress "element_input_complete"
    TimeInput = TimerElapsed(stageStart)
    If SuppressUserMessages Then SaveP0Progress "input_complete"
    stageStart = Timer
    FEMProgressPhase "要素剛性の作成"
    If Not SetElmMat() Then GoTo AnalysisFailure
    TimeElementStiffness = TimerElapsed(stageStart)
    If SuppressUserMessages Then SaveP0Progress "element_stiffness_complete"
    stageStart = Timer
    FEMProgressPhase "全体剛性・境界条件の準備"
    SetTotalMat
    WeightUpdate
    SetBoundaryCondition
    TimeAssembly = TimerElapsed(stageStart)
    If SuppressUserMessages Then SaveP0Progress "assembly_complete"
    stageStart = Timer
    CurrentIncrement = 0
    CurrentIteration = 0
    ' 線形予備求解はしない。P3が零変位からNewtonする。
    TimeSolve = 0#
    If SuppressUserMessages Then SaveP0Progress "linear_presolve_skipped"
    FEMProgressPhase "ステージ解析", "初期状態を準備しています"
    If Not PlaCalc() Then GoTo AnalysisFailure
    TimePlastic = TimerElapsed(stageStart)
    If SuppressUserMessages Then SaveP0Progress "plastic_complete"
    FEMProgressPhase "結果出力", "応力・変位・解析結果を保存しています"
    AnalysisOK = True
    BuildP1Results
    If SuppressUserMessages Then SaveP0Progress "p1_result_recovery_complete"
    stageStart = Timer
    P3PrepareOutputResidual
    ComputeResidual
    AnalysisOK = True
    ResultStatus = RESULT_PASS
    ResultRevision = ModelRevision
    If UCase$(Trim$(FEMReadTextSetting("OUTPUT_STAGE_MODE", "FINAL"))) <> "ALL" Then
        SaveStress
        SaveDisp
    End If
    SaveP1Results
    SaveP2MaterialResults
    TimeOutput = TimerElapsed(stageStart)
    SaveAnalysisReport
    FEMIoClose
    FEMViewerMarkClean
    MeshStampResult
    now_day = CLng(Date)
    now_time = Timer
    calc_time = CLng(now_time - start_time) + (now_day - start_day) * CLng(3600) * CLng(24)
    RestoreApplicationState
    AnalysisRunning = False
    If P3SrmEnabled Then
        FEMProgressFinish True, P3SrmNote
    Else
        FEMProgressFinish True, "すべてのステージが終了しました"
    End If
    If Not SuppressUserMessages Then
        If P3SrmEnabled Then
            MsgBox "計算終了" & vbCrLf & P3SrmNote
        Else
            MsgBox "計算終了"
        End If
        MsgBox Int(CLng(calc_time) / (CLng(3600) * CLng(24))) & "日" & _
            Int((calc_time) / 3600) Mod 24 & "時間" & _
            Int((calc_time) / 60) Mod 60 & "分" & _
            Int(calc_time) Mod 60 & "秒"
    End If
    Exit Sub
AnalysisFailure:
    If Len(AnalysisMessage) = 0 Then
        SetAnalysisFailure RESULT_RUNTIME_ERROR, "解析に失敗しました。", vbObjectError + 3399, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
    End If
    GoTo ErrorHandler
ErrorHandler:
    primaryError = Err.Number
    primaryDescription = Err.Description
    If primaryError = 0 Then primaryError = AnalysisErrorNumber
    If primaryError = 0 Then primaryError = vbObjectError + 3399
    If Len(primaryDescription) = 0 Then primaryDescription = AnalysisMessage
    If Len(primaryDescription) = 0 Then primaryDescription = "解析に失敗しました。"
    If Len(FEMLastProc) > 0 Then primaryDescription = primaryDescription & " 場所=" & FEMLastProc
    If primaryError = 18 Then
        SetAnalysisFailure RESULT_NONCONVERGED, "ユーザーが解析を中断しました。", 18, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
        GoTo SkipFailureAssign
    End If
    If ResultStatus = RESULT_NOT_RUN Or ResultStatus = RESULT_PASS Then
        SetAnalysisFailure ErrorStatus(primaryError), primaryDescription, primaryError, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
    Else
        AnalysisErrorNumber = primaryError
        If Len(AnalysisMessage) = 0 Then AnalysisMessage = primaryDescription
    End If
SkipFailureAssign:
    On Error Resume Next
    FEMIoClose
    SaveAnalysisReport
    reportError = Err.Number
    reportDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    RestoreApplicationState
    AnalysisRunning = False
    FEMProgressFinish False, ResultStatus & "  " & AnalysisMessage
    If reportError <> 0 Then AnalysisMessage = AnalysisMessage & " 診断レポート保存失敗: " & reportDescription
    If Not SuppressUserMessages Then MsgBox "計算失敗 [" & ResultStatus & "] " & AnalysisMessage, vbCritical
End Sub
Sub ボタン6_Click()
Dim dscale As Variant: Dim i1 As Variant: Dim i2 As Variant: Dim x1 As Variant: Dim x2 As Variant
Dim x3 As Variant: Dim x4 As Variant: Dim x5 As Variant: Dim x6 As Variant: Dim x7 As Variant
Dim x8 As Variant: Dim y1 As Variant: Dim y2 As Variant: Dim y3 As Variant: Dim y4 As Variant
Dim y5 As Variant: Dim y6 As Variant: Dim y7 As Variant: Dim y8 As Variant
Dim i As Long, CoutN As Long
 P1DrawDeformedViewer
 Exit Sub
 
 NumberOfNode = SetNumberOfNode(): SetNodeData
 NumberOfElement = SetNumberOfElement(): SetElementData
 
 ReadFreeNode
 
 CoutN = NumberOfFreeNode * 2 + 1
 ReDim Disp(CoutN + 1)
 ReDim XXX(CoutN + 1)
    


 With ThisWorkbook.Worksheets("結果節点")
     For i = 0 To NumberOfFreeNode
        i1 = i * 2: i2 = i1 + 1
        XXX(i1) = .Cells(i + 2, 2)
        XXX(i2) = .Cells(i + 2, 3)
        Disp(i1) = .Cells(i + 2, 4)
        Disp(i2) = .Cells(i + 2, 5)
     Next i
 End With
 
 
 dscale = FEMReadSetting("VIEW_DISP_SCALE", 100#)
 For i = 0 To NumberOfElement - 1
    With Elem(i)
    'X(0), .Y(0), .X(4), .Y(4), .X(1), .Y(1), .X(5), .Y(5), .X(2), Y(2), .X(6), .Y(6), .X(3), .Y(3), .X(7), .Y(7)
    x1 = .x(0) + dscale * Disp(.node(0) * 2)
    y1 = .y(0) + dscale * Disp(.node(0) * 2 + 1)
    x2 = .x(4) + dscale * Disp(.node(4) * 2)
    y2 = .y(4) + dscale * Disp(.node(4) * 2 + 1)
    x3 = .x(1) + dscale * Disp(.node(1) * 2)
    y3 = .y(1) + dscale * Disp(.node(1) * 2 + 1)
    x4 = .x(5) + dscale * Disp(.node(5) * 2)
    y4 = .y(5) + dscale * Disp(.node(5) * 2 + 1)
    x5 = .x(2) + dscale * Disp(.node(2) * 2)
    y5 = .y(2) + dscale * Disp(.node(2) * 2 + 1)
    x6 = .x(6) + dscale * Disp(.node(6) * 2)
    y6 = .y(6) + dscale * Disp(.node(6) * 2 + 1)
    x7 = .x(3) + dscale * Disp(.node(3) * 2)
    y7 = .y(3) + dscale * Disp(.node(3) * 2 + 1)
    x8 = .x(7) + dscale * Disp(.node(7) * 2)
    y8 = .y(7) + dscale * Disp(.node(7) * 2 + 1)
    DrawDispElement i + 1, x1, y1, x2, y2, x3, y3, x4, y4, x5, y5, x6, y6, x7, y7, x8, y8
    End With
 Next i
End Sub
Sub ボタン7_Click()
Dim ds1 As Variant: Dim ds2 As Variant: Dim idd As Variant: Dim ii As Variant: Dim sMax As Variant
Dim sMin As Variant: Dim stressforcolor() As Variant
Dim i As Long, CC As Long
 P1DrawContourViewer
 Exit Sub
 NumberOfNode = SetNumberOfNode(): SetNodeData
 NumberOfElement = SetNumberOfElement(): SetElementData
 NumberOfMaterial = SetNumberOfMaterial(): SetMaterialData
 
 ReDim stressforcolor(NumberOfElement)
 idd = FEMReadSetting("VIEW_COLOR_RESULT", P1_RESULT_VON_MISES_TOP)
 With ThisWorkbook.Worksheets("結果要素")
 sMax = 0: sMin = 0
 For i = 0 To NumberOfElement - 1
   stressforcolor(i) = .Cells(i + 2, idd + 1): CC = stressforcolor(i)
   If CC > 0 And CC > sMax Then sMax = CC
   If CC < 0 And CC < sMin Then sMin = CC
 Next i
 End With
 With ThisWorkbook.Worksheets("図")
 If sMin = sMax Then
  MsgBox "最大・最小がないので色付けできません"
  sMax = 255: sMin = -255
 End If
 ds1 = sMax / 255: ds2 = sMin / 255
 If Abs(ds2) > ds1 Then ds1 = Abs(ds2)
 
 For i = 0 To NumberOfElement - 1
   ActiveSheet.Shapes("E_" & (i + 1)).Select
   If stressforcolor(i) >= 0 Then
     ii = Int(stressforcolor(i) / ds1)
     If ii > 255 Then ii = 255
     CC = RGB(255, 255 - ii, 255 - ii)
   Else
     ii = Int(Abs(stressforcolor(i)) / ds1)
     If ii > 255 Then ii = 255
     CC = RGB(255 - ii, 255 - ii, 255)
   End If
   selection.ShapeRange.fill.ForeColor.RGB = CC
   selection.ShapeRange.fill.Visible = msoTrue
   selection.ShapeRange.fill.Solid
 Next i
 End With
End Sub

Sub ボタン20_Click()
    MeshOptimizeExisting
End Sub

Private Function P1LoadViewerMeshFromSheets() As Boolean
    Dim wsNode As Worksheet, wsElement As Worksheet
    Dim nodeValues As Variant, elementValues As Variant
    Dim nodeCount As Long, elementCount As Long
    Dim rowNo As Long, localNode As Long, nodeId As Long, elementId As Long
    Dim lastNodeRow As Long, lastElementRow As Long
    Dim errorNumber As Long, errorDescription As String

    On Error GoTo P1ViewerMeshLoadError
    Set wsNode = ThisWorkbook.Worksheets("節点データ")
    Set wsElement = ThisWorkbook.Worksheets("要素データ")
    lastNodeRow = wsNode.Cells(wsNode.rows.count, 1).End(xlUp).row
    lastElementRow = wsElement.Cells(wsElement.rows.count, 1).End(xlUp).row
    nodeCount = lastNodeRow - 1
    elementCount = lastElementRow - 1
    If nodeCount < 1 Then Err.Raise vbObjectError + 3490, "CONTROL.P1LoadViewerMeshFromSheets", "図形表示用の節点データがありません。"
    If elementCount < 1 Then Err.Raise vbObjectError + 3491, "CONTROL.P1LoadViewerMeshFromSheets", "図形表示用の要素データがありません。"

    ' 図形表示は材料番号を必要としないため、解析用ValidateInputDataを経由しない。
    ' 要素作成②の材料番号が材料データに未登録でも、Q8形状そのものは表示できる。
    nodeValues = wsNode.range(wsNode.Cells(2, 1), wsNode.Cells(nodeCount + 1, 3)).value2
    elementValues = wsElement.range(wsElement.Cells(2, 1), wsElement.Cells(elementCount + 1, 9)).value2
    NumberOfNode = nodeCount
    NumberOfElement = elementCount
    NumberOfFreeNode = nodeCount
    ReDim x(nodeCount - 1)
    ReDim y(nodeCount - 1)
    ReDim Elem(elementCount - 1)

    For rowNo = 1 To nodeCount
        nodeId = CLng(nodeValues(rowNo, 1))
        If nodeId < 1 Or nodeId > nodeCount Then Err.Raise vbObjectError + 3492, "CONTROL.P1LoadViewerMeshFromSheets", "節点番号が連番範囲外です。"
        x(nodeId - 1) = CDbl(nodeValues(rowNo, 2))
        y(nodeId - 1) = CDbl(nodeValues(rowNo, 3))
    Next rowNo

    For rowNo = 1 To elementCount
        elementId = CLng(elementValues(rowNo, 1))
        If elementId < 1 Or elementId > elementCount Then Err.Raise vbObjectError + 3493, "CONTROL.P1LoadViewerMeshFromSheets", "要素番号が連番範囲外です。"
        Elem(elementId - 1).ElNo = elementId - 1
        For localNode = 0 To 7
            If IsError(elementValues(rowNo, localNode + 2)) Then
                nodeId = 0
            ElseIf IsEmpty(elementValues(rowNo, localNode + 2)) Then
                nodeId = 0
            ElseIf Not IsNumeric(elementValues(rowNo, localNode + 2)) Then
                nodeId = 0
            Else
                nodeId = CLng(elementValues(rowNo, localNode + 2))
            End If
            If nodeId < 1 Then
                If localNode >= 6 Then
                    Elem(elementId - 1).node(localNode) = -1
                    Elem(elementId - 1).IsJoint = True
                Else
                    Err.Raise vbObjectError + 3494, "CONTROL.P1LoadViewerMeshFromSheets", "Q8接続節点番号が範囲外です。要素=" & CStr(elementId)
                End If
            Else
                If nodeId > nodeCount Then Err.Raise vbObjectError + 3494, "CONTROL.P1LoadViewerMeshFromSheets", "Q8接続節点番号が範囲外です。要素=" & CStr(elementId)
                Elem(elementId - 1).node(localNode) = nodeId - 1
                Elem(elementId - 1).x(localNode) = x(nodeId - 1)
                Elem(elementId - 1).y(localNode) = y(nodeId - 1)
            End If
        Next localNode
    Next rowNo
    P1LoadViewerMeshFromSheets = True
    Exit Function
P1ViewerMeshLoadError:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, "CONTROL.P1LoadViewerMeshFromSheets", errorDescription
End Function

Private Sub P1ApplyAutoFit(ByRef view As P1_ViewState)
    Const VIEW_LEFT As Double = 40#
    Const VIEW_TOP As Double = 40#
    Const VIEW_WIDTH As Double = 640#
    Const VIEW_HEIGHT As Double = 400#
    Const VIEW_PADDING As Double = 20#
    Dim elementId As Long, localNode As Long
    Dim xValue As Double, yValue As Double
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim widthValue As Double, heightValue As Double, scaleX As Double, scaleY As Double
    Dim fitScale As Double, firstPoint As Boolean

    If NumberOfElement <= 0 Then Exit Sub
    firstPoint = True
    For elementId = 0 To NumberOfElement - 1
        For localNode = 0 To 7
            xValue = Elem(elementId).x(localNode)
            yValue = Elem(elementId).y(localNode)
            If firstPoint Then
                minX = xValue: maxX = xValue: minY = yValue: maxY = yValue
                firstPoint = False
            Else
                If xValue < minX Then minX = xValue
                If xValue > maxX Then maxX = xValue
                If yValue < minY Then minY = yValue
                If yValue > maxY Then maxY = yValue
            End If
        Next localNode
    Next elementId
    widthValue = maxX - minX
    heightValue = maxY - minY
    If widthValue <= 1E-30 Then widthValue = 1#
    If heightValue <= 1E-30 Then heightValue = 1#
    scaleX = (VIEW_WIDTH - 2# * VIEW_PADDING) / widthValue
    scaleY = (VIEW_HEIGHT - 2# * VIEW_PADDING) / heightValue
    fitScale = scaleX
    If scaleY < fitScale Then fitScale = scaleY
    If fitScale <= 0# Then fitScale = 1#
    view.drawScale = fitScale
    view.originX = VIEW_LEFT + VIEW_PADDING - minX * fitScale
    view.originY = VIEW_TOP + VIEW_PADDING + maxY * fitScale
End Sub

Private Sub P1ReadViewState(ByRef view As P1_ViewState)
    Dim candidateResultId As Long
    view.originX = 0#: view.originY = 0#: view.drawScale = 1#: view.DeformedScale = 1#
    view.originX = FEMReadSetting("VIEW_MIN_X", 0#)
    view.originY = FEMReadSetting("VIEW_MAX_Y", 0#)
    view.drawScale = FEMReadSetting("VIEW_SCALE", 1#)
    view.DeformedScale = FEMReadSetting("VIEW_DISP_SCALE", 1#)
    view.resultId = P1_RESULT_VON_MISES_TOP
    candidateResultId = CLng(FEMReadSetting("VIEW_COLOR_RESULT", P1_RESULT_VON_MISES_TOP))
    If candidateResultId >= 0 And candidateResultId < P1_RESULT_COUNT Then view.resultId = candidateResultId
    If view.drawScale <= 0# Then view.drawScale = 1#
    If view.DeformedScale <= 0# Then view.DeformedScale = 1#
    view.ProjectionMatrix(0, 0) = 1#: view.ProjectionMatrix(0, 1) = 0#: view.ProjectionMatrix(0, 2) = 0#
    view.ProjectionMatrix(1, 0) = 0#: view.ProjectionMatrix(1, 1) = 1#: view.ProjectionMatrix(1, 2) = 0#
    view.ProjectionMatrix(2, 0) = 0#: view.ProjectionMatrix(2, 1) = 0#: view.ProjectionMatrix(2, 2) = 1#
    view.AngleX = 0#: view.AngleY = 0#: view.AngleZ = 0#: view.AutoFit = False
    view.ShowNodeLabels = NumberOfFreeNode <= 500
    view.ShowElementLabels = False
    view.ShowOriginal = True: view.ShowDeformed = False: view.ShowContour = False
    view.scaleMode = UCase$(Trim$(FEMReadTextSetting("VIEW_SCALE_MODE", "SYMMETRIC")))
    If view.scaleMode <> "AUTO" And view.scaleMode <> "SYMMETRIC" And view.scaleMode <> "USER" Then view.scaleMode = "SYMMETRIC"
    view.DetailMode = UCase$(Trim$(FEMReadTextSetting("VIEW_DETAIL_MODE", "SIMPLE")))
    If view.DetailMode <> "SIMPLE" And view.DetailMode <> "DETAIL" Then view.DetailMode = "SIMPLE"
    view.ResultSource = UCase$(Trim$(FEMReadTextSetting("VIEW_RESULT_SOURCE", "SMOOTHED")))
    If view.ResultSource <> "ELEMENT" And view.ResultSource <> "RAW_GAUSS" Then view.ResultSource = "SMOOTHED"
    '旧版のVIEW_MIN_X/VIEW_MAX_Y/VIEW_SCALEが残っていても、通常表示は現行メッシュへ自動フィットする。
    'USER指定時だけ設定値を尊重し、①・②のメッシュ変更後の画面外表示を防止する。
    If view.scaleMode <> "USER" Then P1ApplyAutoFit view
End Sub

Private Function P1EnsureViewerReady(ByVal operationName As String) As Boolean
    Dim upperBound As Long, errorNumber As Long
    P1EnsureViewerReady = False
    If Not AnalysisOK Then
        If Not SuppressUserMessages Then MsgBox operationName & "は解析成功後に実行してください。", vbExclamation
        Exit Function
    End If
    If ResultRevision <> ModelRevision Or P1ResultModelRevision <> ModelRevision Then
        If Not SuppressUserMessages Then MsgBox "解析結果の版が現在のモデルと一致しません。再解析してください。", vbExclamation
        Exit Function
    End If
    If Not P1ResultReady Then
        If Not SuppressUserMessages Then MsgBox "P1結果が未作成です。再解析してください。", vbExclamation
        Exit Function
    End If
    On Error Resume Next
    upperBound = UBound(TDisp)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    If errorNumber <> 0 Then
        If Not SuppressUserMessages Then MsgBox "変位配列が初期化されていません。", vbExclamation
        Exit Function
    End If
    If upperBound < lastDof Then
        If Not SuppressUserMessages Then MsgBox "変位配列の件数が自由度数と一致しません。", vbExclamation
        Exit Function
    End If
    P1EnsureViewerReady = True
End Function

Private Sub P1ClearViewerLayer(ByRef ws As Worksheet, ByVal prefix As String)
    Dim shapeNo As Long, shapeName As String
    For shapeNo = ws.Shapes.count To 1 Step -1
        shapeName = ws.Shapes.item(shapeNo).name
        If left$(shapeName, Len(prefix)) = prefix Then ws.Shapes.item(shapeNo).Delete
    Next shapeNo
End Sub

Private Function P1IsViewerShape(ByVal shapeName As String) As Boolean
    P1IsViewerShape = False
    If left$(shapeName, 4) = "P1E_" Then P1IsViewerShape = True
    If left$(shapeName, 4) = "P1D_" Then P1IsViewerShape = True
    If left$(shapeName, 4) = "P1C_" Then P1IsViewerShape = True
    If left$(shapeName, 7) = "P1Node_" Then P1IsViewerShape = True
    If left$(shapeName, 9) = "P1Legend_" Then P1IsViewerShape = True
    If left$(shapeName, 6) = "P1Agg_" Then P1IsViewerShape = True
End Function

Public Sub P1ClearViewerLayers()
    Dim ws As Worksheet
    Dim shapeNo As Long, shapeName As String
    Set ws = ThisWorkbook.Worksheets("図")
    For shapeNo = ws.Shapes.count To 1 Step -1
        shapeName = ws.Shapes.item(shapeNo).name
        If P1IsViewerShape(shapeName) Then ws.Shapes.item(shapeNo).Delete
    Next shapeNo
End Sub

Private Sub P1PrepareFigureViewport(ByRef ws As Worksheet)
    Dim figureWindow As Object
    On Error Resume Next
    ws.Activate
    Set figureWindow = ws.parent.Windows(1)
    figureWindow.Zoom = 100
    figureWindow.ScrollRow = 1
    figureWindow.ScrollColumn = 1
    On Error GoTo 0
End Sub

Private Sub P1ProjectPointCached(ByRef view As P1_ViewState, ByVal xValue As Double, ByVal yValue As Double, ByVal zValue As Double, ByRef screenX As Double, ByRef screenY As Double)
    Dim projectedX As Double, projectedY As Double
    projectedX = view.ProjectionMatrix(0, 0) * xValue + view.ProjectionMatrix(0, 1) * yValue + view.ProjectionMatrix(0, 2) * zValue
    projectedY = view.ProjectionMatrix(1, 0) * xValue + view.ProjectionMatrix(1, 1) * yValue + view.ProjectionMatrix(1, 2) * zValue
    screenX = view.originX + view.drawScale * projectedX
    screenY = view.originY - view.drawScale * projectedY
End Sub

Private Sub P1AddPolygonShape(ByRef ws As Worksheet, ByVal shapeName As String, ByVal fillColor As Long, ByVal useFill As Boolean, ByVal pointCount As Long, ByRef pointX() As Double, ByRef pointY() As Double)
    Dim builder As Object, shapeItem As Object, pointNo As Long
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    If useFill Then
        Set builder = ws.Shapes.BuildFreeform(msoEditingAuto, pointX(0), pointY(0))
        For pointNo = 1 To pointCount - 1
            builder.AddNodes msoSegmentLine, msoEditingAuto, pointX(pointNo), pointY(pointNo)
        Next pointNo
        builder.AddNodes msoSegmentLine, msoEditingAuto, pointX(0), pointY(0)
        Set shapeItem = builder.ConvertToShape
        shapeItem.name = shapeName
        shapeItem.fill.ForeColor.RGB = fillColor
        shapeItem.fill.Visible = msoTrue
        shapeItem.line.Visible = msoFalse
        Exit Sub
    End If
    Set builder = ws.Shapes.BuildFreeform(msoEditingAuto, pointX(0), pointY(0))
    For pointNo = 1 To pointCount - 1
        builder.AddNodes msoSegmentLine, msoEditingAuto, pointX(pointNo), pointY(pointNo)
    Next pointNo
    builder.AddNodes msoSegmentLine, msoEditingAuto, pointX(0), pointY(0)
    Set shapeItem = builder.ConvertToShape
    shapeItem.name = shapeName
    shapeItem.fill.Visible = msoFalse
    shapeItem.line.ForeColor.RGB = RGB(0, 0, 0)
End Sub

Private Function P1ElementStride(ByVal elementCount As Long, ByVal subdiv As Long) As Long
    Dim estimatedCount As Double
    estimatedCount = CDbl(elementCount) * CDbl(subdiv) * CDbl(subdiv)
    If estimatedCount <= P1_VIEW_SAFE_SHAPE_LIMIT Then
        P1ElementStride = 1
    Else
        P1ElementStride = CLng((estimatedCount + P1_VIEW_SAFE_SHAPE_LIMIT - 1#) \ P1_VIEW_SAFE_SHAPE_LIMIT)
        If P1ElementStride < 1 Then P1ElementStride = 1
    End If
End Function

Private Sub P1ConfigureAggregatedChart(ByRef chartObject As Object, ByRef chart As Object, _
                                       ByVal minX As Double, ByVal maxX As Double, _
                                       ByVal minY As Double, ByVal maxY As Double)
    Dim axisObject As Object
    If maxX <= minX Then maxX = minX + 1#
    If maxY <= minY Then maxY = minY + 1#
    chartObject.left = minX
    chartObject.Top = minY
    chartObject.width = maxX - minX + 20#
    chartObject.Height = maxY - minY + 20#
    On Error Resume Next
    chart.HasLegend = False
    chart.HasTitle = False
    chart.ChartArea.Format.fill.Visible = 0
    chart.ChartArea.Format.line.Visible = 0
    chart.PlotArea.Format.fill.Visible = 0
    chart.PlotArea.Format.line.Visible = 0
    chart.Axes(1).MinimumScale = minX
    chart.Axes(1).MaximumScale = maxX
    chart.Axes(2).MinimumScale = -maxY
    chart.Axes(2).MaximumScale = -minY
    chart.Axes(1).Visible = False
    chart.Axes(2).Visible = False
    chartObject.ShapeRange.line.Visible = 0
    chartObject.Placement = 3
    On Error GoTo 0
End Sub

Private Sub P1AddAggregatedChartSeries(ByRef chart As Object, ByRef chunkX() As Variant, _
                                       ByRef chunkY() As Variant, ByVal pointCount As Long, _
                                       ByVal lineColor As Long)
    Dim seriesX() As Variant, seriesY() As Variant
    Dim pointNo As Long, series As Object
    If pointCount <= 0 Then Exit Sub
    ReDim seriesX(0 To pointCount - 1)
    ReDim seriesY(0 To pointCount - 1)
    For pointNo = 0 To pointCount - 1
        seriesX(pointNo) = chunkX(pointNo)
        seriesY(pointNo) = chunkY(pointNo)
    Next pointNo
    Set series = chart.SeriesCollection.NewSeries
    series.xValues = seriesX
    series.values = seriesY
    series.Format.line.ForeColor.RGB = lineColor
    series.Format.line.weight = 0.75
    series.MarkerStyle = -4142
End Sub

Private Sub P1AddAggregatedScatterSeries(ByRef chart As Object, ByRef chunkX() As Variant, _
                                         ByRef chunkY() As Variant, ByVal pointCount As Long, _
                                         ByVal markerColor As Long)
    Dim seriesX() As Variant, seriesY() As Variant
    Dim pointNo As Long, series As Object
    If pointCount <= 0 Then Exit Sub
    ReDim seriesX(0 To pointCount - 1)
    ReDim seriesY(0 To pointCount - 1)
    For pointNo = 0 To pointCount - 1
        seriesX(pointNo) = chunkX(pointNo)
        seriesY(pointNo) = chunkY(pointNo)
    Next pointNo
    Set series = chart.SeriesCollection.NewSeries
    series.xValues = seriesX
    series.values = seriesY
    On Error Resume Next
    series.Format.line.Visible = 0
    series.MarkerStyle = 8
    series.markerSize = 4
    series.MarkerForegroundColor = markerColor
    series.MarkerBackgroundColor = markerColor
    On Error GoTo 0
End Sub

Private Sub P1AddAggregatedNodeSeries(ByRef chart As Object, ByRef view As P1_ViewState, _
                                      ByVal deformed As Boolean, ByVal lineColor As Long)
    Dim nodeXValues() As Variant, nodeYValues() As Variant
    Dim nodeId As Long, originalNode As Long
    Dim xValue As Double, yValue As Double, uxValue As Double, uyValue As Double
    Dim pointX As Double, pointY As Double
    Dim series As Object
    If NumberOfFreeNode <= 0 Then Exit Sub
    ReDim nodeXValues(0 To NumberOfFreeNode - 1)
    ReDim nodeYValues(0 To NumberOfFreeNode - 1)
    For nodeId = 0 To NumberOfFreeNode - 1
        originalNode = P6GetOriginalFreeNode(nodeId)
        If originalNode < 0 Or originalNode >= NumberOfNode Then originalNode = nodeId
        If originalNode >= 0 And originalNode < NumberOfNode Then
            xValue = x(originalNode)
            yValue = y(originalNode)
            If deformed Then
                P1GetNodeDisplacement nodeId, uxValue, uyValue
                xValue = xValue + view.DeformedScale * uxValue
                yValue = yValue + view.DeformedScale * uyValue
            End If
            P1ProjectPointCached view, xValue, yValue, 0#, pointX, pointY
            nodeXValues(nodeId) = pointX
            nodeYValues(nodeId) = -pointY
        End If
    Next nodeId
    Set series = chart.SeriesCollection.NewSeries
    series.xValues = nodeXValues
    series.values = nodeYValues
    On Error Resume Next
    series.Format.line.Visible = 0
    series.MarkerStyle = 8
    series.markerSize = 2
    series.MarkerForegroundColor = lineColor
    series.MarkerBackgroundColor = lineColor
    On Error GoTo 0
End Sub

Private Sub P1AddAggregatedMeshChart(ByRef ws As Worksheet, ByRef view As P1_ViewState, _
                                     ByVal deformed As Boolean, ByVal shapeName As String, _
                                     ByVal lineColor As Long)
    Const P1_CHART_POINT_LIMIT As Long = 12000
    Dim stride As Long, visibleCount As Long
    Dim chunkX() As Variant, chunkY() As Variant
    Dim elementId As Long, pointNo As Long, pointIndex As Long, sourceNode As Long, nodeId As Long
    Dim pointX As Double, pointY As Double, xValue As Double, yValue As Double
    Dim uxValue As Double, uyValue As Double
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim chartObject As Object, chart As Object, series As Object
    Dim firstPoint As Boolean
    Dim sourceOrder(0 To 7) As Long

    stride = P1ElementStride(NumberOfElement, 1)
    visibleCount = (NumberOfElement + stride - 1) \ stride
    If visibleCount <= 0 Then Exit Sub
    ReDim chunkX(0 To P1_CHART_POINT_LIMIT - 1)
    ReDim chunkY(0 To P1_CHART_POINT_LIMIT - 1)
    sourceOrder(0) = 0: sourceOrder(1) = 4: sourceOrder(2) = 1: sourceOrder(3) = 5
    sourceOrder(4) = 2: sourceOrder(5) = 6: sourceOrder(6) = 3: sourceOrder(7) = 7
    firstPoint = True
    Set chartObject = ws.ChartObjects.Add(0#, 0#, 100#, 100#)
    chartObject.name = shapeName
    Set chart = chartObject.chart
    chart.ChartType = 75
    For elementId = 0 To NumberOfElement - 1 Step stride
        For pointNo = 0 To 8
            If pointNo < 8 Then
                sourceNode = sourceOrder(pointNo)
            Else
                sourceNode = sourceOrder(0)
            End If
            xValue = Elem(elementId).x(sourceNode)
            yValue = Elem(elementId).y(sourceNode)
            If deformed Then
                nodeId = Elem(elementId).node(sourceNode)
                P1GetNodeDisplacement nodeId, uxValue, uyValue
                xValue = xValue + view.DeformedScale * uxValue
                yValue = yValue + view.DeformedScale * uyValue
            End If
            P1ProjectPointCached view, xValue, yValue, 0#, pointX, pointY
            chunkX(pointIndex) = pointX
            chunkY(pointIndex) = -pointY
            If firstPoint Then
                minX = pointX: maxX = pointX: minY = pointY: maxY = pointY
                firstPoint = False
            Else
                If pointX < minX Then minX = pointX
                If pointX > maxX Then maxX = pointX
                If pointY < minY Then minY = pointY
                If pointY > maxY Then maxY = pointY
            End If
            pointIndex = pointIndex + 1
        Next pointNo
        chunkX(pointIndex) = CVErr(2042)
        chunkY(pointIndex) = CVErr(2042)
        pointIndex = pointIndex + 1
        ' Excelの1系列上限（16,384点）を超えないよう要素境界で分割する。
        If pointIndex >= P1_CHART_POINT_LIMIT - 10 Or elementId + stride >= NumberOfElement Then
            P1AddAggregatedChartSeries chart, chunkX, chunkY, pointIndex, lineColor
            pointIndex = 0
        End If
    Next elementId
    '要素線とは独立した節点点系列を追加し、Q8の全節点を表示する。
    P1AddAggregatedNodeSeries chart, view, deformed, RGB(31, 78, 121)
    P1ConfigureAggregatedChart chartObject, chart, minX, maxX, minY, maxY
End Sub

Private Sub P1DrawAggregatedContourLayer(ByRef ws As Worksheet, ByRef view As P1_ViewState, _
                                         ByVal minValue As Double, ByVal maxValue As Double)
    Dim stride As Long, elementId As Long, pointNo As Long, sourceNode As Long, nodeId As Long
    Dim pointX As Double, pointY As Double, xValue As Double, yValue As Double
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim uxValue As Double, uyValue As Double, elementValue As Double
    Dim chartObject As Object, chart As Object
    Dim firstPoint As Boolean, binIndex As Long, pointIndex As Long
    Dim sourceOrder(0 To 7) As Long
    Dim pointCountByBin(0 To P1_CONTOUR_BIN_COUNT - 1) As Long
    Dim seriesX() As Variant, seriesY() As Variant

    stride = P1ElementStride(NumberOfElement, 1)
    sourceOrder(0) = 0: sourceOrder(1) = 4: sourceOrder(2) = 1: sourceOrder(3) = 5
    sourceOrder(4) = 2: sourceOrder(5) = 6: sourceOrder(6) = 3: sourceOrder(7) = 7

    ' 概略表示では要素ごとにChart系列を作らない。系列数を固定し、同色要素を
    ' CVErr(2042)で区切った少数の系列へ集約してExcelの応答停止を防ぐ。
    For elementId = 0 To NumberOfElement - 1 Step stride
        elementValue = ElementAverageResult(elementId, view.resultId)
        If Not P1IsFiniteValue(elementValue) Then elementValue = 0#
        binIndex = P1ContourBinIndex(elementValue, minValue, maxValue)
        ' 8頂点+始点へ戻る点+要素間区切り(#N/A)の10点を確保する。
        pointCountByBin(binIndex) = pointCountByBin(binIndex) + 10
        For pointNo = 0 To 8
            If pointNo < 8 Then
                sourceNode = sourceOrder(pointNo)
            Else
                sourceNode = sourceOrder(0)
            End If
            xValue = Elem(elementId).x(sourceNode)
            yValue = Elem(elementId).y(sourceNode)
            ' 応力コンターは元形状へ重ねる。変形倍率は変形表示専用とし、
            ' 大きな変位倍率で応力図形がつぶれることを防ぐ。
            P1ProjectPointCached view, xValue, yValue, 0#, pointX, pointY
            If firstPoint Then
                minX = pointX: maxX = pointX: minY = pointY: maxY = pointY
                firstPoint = False
            Else
                If pointX < minX Then minX = pointX
                If pointX > maxX Then maxX = pointX
                If pointY < minY Then minY = pointY
                If pointY > maxY Then maxY = pointY
            End If
        Next pointNo
    Next elementId
    If firstPoint Then Exit Sub

    Set chartObject = ws.ChartObjects.Add(0#, 0#, 100#, 100#)
    chartObject.name = "P1Agg_C"
    Set chart = chartObject.chart
    chart.ChartType = 75
    For binIndex = 0 To P1_CONTOUR_BIN_COUNT - 1
        If pointCountByBin(binIndex) > 0 Then
            ReDim seriesX(0 To pointCountByBin(binIndex) - 1)
            ReDim seriesY(0 To pointCountByBin(binIndex) - 1)
            pointIndex = 0
            For elementId = 0 To NumberOfElement - 1 Step stride
                elementValue = ElementAverageResult(elementId, view.resultId)
                If Not P1IsFiniteValue(elementValue) Then elementValue = 0#
                If P1ContourBinIndex(elementValue, minValue, maxValue) = binIndex Then
                    For pointNo = 0 To 8
                        If pointNo < 8 Then
                            sourceNode = sourceOrder(pointNo)
                        Else
                            sourceNode = sourceOrder(0)
                        End If
                        xValue = Elem(elementId).x(sourceNode)
                        yValue = Elem(elementId).y(sourceNode)
                        ' 応力コンターは元形状へ重ねる。
                        P1ProjectPointCached view, xValue, yValue, 0#, pointX, pointY
                        seriesX(pointIndex) = pointX
                        seriesY(pointIndex) = -pointY
                        pointIndex = pointIndex + 1
                    Next pointNo
                    seriesX(pointIndex) = CVErr(2042)
                    seriesY(pointIndex) = CVErr(2042)
                    pointIndex = pointIndex + 1
                End If
            Next elementId
            P1AddAggregatedChartSeries chart, seriesX, seriesY, pointIndex, _
                                       P1ColorForValue(minValue + (CDbl(binIndex) + 0.5) * (maxValue - minValue) / CDbl(P1_CONTOUR_BIN_COUNT), minValue, maxValue)
        End If
    Next binIndex
    P1ConfigureAggregatedChart chartObject, chart, minX, maxX, minY, maxY
End Sub

Private Function P1ContourSubdiv(ByVal elementCount As Long) As Long
    'Excelの図形描画は要素数に比例して急激に遅くなるため、現在は
    'Q8要素単位の1セル表示を標準とし、要素strideでLODを制御する。
    P1ContourSubdiv = 1
End Function

Private Function P1IsFiniteValue(ByVal value As Double) As Boolean
    P1IsFiniteValue = False
    If value <> value Then Exit Function
    If Abs(value) > 1E+300 Then Exit Function
    P1IsFiniteValue = True
End Function

Private Function P1ContourBinIndex(ByVal value As Double, ByVal minValue As Double, ByVal maxValue As Double) As Long
    Dim ratio As Double, indexValue As Long
    If Not P1IsFiniteValue(value) Then value = 0#
    If maxValue <= minValue Then
        ratio = 0.5
    Else
        ratio = (value - minValue) / (maxValue - minValue)
    End If
    If ratio < 0# Then ratio = 0#
    If ratio > 1# Then ratio = 1#
    indexValue = CLng(Int(ratio * CDbl(P1_CONTOUR_BIN_COUNT)))
    If indexValue < 0 Then indexValue = 0
    If indexValue >= P1_CONTOUR_BIN_COUNT Then indexValue = P1_CONTOUR_BIN_COUNT - 1
    P1ContourBinIndex = indexValue
End Function

Private Sub P1GetNodeDisplacement(ByVal nodeId As Long, ByRef uxValue As Double, ByRef uyValue As Double)
    If nodeId < 0 Or nodeId >= NumberOfFreeNode Then Err.Raise vbObjectError + 3420, "CONTROL.P1GetNodeDisplacement", "Viewerの節点番号が範囲外です。"
    uxValue = TDisp(2 * nodeId)
    uyValue = TDisp(2 * nodeId + 1)
End Sub

Private Sub P1DrawElementShape(ByRef ws As Worksheet, ByRef view As P1_ViewState, ByVal elementId As Long, ByVal deformed As Boolean, ByVal prefix As String)
    Dim pointX(0 To 7) As Double, pointY(0 To 7) As Double
    Dim sourceNode(0 To 7) As Long, pointNo As Long, nodeId As Long, useCount As Long
    Dim uxValue As Double, uyValue As Double, xValue As Double, yValue As Double
    If Elem(elementId).IsJoint Or Elem(elementId).node(6) < 0 Then
        sourceNode(0) = 0: sourceNode(1) = 2: sourceNode(2) = 1
        sourceNode(3) = 4: sourceNode(4) = 5: sourceNode(5) = 3
        useCount = 6
    Else
        sourceNode(0) = 0: sourceNode(1) = 4: sourceNode(2) = 1: sourceNode(3) = 5
        sourceNode(4) = 2: sourceNode(5) = 6: sourceNode(6) = 3: sourceNode(7) = 7
        useCount = 8
    End If
    For pointNo = 0 To useCount - 1
        If Elem(elementId).node(sourceNode(pointNo)) < 0 Then
            xValue = Elem(elementId).x(sourceNode(pointNo))
            yValue = Elem(elementId).y(sourceNode(pointNo))
        Else
            xValue = Elem(elementId).x(sourceNode(pointNo))
            yValue = Elem(elementId).y(sourceNode(pointNo))
            If deformed Then
                nodeId = Elem(elementId).node(sourceNode(pointNo))
                P1GetNodeDisplacement nodeId, uxValue, uyValue
                xValue = xValue + view.DeformedScale * uxValue
                yValue = yValue + view.DeformedScale * uyValue
            End If
        End If
        P1ProjectPointCached view, xValue, yValue, 0#, pointX(pointNo), pointY(pointNo)
    Next pointNo
    P1AddPolygonShape ws, prefix & CStr(elementId + 1), 0, False, useCount, pointX, pointY
End Sub

Private Sub P1DrawNodeLabels(ByRef ws As Worksheet, ByRef view As P1_ViewState)
    Dim nodeId As Long, originalNode As Long, pointX As Double, pointY As Double, labelShape As Object
    If Not view.ShowNodeLabels Then Exit Sub
    If view.DetailMode <> "DETAIL" Then Exit Sub
    If NumberOfFreeNode > 500 Then Exit Sub
    For nodeId = 0 To NumberOfFreeNode - 1
        'NodalX/NodalYはP1結果作成後の配列なので、計算前は入力節点座標を使う。
        originalNode = P6GetOriginalFreeNode(nodeId)
        P1ProjectPointCached view, x(originalNode), y(originalNode), 0#, pointX, pointY
        Set labelShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, pointX - 8#, pointY - 8#, 36#, 16#)
        labelShape.name = "P1Node_" & CStr(nodeId + 1)
        labelShape.fill.Visible = msoFalse
        labelShape.line.Visible = msoFalse
        labelShape.TextFrame.Characters.text = CStr(originalNode + 1)
    Next nodeId
End Sub

Private Sub P1DrawOriginalLayer(ByRef ws As Worksheet, ByRef view As P1_ViewState, ByVal clearLayer As Boolean)
    Dim elementId As Long, stride As Long
    If clearLayer Then P1ClearViewerLayers
    If view.DetailMode <> "DETAIL" Then
        P1AddAggregatedMeshChart ws, view, False, "P1Agg_E", RGB(0, 0, 0)
        Exit Sub
    End If
    stride = P1ElementStride(NumberOfElement, 1)
    For elementId = 0 To NumberOfElement - 1 Step stride
        P1DrawElementShape ws, view, elementId, False, "P1E_"
    Next elementId
    P1DrawNodeLabels ws, view
End Sub

Private Function P1InterpolateElementResult(ByVal elementId As Long, ByVal resultId As Long, ByVal xi As Double, ByVal eta As Double) As Double
    Dim shapeValues(0 To 3) As Double, cornerNo As Long, nodeId As Long
    P1CornerShape xi, eta, shapeValues
    For cornerNo = 0 To 3
        nodeId = Elem(elementId).node(cornerNo)
        P1InterpolateElementResult = P1InterpolateElementResult + shapeValues(cornerNo) * NodalResult(nodeId, resultId)
    Next cornerNo
End Function

Private Function P1ColorForValue(ByVal value As Double, ByVal minValue As Double, ByVal maxValue As Double) As Long
    Dim ratio As Double, redValue As Long, greenValue As Long, blueValue As Long
    If maxValue <= minValue Then
        ratio = 0.5
    Else
        ratio = (value - minValue) / (maxValue - minValue)
    End If
    If ratio < 0# Then ratio = 0#
    If ratio > 1# Then ratio = 1#
    If ratio <= 0.5 Then
        redValue = 0
        greenValue = CLng(255# * ratio * 2#)
        blueValue = 255
    Else
        redValue = CLng(255# * (ratio - 0.5) * 2#)
        greenValue = CLng(255# * (1# - ratio) * 2#)
        blueValue = 0
    End If
    P1ColorForValue = RGB(redValue, greenValue, blueValue)
End Function

Private Sub P1ComputeRange(ByRef view As P1_ViewState, ByRef minValue As Double, ByRef maxValue As Double, ByRef minNode As Long, ByRef maxNode As Long)
    Dim nodeId As Long, elementId As Long, gaussNo As Long
    Dim currentValue As Double, boundValue As Double, validCount As Long
    minValue = 1E+308: maxValue = -1E+308: minNode = -1: maxNode = -1
    validCount = 0
    If view.ResultSource = "RAW_GAUSS" Then
        For elementId = 0 To NumberOfElement - 1
            For gaussNo = 0 To 3
                currentValue = GaussResult(elementId, gaussNo, view.resultId)
                If P1IsFiniteValue(currentValue) Then
                    validCount = validCount + 1
                    If currentValue < minValue Then minValue = currentValue: minNode = elementId: maxNode = gaussNo
                    If currentValue > maxValue Then maxValue = currentValue
                End If
            Next gaussNo
        Next elementId
    ElseIf view.ResultSource = "ELEMENT" Then
        For elementId = 0 To NumberOfElement - 1
            currentValue = ElementAverageResult(elementId, view.resultId)
            If P1IsFiniteValue(currentValue) Then
                validCount = validCount + 1
                If currentValue < minValue Then minValue = currentValue: minNode = elementId
                If currentValue > maxValue Then maxValue = currentValue: maxNode = elementId
            End If
        Next elementId
    Else
        For nodeId = 0 To NumberOfFreeNode - 1
            currentValue = NodalResult(nodeId, view.resultId)
            If P1IsFiniteValue(currentValue) Then
                validCount = validCount + 1
                If currentValue < minValue Then minValue = currentValue: minNode = nodeId
                If currentValue > maxValue Then maxValue = currentValue: maxNode = nodeId
            End If
        Next nodeId
    End If
    If validCount = 0 Then
        minValue = -1#: maxValue = 1#: minNode = 0: maxNode = 0
        Exit Sub
    End If
    If view.scaleMode = "SYMMETRIC" Then
        boundValue = Abs(minValue)
        If Abs(maxValue) > boundValue Then boundValue = Abs(maxValue)
        minValue = -boundValue: maxValue = boundValue
    End If
    If maxValue <= minValue Then
        If Abs(maxValue) <= 1E-30 Then
            minValue = -1#: maxValue = 1#
        Else
            boundValue = Abs(maxValue) * 0.01
            If boundValue <= 1E-30 Then boundValue = 1#
            minValue = minValue - boundValue: maxValue = maxValue + boundValue
        End If
    End If
End Sub

Private Sub P1DrawContourLayerFast(ByRef ws As Worksheet, ByRef view As P1_ViewState, ByVal minValue As Double, ByVal maxValue As Double)
    Dim elementId As Long, stride As Long, elementValue As Double
    Dim shapeItem As Object
    If view.DetailMode <> "DETAIL" Then
        If view.ResultSource = "RAW_GAUSS" Then
            P1DrawAggregatedGaussLayer ws, view, minValue, maxValue
        Else
            P1DrawAggregatedContourLayer ws, view, minValue, maxValue
        End If
        Exit Sub
    End If
    If view.ResultSource = "RAW_GAUSS" Then
        P1DrawGaussPointLayer ws, view, minValue, maxValue
        Exit Sub
    End If
    stride = P1ElementStride(NumberOfElement, 1)
    For elementId = 0 To NumberOfElement - 1 Step stride
        elementValue = ElementAverageResult(elementId, view.resultId)
        If view.ResultSource = "ELEMENT" Then
            P1DrawFilledElement ws, view, elementId, elementValue, minValue, maxValue, "P1C_"
        Else
            P1DrawElementShape ws, view, elementId, False, "P1C_"
            Set shapeItem = ws.Shapes.item("P1C_" & CStr(elementId + 1))
            shapeItem.line.ForeColor.RGB = P1ColorForValue(elementValue, minValue, maxValue)
        End If
    Next elementId
End Sub

Private Sub P1DrawFilledElement(ByRef ws As Worksheet, ByRef view As P1_ViewState, ByVal elementId As Long, ByVal elementValue As Double, ByVal minValue As Double, ByVal maxValue As Double, ByVal prefix As String)
    Dim pointX(0 To 7) As Double, pointY(0 To 7) As Double
    Dim sourceNode(0 To 7) As Long, pointNo As Long
    Dim xValue As Double, yValue As Double
    sourceNode(0) = 0: sourceNode(1) = 4: sourceNode(2) = 1: sourceNode(3) = 5
    sourceNode(4) = 2: sourceNode(5) = 6: sourceNode(6) = 3: sourceNode(7) = 7
    For pointNo = 0 To 7
        xValue = Elem(elementId).x(sourceNode(pointNo))
        yValue = Elem(elementId).y(sourceNode(pointNo))
        P1ProjectPointCached view, xValue, yValue, 0#, pointX(pointNo), pointY(pointNo)
    Next pointNo
    P1AddPolygonShape ws, prefix & CStr(elementId + 1), P1ColorForValue(elementValue, minValue, maxValue), True, 8, pointX, pointY
End Sub

Private Sub P1DrawGaussPointLayer(ByRef ws As Worksheet, ByRef view As P1_ViewState, ByVal minValue As Double, ByVal maxValue As Double)
    Dim elementId As Long, gaussNo As Long, stride As Long
    Dim xValue As Double, yValue As Double, pointX As Double, pointY As Double
    Dim gaussValue As Double, markerSize As Double
    Dim shapeItem As Object
    stride = P1ElementStride(NumberOfElement, 4)
    markerSize = 6#
    For elementId = 0 To NumberOfElement - 1 Step stride
        For gaussNo = 0 To 3
            FEMGaussPhysicalXY elementId, gaussNo, xValue, yValue
            P1ProjectPointCached view, xValue, yValue, 0#, pointX, pointY
            gaussValue = GaussResult(elementId, gaussNo, view.resultId)
            Set shapeItem = ws.Shapes.AddShape(msoShapeOval, pointX - markerSize * 0.5, pointY - markerSize * 0.5, markerSize, markerSize)
            shapeItem.name = "P1G_" & CStr(elementId + 1) & "_" & CStr(gaussNo)
            shapeItem.fill.ForeColor.RGB = P1ColorForValue(gaussValue, minValue, maxValue)
            shapeItem.fill.Visible = msoTrue
            shapeItem.line.Visible = msoFalse
        Next gaussNo
    Next elementId
End Sub

Private Sub P1DrawAggregatedGaussLayer(ByRef ws As Worksheet, ByRef view As P1_ViewState, ByVal minValue As Double, ByVal maxValue As Double)
    Dim elementId As Long, gaussNo As Long, stride As Long, binIndex As Long, pointIndex As Long
    Dim xValue As Double, yValue As Double, pointX As Double, pointY As Double, gaussValue As Double
    Dim pointCountByBin(0 To P1_CONTOUR_BIN_COUNT - 1) As Long
    Dim seriesX() As Variant, seriesY() As Variant
    Dim chartObject As Object, chart As Object
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim firstPoint As Boolean
    stride = P1ElementStride(NumberOfElement, 4)
    firstPoint = True
    For elementId = 0 To NumberOfElement - 1 Step stride
        For gaussNo = 0 To 3
            gaussValue = GaussResult(elementId, gaussNo, view.resultId)
            If Not P1IsFiniteValue(gaussValue) Then gaussValue = 0#
            binIndex = P1ContourBinIndex(gaussValue, minValue, maxValue)
            pointCountByBin(binIndex) = pointCountByBin(binIndex) + 1
            FEMGaussPhysicalXY elementId, gaussNo, xValue, yValue
            P1ProjectPointCached view, xValue, yValue, 0#, pointX, pointY
            If firstPoint Then
                minX = pointX: maxX = pointX: minY = pointY: maxY = pointY
                firstPoint = False
            Else
                If pointX < minX Then minX = pointX
                If pointX > maxX Then maxX = pointX
                If pointY < minY Then minY = pointY
                If pointY > maxY Then maxY = pointY
            End If
        Next gaussNo
    Next elementId
    If firstPoint Then Exit Sub
    Set chartObject = ws.ChartObjects.Add(minX - 8#, minY - 8#, maxX - minX + 16#, maxY - minY + 16#)
    chartObject.name = "P1Agg_G"
    Set chart = chartObject.chart
    chart.ChartType = 75
    For binIndex = 0 To P1_CONTOUR_BIN_COUNT - 1
        If pointCountByBin(binIndex) > 0 Then
            ReDim seriesX(0 To pointCountByBin(binIndex) - 1)
            ReDim seriesY(0 To pointCountByBin(binIndex) - 1)
            pointIndex = 0
            For elementId = 0 To NumberOfElement - 1 Step stride
                For gaussNo = 0 To 3
                    gaussValue = GaussResult(elementId, gaussNo, view.resultId)
                    If Not P1IsFiniteValue(gaussValue) Then gaussValue = 0#
                    If P1ContourBinIndex(gaussValue, minValue, maxValue) = binIndex Then
                        FEMGaussPhysicalXY elementId, gaussNo, xValue, yValue
                        P1ProjectPointCached view, xValue, yValue, 0#, pointX, pointY
                        seriesX(pointIndex) = pointX
                        seriesY(pointIndex) = -pointY
                        pointIndex = pointIndex + 1
                    End If
                Next gaussNo
            Next elementId
            P1AddAggregatedScatterSeries chart, seriesX, seriesY, pointIndex, _
                P1ColorForValue(minValue + (CDbl(binIndex) + 0.5) * (maxValue - minValue) / CDbl(P1_CONTOUR_BIN_COUNT), minValue, maxValue)
        End If
    Next binIndex
    P1ConfigureAggregatedChart chartObject, chart, minX, maxX, minY, maxY
End Sub

Private Sub P1AddLegend(ByRef ws As Worksheet, ByRef view As P1_ViewState, ByVal titleText As String, ByVal resultId As Long, ByVal minValue As Double, ByVal maxValue As Double, ByVal minNode As Long, ByVal maxNode As Long, ByVal sourceText As String)
    Dim legendShape As Object, legendText As String
    If minNode < 0 Or minNode >= NumberOfFreeNode Then minNode = 0
    If maxNode < 0 Or maxNode >= NumberOfFreeNode Then maxNode = 0
    legendText = titleText & "  [" & P1ResultUnit(resultId) & "]" & vbCrLf
    legendText = legendText & "scale=" & view.scaleMode & ", source=" & sourceText & vbCrLf
    legendText = legendText & "min=" & Format$(minValue, "0.000E+00") & "  max=" & Format$(maxValue, "0.000E+00")
    Set legendShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 10#, 10#, 330#, 58#)
    legendShape.name = "P1Legend_" & CStr(resultId)
    legendShape.fill.Visible = msoFalse
    legendShape.line.Visible = msoFalse
    legendShape.TextFrame.Characters.text = legendText
End Sub

Public Function P1GetLastViewerError() As String
    P1GetLastViewerError = P1ViewerLastError
End Function

Public Sub P1DrawOriginalViewer()
    Dim ws As Worksheet, view As P1_ViewState
    On Error GoTo P1OriginalViewerError
    P1ViewerLastError = ""
    '変形前の図形表示は解析結果を必要としない。
    '初回表示またはメッシュ変更後だけ、材料検証を伴わない表示専用Q8読込を行う。
    If NumberOfNode <= 0 Or NumberOfElement <= 0 Or NumberOfFreeNode <= 0 Then
        P1LoadViewerMeshFromSheets
    End If
    Set ws = ThisWorkbook.Worksheets("図")
    P1PrepareFigureViewport ws
    P1ReadViewState view
    P1ClearViewerLayers
    P1DrawOriginalLayer ws, view, False
    Exit Sub
P1OriginalViewerError:
    P1ViewerLastError = CStr(Err.Number) & " " & Err.Description
    If Not SuppressUserMessages Then MsgBox "図形表示に失敗しました。" & vbCrLf & P1ViewerLastError, vbExclamation
End Sub

Public Sub P1DrawDeformedViewer()
    Dim ws As Worksheet, view As P1_ViewState, elementId As Long, stride As Long
    If Not P1EnsureViewerReady("変形後メッシュ表示") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("図")
    P1PrepareFigureViewport ws
    P1ReadViewState view
    P1ClearViewerLayers
    P1DrawOriginalLayer ws, view, False
    If view.DetailMode <> "DETAIL" Then
        P1AddAggregatedMeshChart ws, view, True, "P1Agg_D", RGB(192, 0, 0)
        Exit Sub
    End If
    stride = P1ElementStride(NumberOfElement, 1)
    For elementId = 0 To NumberOfElement - 1 Step stride
        P1DrawElementShape ws, view, elementId, True, "P1D_"
    Next elementId
End Sub

Private Function P1ResultSourceLabel(ByRef view As P1_ViewState) As String
    If view.ResultSource = "RAW_GAUSS" Then
        P1ResultSourceLabel = "RAW_GAUSS (Gauss点)"
    ElseIf view.ResultSource = "ELEMENT" Then
        P1ResultSourceLabel = "ELEMENT (要素平均)"
    Else
        P1ResultSourceLabel = "SMOOTHED (節点平滑)"
    End If
End Function

Public Sub P1DrawContourViewer()
    Dim ws As Worksheet, view As P1_ViewState
    Dim minValue As Double, maxValue As Double, minNode As Long, maxNode As Long
    On Error GoTo P1ContourViewerError
    P1ViewerLastError = ""
    If Not P1EnsureViewerReady("コンター表示") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("図")
    P1PrepareFigureViewport ws
    P1ReadViewState view
    If view.resultId < 0 Or view.resultId >= P1_RESULT_COUNT Then
        If Not SuppressUserMessages Then MsgBox "結果IDが範囲外です。", vbExclamation
        Exit Sub
    End If
    If Not P1ResultAvailable(view.resultId) Then
        If Not SuppressUserMessages Then MsgBox "選択した結果は現行2D要素では未対応です。", vbExclamation
        Exit Sub
    End If
    P1ComputeRange view, minValue, maxValue, minNode, maxNode
    P1ClearViewerLayers
    P1DrawOriginalLayer ws, view, False
    ' 元形状を先に描き、応力コンターを後から重ねる。逆順では黒い元形状の
    ' Chart系列が色付きコンターを覆い、応力表示が一色に見える。
    P1DrawContourLayerFast ws, view, minValue, maxValue
    P1AddLegend ws, view, P1ResultLabel(view.resultId), view.resultId, minValue, maxValue, minNode, maxNode, P1ResultSourceLabel(view)
    Exit Sub
P1ContourViewerError:
    P1ViewerLastError = CStr(Err.Number) & " " & Err.Description
    If Not SuppressUserMessages Then MsgBox "コンター表示に失敗しました。" & vbCrLf & P1ViewerLastError, vbExclamation
End Sub





