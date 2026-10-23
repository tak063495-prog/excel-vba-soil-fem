Option Explicit

Public FEMViewerLastSvg As String
Public FEMViewerCanvasW As Double, FEMViewerCanvasH As Double
Public FEMViewerLastError As String
Public FEMViewerElementCount As Long
Private viewerBusy As Boolean
Private decimalMark As String
Private Const DIRTY_NAME As String = "_FEMViewerResultsDirty"

Public Sub FEMViewerMarkDirty()
    SetResultDirty True
End Sub

Public Sub FEMViewerMarkClean()
    SetResultDirty False
End Sub

Private Sub SetResultDirty(ByVal dirty As Boolean)
    Dim item As name, formula As String
    formula = "=FALSE"
    If dirty Then formula = "=TRUE"
    On Error Resume Next
    Set item = ThisWorkbook.names(DIRTY_NAME)
    On Error GoTo 0
    If item Is Nothing Then
        ThisWorkbook.names.Add name:=DIRTY_NAME, RefersTo:=formula, Visible:=False
    Else
        item.RefersTo = formula
    End If
End Sub

Private Function ResultsDirty() As Boolean
    On Error Resume Next
    ResultsDirty = (ThisWorkbook.names(DIRTY_NAME).RefersTo = "=TRUE")
    On Error GoTo 0
End Function

Public Sub FEMViewAll()
    DrawView "ALL"
End Sub

Public Sub FEMViewInitial()
    DrawView "INITIAL"
End Sub

Public Sub FEMViewResult()
    DrawView "RESULT"
End Sub

Private Function ReadTable(ByVal sheetName As String, ByVal columns As Long) As Variant
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row
    If lastRow < 2 Then lastRow = 2
    ReadTable = ws.range(ws.cells(1, 1), ws.cells(lastRow, columns)).value2
End Function

Private Function number(ByVal value As Variant, ByVal field As String) As Double
    If IsError(value) Then Err.Raise vbObjectError + 3700, "FEMViewer", field & "がエラーです。"
    If IsEmpty(value) Then Err.Raise vbObjectError + 3700, "FEMViewer", field & "が空欄です。"
    If Not IsNumeric(value) Then Err.Raise vbObjectError + 3700, "FEMViewer", field & "が数値ではありません。"
    number = CDbl(value)
    If Abs(number) > 1E+100 Then Err.Raise vbObjectError + 3700, "FEMViewer", field & "が表示範囲外です。"
End Function

Private Function id(ByVal value As Variant) As Long
    Dim n As Double
    n = number(value, "番号")
    If n < 1# Or n > 2147483647# Or n <> Fix(n) Then Err.Raise vbObjectError + 3701, "FEMViewer", "番号が正の整数ではありません。"
    id = CLng(n)
End Function

Private Function IsOn(ByVal value As Variant) As Boolean
    Dim token As String
    IsOn = True
    If IsEmpty(value) Then Exit Function
    If IsNumeric(value) Then
        IsOn = (Abs(CDbl(value)) >= 0.5)
        Exit Function
    End If
    token = UCase$(Trim$(CStr(value)))
    Select Case token
        Case "OFF", "NO", "N", "しない", "無効", "FALSE": IsOn = False
    End Select
End Function

Private Function Xml(ByVal text As String) As String
    text = Replace$(text, "&", "&amp;")
    text = Replace$(text, "<", "&lt;")
    text = Replace$(text, ">", "&gt;")
    Xml = Replace$(text, """", "&quot;")
End Function

Private Function num(ByVal value As Double) As String
    num = Replace$(Format$(value, "0.000"), decimalMark, ".")
End Function

Private Function SvgText(ByVal x As Double, ByVal y As Double, ByVal text As String, Optional ByVal size As Long = 12) As String
    SvgText = "<text x=""" & num(x) & """ y=""" & num(y) & """ font-size=""" & CStr(size) & """ fill=""#263444"">" & Xml(text) & "</text>"
End Function

Private Function Color(ByVal value As Double, ByVal low As Double, ByVal high As Double) As String
    Dim t As Double, r As Long, g As Long, b As Long
    t = 0.5
    If high > low Then t = (value - low) / (high - low)
    If t < 0# Then t = 0#
    If t > 1# Then t = 1#
    If t <= 0.5 Then
        r = CLng(510# * t): g = r: b = 255
    Else
        r = 255: g = CLng(510# * (1# - t)): b = g
    End If
    Color = "#" & right$("0" & Hex$(r), 2) & right$("0" & Hex$(g), 2) & right$("0" & Hex$(b), 2)
End Function

Private Function LatestStage(ByRef rows As Variant) As Long
    Dim r As Long, n As Long
    For r = 2 To UBound(rows, 1)
        If Not IsEmpty(rows(r, 1)) Then
            n = id(rows(r, 1))
            LatestStage = n
        End If
    Next r
End Function

Public Function FEMViewerResultUnit(ByVal selection As Long) As String
    Dim item As Long
    If selection = 71 Then FEMViewerResultUnit = "状態": Exit Function
    If selection >= 69 Then FEMViewerResultUnit = "無次元": Exit Function
    item = (selection - 1) Mod 17 + 1
    If item = 7 Then
        FEMViewerResultUnit = "deg"
    ElseIf item >= 11 And item <= 16 Then
        FEMViewerResultUnit = "無次元"
    Else
        FEMViewerResultUnit = "force/area"
    End If
End Function
Private Function SavedElementValues(ByRef rows As Variant, ByVal selection As Long, ByVal Stage As Long, ByVal active As Object, ByVal jointElements As Object) As Object
    Dim values As Object, r As Long, e As Long, column As Long
    Set values = CreateObject("Scripting.Dictionary")
    If selection < 1 Or selection > 68 Then Err.Raise vbObjectError + 3703, "FEMViewer", "色付けは1～68の整数を指定してください（結果要素 E列～BT列）。"
    column = selection + 4
    If Len(Trim$(CStr(rows(1, column)))) = 0 Then Err.Raise vbObjectError + 3703, "FEMViewer", "選択列の見出しがありません。再解析してください。"
    For r = 2 To UBound(rows, 1)
        If Not IsEmpty(rows(r, 1)) Then
            If CLng(rows(r, 1)) = Stage Then
                e = id(rows(r, 2))
                If active.Exists(e) And Not jointElements.Exists(e) Then
                    If values.Exists(e) Then Err.Raise vbObjectError + 3704, "FEMViewer", "同じステージ・要素の結果が重複しています。要素=" & e
                    If IsEmpty(rows(r, column)) Then Err.Raise vbObjectError + 3704, "FEMViewer", "選択列の結果がありません。要素=" & e
                    values.Add e, number(rows(r, column), "結果要素の選択列")
                End If
            End If
        End If
    Next r
    Set SavedElementValues = values
End Function
Public Sub FEMViewerSetupSelector()
    FEMResultOptionsInstall
End Sub

Private Sub DrawView(ByVal mode As String)
    Dim nodes As Variant, elements As Variant, materials As Variant, stages As Variant, settings As Variant
    Dim nodeResults As Variant, elementResults As Variant, nodeMap As Object, materialMap As Object
    Dim plans As Object, displacements As Object, active As Object, values As Object, config As Object, finalMaterials As Object
    Dim r As Long, j As Long, n As Long, e As Long, mat As Long, flag As Long, Stage As Long, count As Long
    Dim selected() As Boolean, px() As Double, py() As Double, elementIds As Object, key As Variant
    Dim jointMaterials As Object, jointElements As Object, jointRow() As Boolean
    Dim visibleNodes As Object, supportX As Boolean, supportY As Boolean, supportLabel As String
    Dim nodeLimit As Long, continuumCount As Long, kind As String
    Dim stageResults As Variant, geometryMoved As Boolean, resultX As Double, resultY As Double
    Dim node As Variant, displacement As Variant, current As Variant, order As Variant
    Dim low As Double, high As Double, value As Double, bound As Double, dscale As Double, drawScale As Double
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double, sx As Double, sy As Double
    Dim originX As Double, originY As Double, x As Double, y As Double, firstPoint As Boolean
    Dim increments As Object, overlayCells As Object, entry As Variant, overlayLegend As String, overlaySvg As String
    Dim shearOn As Boolean, crossOn As Boolean, overlayA As Boolean, incrementPaint As Boolean
    Dim statePaint As Boolean, cumPaint As Boolean, plasticStates As Object, stateCount(0 To 3) As Long, stateNames As Variant
    Dim vectorScale As Double, baseResultId As Long, incrementCaption As String
    Dim resultId As Long, fill As String, stroke As String, points As String, title As String, subtitle As String
    Dim svg As String, parts() As String, partCount As Long, labels As Object, showLabels As Boolean
    Dim ws As Worksheet, picture As Object, oldScreen As Boolean, oldEvents As Boolean, oldStatus As Variant
    Dim oldDisplayStatus As Boolean, stateSaved As Boolean, errorText As String, lastRow As Long, canvasW As Double, canvasH As Double
    Dim viewPhase As String, errorNumber As Long
    If AnalysisRunning Or viewerBusy Then Exit Sub
    viewerBusy = True
    FEMViewerLastSvg = ""
    FEMViewerLastError = "": FEMViewerElementCount = 0
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    oldStatus = Application.StatusBar: oldDisplayStatus = Application.DisplayStatusBar
    stateSaved = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "メッシュ表示を準備しています"
    viewPhase = "表示設定の読込"
    decimalMark = Application.International(xlDecimalSeparator)
    nodes = ReadTable("節点データ", 5)
    elements = ReadTable("要素データ", 10)
    materials = ReadTable("材料データ", 13)
    stages = ReadTable("ステージ", 5)
    ' The viewer reads sheet values directly, including in manual calculation mode.
    If FEMUiInstalled() Then ThisWorkbook.Worksheets("出力・表示").Calculate
    ThisWorkbook.Worksheets("設定").Calculate
    settings = ReadTable("設定", 5)
    Set nodeMap = CreateObject("Scripting.Dictionary")
    Set materialMap = CreateObject("Scripting.Dictionary")
    Set jointMaterials = CreateObject("Scripting.Dictionary")
    Set jointElements = CreateObject("Scripting.Dictionary")
    Set plans = CreateObject("Scripting.Dictionary")
    Set finalMaterials = CreateObject("Scripting.Dictionary")
    Set elementIds = CreateObject("Scripting.Dictionary")
    Set config = CreateObject("Scripting.Dictionary")
    Set labels = CreateObject("Scripting.Dictionary")
    Set visibleNodes = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(settings, 1)
        If Len(CStr(settings(r, 5))) > 0 Then config(UCase$(CStr(settings(r, 5)))) = settings(r, 3)
    Next r
    For r = 2 To UBound(nodes, 1)
        If Not IsEmpty(nodes(r, 1)) Then
            n = id(nodes(r, 1))
            If nodeMap.Exists(n) Then Err.Raise vbObjectError + 3705, "FEMViewer", "節点番号が重複しています。"
            nodeMap.Add n, Array(number(nodes(r, 2), "節点X"), number(nodes(r, 3), "節点Y"))
        End If
    Next r
    If config.Exists("VIEW_DETAIL_MODE") Then showLabels = (UCase$(Trim$(CStr(config("VIEW_DETAIL_MODE")))) = "DETAIL")
    For r = 2 To UBound(materials, 1)
        If Not IsEmpty(materials(r, 1)) Then
            mat = id(materials(r, 1))
            materialMap(mat) = IsOn(materials(r, 10))
            kind = UCase$(Trim$(CStr(materials(r, 11))))
            If kind = "JOINT" Or kind = "接合" Or kind = "IF" Or kind = "INTERFACE" Then jointMaterials(mat) = True
        End If
    Next r
    For Each key In materialMap.keys
        finalMaterials(key) = materialMap(key)
    Next key
    Dim framePrefix As Long, planIndex As Long
    framePrefix = -1
    If mode = "RESULT" Then framePrefix = FSViewPrefix()
    For r = 2 To UBound(stages, 1)
        If Len(Trim$(CStr(stages(r, 2)))) > 0 Then planIndex = planIndex + 1
        If framePrefix >= 0 And planIndex > framePrefix Then Exit For
        If Not IsNumeric(stages(r, 5)) Then GoTo NextPlan
        If CDbl(stages(r, 5)) <> 0# Then
            flag = 0
            Select Case UCase$(Trim$(CStr(stages(r, 2))))
                Case "BIRTH", "誕生", "盛土": flag = 1
                Case "DEATH", "死滅", "掘削", "KILL": flag = 2
            End Select
            If flag <> 0 Then
                mat = id(stages(r, 3))
                finalMaterials(mat) = (flag = 1)
                If plans.Exists(mat) Then flag = flag Or CLng(plans(mat))
                plans(mat) = flag
            End If
        End If
NextPlan:
    Next r
    ReDim jointRow(2 To UBound(elements, 1))
    For r = 2 To UBound(elements, 1)
        If Not IsEmpty(elements(r, 1)) Then
            jointRow(r) = jointMaterials.Exists(id(elements(r, 10)))
            If jointRow(r) Then jointElements(id(elements(r, 1))) = True
        End If
    Next r
    dscale = 0#: Stage = 0
    viewPhase = "保存結果と入力の照合"
    If mode = "RESULT" Then
        If ResultsDirty() Then Err.Raise vbObjectError + 3706, "FEMViewer", "入力変更後のため、再解析してから結果表示してください。"
        If CStr(ThisWorkbook.Worksheets("診断").cells(2, 2).value2) <> "PASS" Then Err.Raise vbObjectError + 3707, "FEMViewer", "解析成功後に結果表示してください。"
        nodeResults = ReadTable("結果節点", 10)
        elementResults = ReadTable("結果要素", 72)
        Stage = LatestStage(nodeResults)
        If Stage = 0 Then Err.Raise vbObjectError + 3708, "FEMViewer", "保存済みの節点結果がありません。"
        ' RESET_STRESS changes the reference geometry without editing the input mesh.
        ' Use saved stage coordinates only after a recorded successful reset.
        stageResults = ReadTable("ステージ結果", 4)
        ' Stage IDs are user-editable and need not follow execution order.
        ' Locate the saved final result in the successful history first, then
        ' use row order to determine whether RESET_STRESS preceded it.
        Dim finalHistoryRow As Long
        finalHistoryRow = 0
        For r = 2 To UBound(stageResults, 1)
            If IsNumeric(stageResults(r, 1)) Then
                If CLng(stageResults(r, 1)) = Stage Then
                    If UCase$(Trim$(CStr(stageResults(r, 4)))) = "PASS" Then finalHistoryRow = r
                End If
            End If
        Next r
        If finalHistoryRow > 0 Then
            For r = 2 To finalHistoryRow
                If UCase$(Trim$(CStr(stageResults(r, 2)))) = "RESET_STRESS" And CStr(stageResults(r, 4)) = "PASS" Then geometryMoved = True
            Next r
        End If
        Set active = CreateObject("Scripting.Dictionary")
        Set displacements = CreateObject("Scripting.Dictionary")
        For r = 2 To UBound(elementResults, 1)
            If Not IsEmpty(elementResults(r, 1)) Then
                If CLng(elementResults(r, 1)) = Stage And Not IsEmpty(elementResults(r, 5)) Then active(id(elementResults(r, 2))) = True
            End If
        Next r
        ' INCLUDE出力では無効要素にも数値が残るため、施工手順で復元した最終状態を優先する。
        For r = 2 To UBound(elements, 1)
            If Not IsEmpty(elements(r, 1)) Then
                e = id(elements(r, 1)): mat = id(elements(r, 10))
                If Not finalMaterials.Exists(mat) Then Err.Raise vbObjectError + 3713, "FEMViewer", "最終状態を確認できない材料があります。"
                If Not CBool(finalMaterials(mat)) Then
                    If active.Exists(e) Then active.REMOVE e
                ElseIf Not active.Exists(e) Then
                    Err.Raise vbObjectError + 3717, "FEMViewer", "最終の有効要素の結果がありません。要素=" & CStr(e)
                End If
            End If
        Next r
        For r = 2 To UBound(nodeResults, 1)
            If Not IsEmpty(nodeResults(r, 1)) Then
                If CLng(nodeResults(r, 1)) = Stage Then
                    n = id(nodeResults(r, 2))
                    If Not nodeMap.Exists(n) Then Err.Raise vbObjectError + 3709, "FEMViewer", "結果の節点が現在のメッシュにありません。"
                    node = nodeMap(n)
                    resultX = number(nodeResults(r, 3), "結果X"): resultY = number(nodeResults(r, 4), "結果Y")
                    If geometryMoved Then
                        nodeMap(n) = Array(resultX, resultY)
                    Else
                        If Abs(node(0) - resultX) > 0.00000001 * (1# + Abs(node(0))) Or Abs(node(1) - resultY) > 0.00000001 * (1# + Abs(node(1))) Then Err.Raise vbObjectError + 3710, "FEMViewer", "結果と入力の座標が一致しません。再解析してください。"
                    End If
                    displacements(n) = Array(number(nodeResults(r, 5), "X変位"), number(nodeResults(r, 6), "Y変位"))
                End If
            End If
        Next r
        resultId = 18
        If config.Exists("VIEW_COLOR_RESULT") Then resultId = FEMResultChoiceId(config("VIEW_COLOR_RESULT"))
        If resultId = 0 Then Err.Raise vbObjectError + 3703, , "表示する結果は内容名から選択してください。"
        baseResultId = resultId
        overlayA = True: vectorScale = 1#
        If FEMUiInstalled() Then
            shearOn = FEMShearOverlayOn(): crossOn = FEMPlasticCrossOn()
            overlayA = FEMOverlayModeA()
            If crossOn Then vectorScale = FEMPlasticCrossScale()
        End If
        If shearOn Or crossOn Or resultId = 69 Or resultId = 70 Then
            Set increments = FEMIncrementRead(Stage)
            Set overlayCells = CreateObject("Scripting.Dictionary")
        End If
        If shearOn And Not overlayA Then resultId = 69
        incrementPaint = (resultId = 69 Or resultId = 70)
        statePaint = (resultId = 71): cumPaint = (resultId = 72)
        If statePaint Or cumPaint Then
            Set plasticStates = FEMPlasticStateRead(Stage)
            Set values = CreateObject("Scripting.Dictionary")
            For Each key In active.keys
                e = CLng(key)
                If Not jointElements.Exists(e) Then
                    If Not plasticStates.Exists(e) Then Err.Raise vbObjectError + 3914, , "有効要素の塑性状態がありません。20261022版で再解析してください。"
                    entry = plasticStates(e)
                    If statePaint Then
                        values(e) = CDbl(entry(0)): stateCount(CLng(entry(0)) + 1) = stateCount(CLng(entry(0)) + 1) + 1
                    Else
                        values(e) = CDbl(entry(1))
                    End If
                End If
            Next key
        ElseIf incrementPaint Then
            Set values = CreateObject("Scripting.Dictionary")
            For Each key In active.keys
                e = CLng(key)
                If Not jointElements.Exists(e) Then
                    If Not increments.Exists(e) Then Err.Raise vbObjectError + 3912, , "有効要素の増分結果がありません。再解析してください。"
                    entry = increments(e): values(e) = entry(resultId - 69)
                End If
            Next key
        Else
            Set values = SavedElementValues(elementResults, resultId, Stage, active, jointElements)
        End If
        dscale = 1#
        If config.Exists("VIEW_DISP_SCALE") Then dscale = number(config("VIEW_DISP_SCALE"), "変位倍率")
        If dscale < 0# Then Err.Raise vbObjectError + 3711, "FEMViewer", "変位倍率は0以上で指定してください。"
    End If
    Dim mixedT6Count As Long
    ReDim selected(2 To UBound(elements, 1))
    ReDim px(2 To UBound(elements, 1), 0 To 7)
    ReDim py(2 To UBound(elements, 1), 0 To 7)
    firstPoint = True: low = 1E+100: high = -1E+100
    For r = 2 To UBound(elements, 1)
        If IsEmpty(elements(r, 1)) Then GoTo NextElement
        e = id(elements(r, 1))
        If elementIds.Exists(e) Then Err.Raise vbObjectError + 3712, "FEMViewer", "要素番号が重複しています。"
        elementIds.Add e, True
        selected(r) = True
        If mode = "INITIAL" Then
            mat = id(elements(r, 10))
            If Not materialMap.Exists(mat) Then Err.Raise vbObjectError + 3713, "FEMViewer", "初期状態を確認できない材料があります。材料=" & CStr(mat)
            selected(r) = CBool(materialMap(mat))
        ElseIf mode = "RESULT" Then
            selected(r) = active.Exists(e)
        End If
        If Not selected(r) Then GoTo NextElement
        count = count + 1
        If Not jointRow(r) And ViewerT6Row(elements, r) Then mixedT6Count = mixedT6Count + 1
        nodeLimit = 7
        If jointRow(r) Then nodeLimit = 5 Else continuumCount = continuumCount + 1
        For j = 0 To nodeLimit
            n = id(elements(r, j + 2))
            If Not nodeMap.Exists(n) Then Err.Raise vbObjectError + 3714, "FEMViewer", "接続節点がありません。要素=" & CStr(e)
            node = nodeMap(n): x = node(0): y = node(1)
            visibleNodes(n) = True
            If mode = "RESULT" Then
                If Not displacements.Exists(n) Then Err.Raise vbObjectError + 3715, "FEMViewer", "有効要素の変位がありません。節点=" & CStr(n)
                displacement = displacements(n)
                x = x + dscale * displacement(0): y = y + dscale * displacement(1)
            End If
            px(r, j) = x: py(r, j) = y
            If showLabels And nodeMap.count <= 500 Then labels(n) = Array(x, y)
            If firstPoint Then minX = x: maxX = x: minY = y: maxY = y: firstPoint = False
            If x < minX Then minX = x
            If x > maxX Then maxX = x
            If y < minY Then minY = y
            If y > maxY Then maxY = y
        Next j
        If mode = "RESULT" And Not jointRow(r) Then
            If Not values.Exists(e) Then Err.Raise vbObjectError + 3716, "FEMViewer", "有効要素の応力がありません。要素=" & CStr(e)
            value = values(e)
            If value < low Then low = value
            If value > high Then high = value
        End If
NextElement:
    Next r
    If mode = "RESULT" Then
        If count <> active.count Then Err.Raise vbObjectError + 3717, "FEMViewer", "結果要素と入力要素の件数が一致しません。"
        If continuumCount = 0 Then low = 0#: high = 1#
        If config.Exists("VIEW_SCALE_MODE") Then
            If UCase$(CStr(config("VIEW_SCALE_MODE"))) = "SYMMETRIC" Then
                bound = Abs(low): If Abs(high) > bound Then bound = Abs(high)
                low = -bound: high = bound
            End If
        End If
        If high <= low Then bound = Abs(low) * 0.01: If bound < 0.00000001 Then bound = 1#
        If high <= low Then low = low - bound: high = high + bound
    End If
    drawScale = 1#: originX = 40#: originY = 100#: canvasW = 860#: canvasH = 540#
    If count > 0 Then
        sx = 740# / (maxX - minX + 0.000000000001)
        sy = 370# / (maxY - minY + 0.000000000001)
        drawScale = sx: If sy < drawScale Then drawScale = sy
        originX = 40# - minX * drawScale: originY = 95# + maxY * drawScale
        If mode <> "RESULT" And mixedT6Count > 0 Then originY = originY + 25#
    End If
    If config.Exists("VIEW_SCALE_MODE") Then
        If UCase$(CStr(config("VIEW_SCALE_MODE"))) = "USER" Then
            If config.Exists("VIEW_SCALE") Then drawScale = number(config("VIEW_SCALE"), "表示スケール")
            If config.Exists("VIEW_MIN_X") Then originX = number(config("VIEW_MIN_X"), "作図原点X")
            If config.Exists("VIEW_MAX_Y") Then originY = number(config("VIEW_MAX_Y"), "作図原点Y")
            If drawScale <= 0# Then Err.Raise vbObjectError + 3718, "FEMViewer", "表示スケールは正数で指定してください。"
            If originX + maxX * drawScale + 40# > canvasW Then canvasW = originX + maxX * drawScale + 40#
            If originY - minY * drawScale + 80# > canvasH Then canvasH = originY - minY * drawScale + 80#
        End If
    End If
    If mode = "RESULT" Then
        If shearOn Or crossOn Then
            For r = 2 To UBound(elements, 1)
                If selected(r) And Not jointRow(r) Then
                    e = id(elements(r, 1))
                    If Not increments.Exists(e) Then Err.Raise vbObjectError + 3912, , "有効要素の増分結果がありません。再解析してください。"
                    entry = increments(e)
                    overlayCells(e) = Array(originX + drawScale * px(r, 0), originY - drawScale * py(r, 0), originX + drawScale * px(r, 1), originY - drawScale * py(r, 1), originX + drawScale * px(r, 2), originY - drawScale * py(r, 2), originX + drawScale * px(r, 3), originY - drawScale * py(r, 3), entry(0), entry(2), entry(3), entry(4))
                End If
            Next r
            overlaySvg = FEMIncrementOverlaySvg(overlayCells, shearOn, crossOn, overlayA, vectorScale, overlayLegend)
        End If
        If Not increments Is Nothing Then
            For Each key In increments.keys
                entry = increments(key)
                incrementCaption = "最終収束増分 / Fs=" & Format$(entry(5), "0.0000") & " / 増分=" & CStr(entry(6)) & " / 荷重係数=" & Format$(entry(7), "0.0000")
                Exit For
            Next key
        End If
        If incrementPaint Or cumPaint Then low = 0#
        If statePaint Then low = -1#: high = 2#
    End If
    Select Case mode
        Case "ALL": title = "図形表示（全メッシュ）": subtitle = "全ステージの要素を表示。色は有効なステージの予定。"
        Case "INITIAL": title = "初期表示": subtitle = "初期状態ONのみ。BIRTH待ち（OFF）は表示しません。"
        Case "RESULT": title = "変形・結果表示": subtitle = "最終ステージ " & CStr(Stage) & " / 変位倍率 " & Format$(dscale, "0.###") & " / " & FEMResultChoiceLabel(resultId) & " [" & FEMViewerResultUnit(resultId) & "]"
    End Select
    If jointElements.count > 0 Then subtitle = subtitle & " / 接合は紫線：数値は接合結果"
    ReDim parts(0 To count + labels.count + nodeMap.count + 100)
    parts(0) = "<svg xmlns=""http://www.w3.org/2000/svg"" width=""" & num(canvasW) & """ height=""" & num(canvasH) & """ viewBox=""0 0 " & num(canvasW) & " " & num(canvasH) & """><g font-family=""Meiryo, sans-serif""><rect width=""100%"" height=""100%"" fill=""white""/>" & SvgText(30, 28, title & "  " & CStr(count) & "要素", 18) & SvgText(30, 51, subtitle, 11)
    partCount = 1
    If mode <> "RESULT" Then
        parts(partCount) = SvgText(30, 72, "X/Yは拘束方向（赤）。灰：通常 / 青：BIRTH予定 / 赤：DEATH予定", 10): partCount = partCount + 1
        If Len(FEMUiMeshQualityNote) > 0 Then parts(partCount) = SvgText(30, 87, FEMUiMeshQualityNote, 10): partCount = partCount + 1
    End If
    If mode = "RESULT" And Len(incrementCaption) > 0 Then
        parts(partCount) = SvgText(30, 72, incrementCaption, 10): partCount = partCount + 1
        If Len(overlayLegend) > 0 Then parts(partCount) = SvgText(30, 87, overlayLegend, 9): partCount = partCount + 1
    End If
    order = Array(0, 4, 1, 5, 2, 6, 3, 7)
    If mode <> "RESULT" And mixedT6Count > 0 Then
        parts(partCount) = SvgText(30, 101, "緑の輪郭：T6 " & CStr(mixedT6Count) & "要素 / Q8 " & CStr(continuumCount - mixedT6Count) & "要素", 10): partCount = partCount + 1
    End If
    For r = 2 To UBound(elements, 1)
        If selected(r) Then
            e = id(elements(r, 1)): points = ""
            If jointRow(r) Then
                points = num(originX + drawScale * px(r, 0)) & "," & num(originY - drawScale * py(r, 0)) & " " & num(originX + drawScale * px(r, 2)) & "," & num(originY - drawScale * py(r, 2)) & " " & num(originX + drawScale * px(r, 1)) & "," & num(originY - drawScale * py(r, 1))
                current = num(originX + drawScale * px(r, 3)) & "," & num(originY - drawScale * py(r, 3)) & " " & num(originX + drawScale * px(r, 5)) & "," & num(originY - drawScale * py(r, 5)) & " " & num(originX + drawScale * px(r, 4)) & "," & num(originY - drawScale * py(r, 4))
                parts(partCount) = "<g stroke=""#7b1fa2"" fill=""none"" stroke-width=""1.4""><title>接合要素 " & CStr(e) & "：接合結果参照</title><polyline points=""" & points & """/><polyline points=""" & CStr(current) & """ stroke-dasharray=""3,2""/></g>"
                partCount = partCount + 1
                GoTo NextDrawElement
            End If
            For j = 0 To 7
                points = points & num(originX + drawScale * px(r, order(j))) & "," & num(originY - drawScale * py(r, order(j))) & " "
            Next j
            fill = "#e2e8ef": stroke = "#718096"
            If mode = "RESULT" Then
                fill = Color(CDbl(values(e)), low, high): stroke = "#4b5563"
                If incrementPaint Or cumPaint Then fill = FEMIncrementFill(CDbl(values(e)), high)
                If statePaint Then fill = FEMPlasticStateFill(CLng(values(e)))
            Else
                flag = 0
                If IsNumeric(elements(r, 10)) Then
                    mat = CLng(elements(r, 10))
                    If plans.Exists(mat) Then flag = CLng(plans(mat))
                End If
                Select Case flag
                    Case 1: fill = "#bbdefb": stroke = "#1976d2"
                    Case 2: fill = "#ffcdd2": stroke = "#c62828"
                    Case 3: fill = "#e1bee7": stroke = "#7b1fa2"
                End Select
            End If
            If mode <> "RESULT" And ViewerT6Row(elements, r) Then stroke = "#008b6d"
            parts(partCount) = "<polygon points=""" & points & """ fill=""" & fill & """ stroke=""" & stroke & """ stroke-width=""0.45""><title>要素 " & CStr(e) & IIf(ViewerT6Row(elements, r), " T6", " Q8") & "</title></polygon>"
            partCount = partCount + 1
        End If
NextDrawElement:
    Next r
    If Len(overlaySvg) > 0 Then parts(partCount) = overlaySvg: partCount = partCount + 1
    If mode <> "RESULT" Then
        For r = 2 To UBound(nodes, 1)
            If Not IsEmpty(nodes(r, 1)) Then
                n = id(nodes(r, 1))
                If visibleNodes.Exists(n) Then
                    supportX = False: supportY = False
                    If Not IsEmpty(nodes(r, 4)) Then supportX = (number(nodes(r, 4), "X拘束") <> 0#)
                    If Not IsEmpty(nodes(r, 5)) Then supportY = (number(nodes(r, 5), "Y拘束") <> 0#)
                    If supportX Or supportY Then
                        node = nodeMap(n): x = originX + drawScale * node(0): y = originY - drawScale * node(1)
                        supportLabel = "": If supportX Then supportLabel = "X"
                        If supportY Then supportLabel = supportLabel & "Y"
                        parts(partCount) = "<g data-node=""" & CStr(n) & """ data-support=""" & supportLabel & """><rect x=""" & num(x - 2#) & """ y=""" & num(y - 2#) & """ width=""4"" height=""4"" fill=""#be3333""/>" & SvgText(x + 4#, y - 3#, supportLabel, 8) & "</g>"
                        partCount = partCount + 1
                    End If
                End If
            End If
        Next r
    End If
    For Each key In labels.keys
        node = labels(key)
        parts(partCount) = SvgText(originX + drawScale * node(0) + 2#, originY - drawScale * node(1) - 2#, CStr(key), 9)
        partCount = partCount + 1
    Next key
    If count = 0 Then parts(partCount) = SvgText(40, 120, "表示対象の要素はありません。", 15): partCount = partCount + 1
    If mode = "RESULT" And statePaint Then
        stateNames = Array("不明（累積なし）", "弾性（塑性ひずみなし）", "過去に降伏（現在は弾性）", "現在降伏")
        For j = 0 To 3
            parts(partCount) = "<rect x=""" & num(40# + j * 205#) & """ y=""" & num(canvasH - 50#) & """ width=""16"" height=""13"" fill=""" & FEMPlasticStateFill(j - 1) & """ stroke=""#4b5563"" stroke-width=""0.5""/>" & SvgText(40# + j * 205# + 20#, canvasH - 39#, stateNames(j) & " " & CStr(stateCount(j)) & "要素", 10)
            partCount = partCount + 1
        Next j
        parts(partCount) = SvgText(40, canvasH - 17#, "4積分点のうち1点でも現在降伏なら赤。黄は塑性ひずみが残るが今は降伏面の内側。累積はBIRTH・RESETで0から。", 10)
    ElseIf mode = "RESULT" Then
        For j = 0 To 39
            parts(partCount) = "<rect x=""" & num(40# + j * 10#) & """ y=""" & num(canvasH - 48#) & """ width=""10.1"" height=""13"" fill=""" & FEMViewerKeyColor(low + (high - low) * j / 39#, low, high, incrementPaint Or cumPaint) & """/>"
            partCount = partCount + 1
        Next j
        parts(partCount) = SvgText(40, canvasH - 17#, Format$(low, "0.000E+00")) & SvgText(355, canvasH - 17#, Format$(high, "0.000E+00"))
    Else
        parts(partCount) = SvgText(40, canvasH - 40#, "灰：通常   青：BIRTH予定   赤：DEATH予定   紫：BIRTH・DEATH両方") & SvgText(40, canvasH - 18#, "予定色は材料番号に対応。無効なステージは色分けに含めません。", 11)
    End If
    parts(partCount + 1) = "</g></svg>"
    ReDim Preserve parts(0 To partCount + 1)
    svg = Join(parts, vbNullString)
    Set ws = ThisWorkbook.Worksheets("図")
    viewPhase = "Excel標準図形への描画"
    Set picture = FEMNativeFigureCreate(ws, svg, canvasW, canvasH)
    ' 新しい図が作成できた後だけ旧図を消す。表示失敗時には旧図を保持する。
    P1ClearViewerLayers
    picture.name = "P1Agg_ViewSVG"
    picture.AlternativeText = title & " / " & CStr(count) & "要素"
    picture.Placement = xlFreeFloating
    If mode = "RESULT" Then
        On Error Resume Next
        ws.Shapes("UI_MeshStart").Visible = msoFalse
        ws.Shapes("UI_MeshRepair").Visible = msoFalse
        On Error GoTo Failed
    End If
    FEMViewerLastSvg = svg: FEMViewerCanvasW = canvasW: FEMViewerCanvasH = canvasH
    FEMViewerElementCount = count
    viewPhase = "図シートの表示"
    ws.Visible = xlSheetVisible
    ws.Activate
    ThisWorkbook.Windows(1).Zoom = 100
    ThisWorkbook.Windows(1).ScrollRow = 1
    ThisWorkbook.Windows(1).ScrollColumn = 1
    GoTo Finish
Failed:
    errorNumber = Err.number
    errorText = CStr(errorNumber) & " [" & viewPhase & "] " & Err.description
    FEMViewerLastError = errorText
Finish:
    On Error Resume Next
    If stateSaved Then
        Application.StatusBar = oldStatus
        Application.DisplayStatusBar = oldDisplayStatus
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    viewerBusy = False
    On Error GoTo 0
    If Len(errorText) > 0 And Not SuppressUserMessages Then MsgBox "表示できませんでした。" & vbCrLf & errorText, vbExclamation
End Sub

Private Function FEMViewerKeyColor(ByVal value As Double, ByVal low As Double, ByVal high As Double, ByVal incrementPaint As Boolean) As String
    If incrementPaint Then FEMViewerKeyColor = FEMIncrementFill(value, high) Else FEMViewerKeyColor = Color(value, low, high)
End Function

Private Function FEMPlasticStateFill(ByVal state As Long) As String
    Select Case state
        Case 2: FEMPlasticStateFill = "#d62728"
        Case 1: FEMPlasticStateFill = "#f6c343"
        Case 0: FEMPlasticStateFill = "#dbe8f5"
        Case Else: FEMPlasticStateFill = "#9e9e9e"
    End Select
End Function

Private Function FEMIncrementFill(ByVal value As Double, ByVal high As Double) As String
    Dim t As Double, g As Long, b As Long
    If high > 0# Then t = value / high
    If t < 0# Then t = 0#
    If t > 1# Then t = 1#
    g = CLng(255# - 220# * t): b = CLng(229# - 209# * t)
    FEMIncrementFill = "#FF" & right$("0" & Hex$(g), 2) & right$("0" & Hex$(b), 2)
End Function

Private Function ViewerT6Row(ByRef elements As Variant, ByVal r As Long) As Boolean
    ViewerT6Row = (elements(r, 4) = elements(r, 5) And elements(r, 4) = elements(r, 8))
End Function
