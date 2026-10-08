Option Explicit

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
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow < 2 Then lastRow = 2
    ReadTable = ws.range(ws.Cells(1, 1), ws.Cells(lastRow, columns)).value2
End Function

Private Function Number(ByVal value As Variant, ByVal field As String) As Double
    If IsError(value) Then Err.Raise vbObjectError + 3700, "FEMViewer", field & "がエラーです。"
    If IsEmpty(value) Then Err.Raise vbObjectError + 3700, "FEMViewer", field & "が空欄です。"
    If Not IsNumeric(value) Then Err.Raise vbObjectError + 3700, "FEMViewer", field & "が数値ではありません。"
    Number = CDbl(value)
    If Abs(Number) > 1E+100 Then Err.Raise vbObjectError + 3700, "FEMViewer", field & "が表示範囲外です。"
End Function

Private Function id(ByVal value As Variant) As Long
    Dim n As Double
    n = Number(value, "番号")
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
                    values.Add e, Number(rows(r, column), "結果要素の選択列")
                End If
            End If
        End If
    Next r
    Set SavedElementValues = values
End Function
Public Sub FEMViewerSetupSelector()
    Dim ws As Worksheet, Table As Worksheet, i As Long, row As Long, last As Long, column As String
    Dim data(1 To 69, 1 To 4) As Variant, oldEvents As Boolean
    oldEvents = Application.EnableEvents: Application.EnableEvents = False
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("設定")
    On Error Resume Next: Set Table = ThisWorkbook.Worksheets("表示量一覧"): On Error GoTo Failed
    If Table Is Nothing Then Set Table = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): Table.name = "表示量一覧"
    data(1, 1) = "選択番号": data(1, 2) = "結果要素の列": data(1, 3) = "表示物理量（積分点別）": data(1, 4) = "単位"
    For i = 1 To 68
        column = Split(ThisWorkbook.Worksheets("結果要素").Cells(1, i + 4).address, "$")(1)
        data(i + 1, 1) = i: data(i + 1, 2) = column
        data(i + 1, 3) = ThisWorkbook.Worksheets("結果要素").Cells(1, i + 4).value2
        data(i + 1, 4) = FEMViewerResultUnit(i)
    Next i
    Table.range("A2:D70").value2 = data
    For i = 1 To 68
        column = CStr(data(i + 1, 2))
        Table.Cells(i + 2, 3).formula = "='結果要素'!" & column & "$1"
    Next i
    Table.range("A1").value2 = "C45で番号を選び、変形・応力表示を押してください。各積分点の値を要素の色として表示します。"
    Table.range("A1:D70").Font.name = "メイリオ": Table.range("A1:D70").Font.size = 10
    Table.columns("A:B").ColumnWidth = 14: Table.columns("C").ColumnWidth = 34: Table.columns("D").ColumnWidth = 18
    Table.range("A2:D2").Interior.Color = RGB(31, 70, 105): Table.range("A2:D2").Font.Color = vbWhite
    Table.range("A2:D2").Font.bOld = True: Table.range("A2:D70").RowHeight = 21
    Table.range("A3:A70").NumberFormat = "0": Table.range("A3:B70").HorizontalAlignment = xlCenter
    On Error Resume Next: ThisWorkbook.names("FEMViewerSelectionIDs").Delete: On Error GoTo Failed
    ThisWorkbook.names.Add name:="FEMViewerSelectionIDs", RefersTo:="='表示量一覧'!$A$3:$A$70"
    last = ws.Cells(ws.rows.count, 5).End(xlUp).row
    For row = 2 To last
        If UCase$(CStr(ws.Cells(row, 5).value2)) = "VIEW_COLOR_RESULT" Then
            ws.Cells(row, 2).value2 = "色付け（結果要素列）"
            ws.Cells(row, 4).formula = "=IF(AND(ISNUMBER(C" & row & "),C" & row & ">=1,C" & row & "<=68,C" & row & "=INT(C" & row & ")),INDEX('結果要素'!$E$1:$BT$1,1,C" & row & "),""1～68を選択"")"
            With ws.Cells(row, 3).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=FEMViewerSelectionIDs"
                .IgnoreBlank = False: .InCellDropdown = True: .ShowError = True: .ShowInput = True
                .InputTitle = "結果要素 E=1 ～ BT=68": .InputMessage = "番号と物理量の対応は「表示量一覧」を参照してください。"
                .ErrorTitle = "表示番号": .ErrorMessage = "1～68の整数を選択してください。"
            End With
        End If
    Next row
    ws.Activate
Finish:
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    Application.EnableEvents = oldEvents
    Err.Raise Err.Number, "FEMViewerSetupSelector", Err.Description
End Sub

Private Sub WriteSvg(ByVal path As String, ByVal svg As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText svg
    stream.SaveToFile path, 2
    stream.Close
End Sub

Private Sub DrawView(ByVal mode As String)
    Dim nodes As Variant, elements As Variant, materials As Variant, stages As Variant, settings As Variant
    Dim nodeResults As Variant, elementResults As Variant, nodeMap As Object, materialMap As Object
    Dim plans As Object, displacements As Object, active As Object, values As Object, config As Object, finalMaterials As Object
    Dim r As Long, j As Long, n As Long, e As Long, mat As Long, flag As Long, Stage As Long, count As Long
    Dim selected() As Boolean, px() As Double, py() As Double, elementIds As Object, key As Variant
    Dim jointMaterials As Object, jointElements As Object, jointRow() As Boolean
    Dim nodeLimit As Long, continuumCount As Long, kind As String
    Dim stageResults As Variant, geometryMoved As Boolean, resultX As Double, resultY As Double
    Dim node As Variant, displacement As Variant, current As Variant, order As Variant
    Dim low As Double, high As Double, value As Double, bound As Double, dscale As Double, drawScale As Double
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double, sx As Double, sy As Double
    Dim originX As Double, originY As Double, x As Double, y As Double, firstPoint As Boolean
    Dim resultId As Long, fill As String, stroke As String, points As String, title As String, subtitle As String
    Dim svg As String, parts() As String, partCount As Long, tempPath As String, labels As Object, showLabels As Boolean
    Dim ws As Worksheet, picture As Object, oldScreen As Boolean, oldEvents As Boolean, oldStatus As Variant
    Dim oldDisplayStatus As Boolean, stateSaved As Boolean, errorText As String, lastRow As Long, canvasW As Double, canvasH As Double
    If AnalysisRunning Or viewerBusy Then Exit Sub
    viewerBusy = True
    FEMViewerLastError = "": FEMViewerElementCount = 0
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    oldStatus = Application.StatusBar: oldDisplayStatus = Application.DisplayStatusBar
    stateSaved = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "メッシュ表示を準備しています"
    decimalMark = Application.International(xlDecimalSeparator)
    nodes = ReadTable("節点データ", 3)
    elements = ReadTable("要素データ", 10)
    materials = ReadTable("材料データ", 13)
    stages = ReadTable("ステージ", 5)
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
    For r = 2 To UBound(settings, 1)
        If Len(CStr(settings(r, 5))) > 0 Then config(UCase$(CStr(settings(r, 5)))) = settings(r, 3)
    Next r
    For r = 2 To UBound(nodes, 1)
        If Not IsEmpty(nodes(r, 1)) Then
            n = id(nodes(r, 1))
            If nodeMap.Exists(n) Then Err.Raise vbObjectError + 3705, "FEMViewer", "節点番号が重複しています。"
            nodeMap.Add n, Array(Number(nodes(r, 2), "節点X"), Number(nodes(r, 3), "節点Y"))
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
    For Each key In materialMap.Keys
        finalMaterials(key) = materialMap(key)
    Next key
    For r = 2 To UBound(stages, 1)
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
    If mode = "RESULT" Then
        If ResultsDirty() Then Err.Raise vbObjectError + 3706, "FEMViewer", "入力変更後のため、再解析してから結果表示してください。"
        If CStr(ThisWorkbook.Worksheets("診断").Cells(2, 2).value2) <> "PASS" Then Err.Raise vbObjectError + 3707, "FEMViewer", "解析成功後に結果表示してください。"
        nodeResults = ReadTable("結果節点", 10)
        elementResults = ReadTable("結果要素", 72)
        Stage = LatestStage(nodeResults)
        If Stage = 0 Then Err.Raise vbObjectError + 3708, "FEMViewer", "保存済みの節点結果がありません。"
        ' RESET_STRESS changes the reference geometry without editing the input mesh.
        ' Use saved stage coordinates only after a recorded successful reset.
        stageResults = ReadTable("ステージ結果", 4)
        For r = 2 To UBound(stageResults, 1)
            If UCase$(Trim$(CStr(stageResults(r, 2)))) = "RESET_STRESS" And CStr(stageResults(r, 4)) = "PASS" Then
                If IsNumeric(stageResults(r, 1)) Then
                    If CLng(stageResults(r, 1)) <= Stage Then geometryMoved = True
                End If
            End If
        Next r
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
                    resultX = Number(nodeResults(r, 3), "結果X"): resultY = Number(nodeResults(r, 4), "結果Y")
                    If geometryMoved Then
                        nodeMap(n) = Array(resultX, resultY)
                    Else
                        If Abs(node(0) - resultX) > 0.00000001 * (1# + Abs(node(0))) Or Abs(node(1) - resultY) > 0.00000001 * (1# + Abs(node(1))) Then Err.Raise vbObjectError + 3710, "FEMViewer", "結果と入力の座標が一致しません。再解析してください。"
                    End If
                    displacements(n) = Array(Number(nodeResults(r, 5), "X変位"), Number(nodeResults(r, 6), "Y変位"))
                End If
            End If
        Next r
        resultId = 18
        If config.Exists("VIEW_COLOR_RESULT") Then
            value = Number(config("VIEW_COLOR_RESULT"), "色付け番号")
            If value < 1# Or value > 68# Or value <> Fix(value) Then Err.Raise vbObjectError + 3703, "FEMViewer", "色付けは1～68の整数を指定してください。"
            resultId = CLng(value)
        End If
        Set values = SavedElementValues(elementResults, resultId, Stage, active, jointElements)
        dscale = 1#
        If config.Exists("VIEW_DISP_SCALE") Then dscale = Number(config("VIEW_DISP_SCALE"), "変位倍率")
        If dscale < 0# Then Err.Raise vbObjectError + 3711, "FEMViewer", "変位倍率は0以上で指定してください。"
    End If
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
        nodeLimit = 7
        If jointRow(r) Then nodeLimit = 5 Else continuumCount = continuumCount + 1
        For j = 0 To nodeLimit
            n = id(elements(r, j + 2))
            If Not nodeMap.Exists(n) Then Err.Raise vbObjectError + 3714, "FEMViewer", "接続節点がありません。要素=" & CStr(e)
            node = nodeMap(n): x = node(0): y = node(1)
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
    End If
    If config.Exists("VIEW_SCALE_MODE") Then
        If UCase$(CStr(config("VIEW_SCALE_MODE"))) = "USER" Then
            If config.Exists("VIEW_SCALE") Then drawScale = Number(config("VIEW_SCALE"), "表示スケール")
            If config.Exists("VIEW_MIN_X") Then originX = Number(config("VIEW_MIN_X"), "作図原点X")
            If config.Exists("VIEW_MAX_Y") Then originY = Number(config("VIEW_MAX_Y"), "作図原点Y")
            If drawScale <= 0# Then Err.Raise vbObjectError + 3718, "FEMViewer", "表示スケールは正数で指定してください。"
            If originX + maxX * drawScale + 40# > canvasW Then canvasW = originX + maxX * drawScale + 40#
            If originY - minY * drawScale + 80# > canvasH Then canvasH = originY - minY * drawScale + 80#
        End If
    End If
    Select Case mode
        Case "ALL": title = "図形表示（全メッシュ）": subtitle = "全ステージの要素を表示。色は有効なステージの予定。"
        Case "INITIAL": title = "初期表示": subtitle = "初期状態ONのみ。BIRTH待ち（OFF）は表示しません。"
        Case "RESULT": title = "変形・応力表示": subtitle = "最終ステージ " & CStr(Stage) & " / 変位倍率 " & Format$(dscale, "0.###") & " / " & CStr(elementResults(1, resultId + 4)) & " [" & FEMViewerResultUnit(resultId) & "] / 結果要素 " & Split(ThisWorkbook.Worksheets("結果要素").Cells(1, resultId + 4).address, "$")(1) & "列（番号 " & resultId & "）"
    End Select
    If jointElements.count > 0 Then subtitle = subtitle & " / 接合は紫線：数値は接合結果"
    ReDim parts(0 To count + labels.count + 90)
    parts(0) = "<svg xmlns=""http://www.w3.org/2000/svg"" width=""" & num(canvasW) & """ height=""" & num(canvasH) & """ viewBox=""0 0 " & num(canvasW) & " " & num(canvasH) & """><g font-family=""Meiryo, sans-serif""><rect width=""100%"" height=""100%"" fill=""white""/>" & SvgText(30, 28, title & "  " & CStr(count) & "要素", 18) & SvgText(30, 51, subtitle, 11)
    partCount = 1: order = Array(0, 4, 1, 5, 2, 6, 3, 7)
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
            parts(partCount) = "<polygon points=""" & points & """ fill=""" & fill & """ stroke=""" & stroke & """ stroke-width=""0.45""><title>要素 " & CStr(e) & "</title></polygon>"
            partCount = partCount + 1
        End If
NextDrawElement:
    Next r
    For Each key In labels.Keys
        node = labels(key)
        parts(partCount) = SvgText(originX + drawScale * node(0) + 2#, originY - drawScale * node(1) - 2#, CStr(key), 9)
        partCount = partCount + 1
    Next key
    If count = 0 Then parts(partCount) = SvgText(40, 120, "表示対象の要素はありません。", 15): partCount = partCount + 1
    If mode = "RESULT" Then
        For j = 0 To 39
            parts(partCount) = "<rect x=""" & num(40# + j * 10#) & """ y=""" & num(canvasH - 48#) & """ width=""10.1"" height=""13"" fill=""" & Color(low + (high - low) * j / 39#, low, high) & """/>"
            partCount = partCount + 1
        Next j
        parts(partCount) = SvgText(40, canvasH - 17#, Format$(low, "0.000E+00")) & SvgText(355, canvasH - 17#, Format$(high, "0.000E+00"))
    Else
        parts(partCount) = SvgText(40, canvasH - 40#, "灰：通常   青：BIRTH予定   赤：DEATH予定   紫：BIRTH・DEATH両方") & SvgText(40, canvasH - 18#, "予定色は材料番号に対応。無効なステージは色分けに含めません。", 11)
    End If
    parts(partCount + 1) = "</g></svg>"
    ReDim Preserve parts(0 To partCount + 1)
    svg = Join(parts, vbNullString)
    tempPath = ThisWorkbook.path & Application.PathSeparator & "~FEMView_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 100#)) & ".svg"
    WriteSvg tempPath, svg
    Set ws = ThisWorkbook.Worksheets("図")
    Set picture = ws.Shapes.AddPicture(tempPath, msoFalse, msoTrue, 10#, 10#, canvasW, canvasH)
    ' 新しい図が作成できた後だけ旧図を消す。表示失敗時には旧図を保持する。
    P1ClearViewerLayers
    picture.name = "P1Agg_ViewSVG"
    picture.AlternativeText = title & " / " & CStr(count) & "要素"
    picture.Placement = xlFreeFloating
    FEMViewerElementCount = count
    ws.Activate
    ThisWorkbook.Windows(1).Zoom = 100
    ThisWorkbook.Windows(1).ScrollRow = 1
    ThisWorkbook.Windows(1).ScrollColumn = 1
    GoTo Finish
Failed:
    errorText = Err.Description
    FEMViewerLastError = errorText
Finish:
    On Error Resume Next
    If Len(tempPath) > 0 Then Kill tempPath
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
