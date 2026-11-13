Option Explicit
' Translate the editable shape page into the unchanged native geometry schema.
' The coordinate table is a ledger. Only coordinates that a boundary references
' are used, so unused points can stay registered without entering the mesh.
Public FEMUiGeometryNotice As String
Public FEMUiGeometryUnused As String
Public Sub FEMUiBuildGeometry()
  Dim sourcePage As Worksheet, legacyPage As Worksheet
  Dim coords As Object, allSeen As Object, outerSeen As Object, innerSeen As Object, nativeIds As Object
  Dim coordRows As Object, unusedRows As Object
  Dim outerIds As Collection, innerIds As Collection, boundaries As Collection
  Dim r As Long, lastRow As Long, id As Long, key As Variant, point As Variant
  Dim kind As String, sequence As String, tokens As Variant, record As Variant, token As Variant
  Dim spacing As Double, outerCount As Long, innerCount As Long, count As Long
  Dim i As Long, j As Long, legacyRow As Long, nativeIndex As Long, rowLimit As Long
  Dim registeredCount As Long, usedCount As Long, diagRow As Long, listRow As Long
  Dim pointData() As Variant, memberData() As Variant
  Dim diagnostic() As Variant, noticeText As String, unusedText As String
  Dim eventsSaved As Boolean, errorNo As Long, errorText As String, errorSource As String
  FEMUiGeometryNotice = "": FEMUiGeometryUnused = ""
  Set sourcePage = ThisWorkbook.Worksheets("形状入力")
  Set legacyPage = ThisWorkbook.Worksheets("図形定義")
  Set coords = CreateObject("Scripting.Dictionary")
  Set allSeen = CreateObject("Scripting.Dictionary")
  Set outerSeen = CreateObject("Scripting.Dictionary"): Set innerSeen = CreateObject("Scripting.Dictionary")
  Set nativeIds = CreateObject("Scripting.Dictionary")
  Set coordRows = CreateObject("Scripting.Dictionary")
  Set unusedRows = CreateObject("Scripting.Dictionary")
  Set outerIds = New Collection: Set innerIds = New Collection: Set boundaries = New Collection
  lastRow = Application.Max(sourcePage.cells(sourcePage.rows.count, 2).End(xlUp).row, sourcePage.cells(sourcePage.rows.count, 3).End(xlUp).row, sourcePage.cells(sourcePage.rows.count, 4).End(xlUp).row)
  For r = 8 To lastRow
    If Len(Trim$(CStr(sourcePage.cells(r, 2).value2))) + Len(Trim$(CStr(sourcePage.cells(r, 3).value2))) + Len(Trim$(CStr(sourcePage.cells(r, 4).value2))) > 0 Then
      id = UIGeometryInteger(sourcePage.cells(r, 2).value2, "座標表の節点番号（行" & CStr(r) & "）")
      If coords.Exists(CStr(id)) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "座標表の節点番号が重複しています: " & CStr(id)
      If Not IsNumeric(sourcePage.cells(r, 3).value2) Or Not IsNumeric(sourcePage.cells(r, 4).value2) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "X/Y座標は数値で入力してください（行" & CStr(r) & "）。"
      If Len(CStr(sourcePage.cells(r, 3).value2)) = 0 Or Len(CStr(sourcePage.cells(r, 4).value2)) = 0 Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "X/Y座標が未入力です（行" & CStr(r) & "）。"
      coords.Add CStr(id), Array(CDbl(sourcePage.cells(r, 3).value2), CDbl(sourcePage.cells(r, 4).value2))
      coordRows.Add CStr(id), r
    End If
  Next r
  If coords.count < 3 Or coords.count > 60000 Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "形状の節点は3～60000点で入力してください。"
  registeredCount = coords.count
  lastRow = Application.Max(sourcePage.cells(sourcePage.rows.count, 8).End(xlUp).row, sourcePage.cells(sourcePage.rows.count, 9).End(xlUp).row, sourcePage.cells(sourcePage.rows.count, 10).End(xlUp).row)
  For r = 8 To lastRow
    kind = Trim$(CStr(sourcePage.cells(r, 8).value2))
    sequence = Replace(Replace(Trim$(CStr(sourcePage.cells(r, 9).value2)), " ", ""), "、", ",")
    If Len(kind) + Len(sequence) + Len(CStr(sourcePage.cells(r, 10).value2)) > 0 Then
      If kind = "外側" Or UCase$(kind) = "OUTER" Then
        kind = "OUTER": outerCount = outerCount + 1
      ElseIf kind = "内側" Or UCase$(kind) = "INNER" Then
        kind = "INNER": innerCount = innerCount + 1
      Else
        Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "境界の種類は外側または内側を選択してください（行" & CStr(r) & "）。"
      End If
      tokens = Split(sequence, ",")
      count = UBound(tokens) - LBound(tokens) + 1
      If count < 3 Or count > 1999 Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "境界の節点順は3～1999点をカンマで区切って入力してください。"
      If Not IsNumeric(sourcePage.cells(r, 10).value2) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "格子間隔は正の数値で入力してください。"
      spacing = CDbl(sourcePage.cells(r, 10).value2)
      If spacing <= 0# Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "格子間隔は0より大きい値にしてください。"
      Dim boundarySeen As Object
      Set boundarySeen = CreateObject("Scripting.Dictionary")
      For i = LBound(tokens) To UBound(tokens)
        id = UIGeometryInteger(tokens(i), "境界の節点番号")
        If Not coords.Exists(CStr(id)) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "境界の節点番号が座標表にありません: " & CStr(id)
        If boundarySeen.Exists(CStr(id)) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "同じ境界で節点番号が重複しています。始点を末尾に重ねず一周分を入力してください。"
        boundarySeen.Add CStr(id), True: tokens(i) = CStr(id)
        If Not allSeen.Exists(CStr(id)) Then allSeen.Add CStr(id), True
        If kind = "OUTER" Then
          If Not outerSeen.Exists(CStr(id)) Then outerSeen.Add CStr(id), True: outerIds.Add CStr(id)
        Else
          If Not innerSeen.Exists(CStr(id)) Then innerSeen.Add CStr(id), True: innerIds.Add CStr(id)
        End If
      Next i
      boundaries.Add Array(kind, tokens, spacing)
    End If
  Next r
  If outerCount < 1 Or outerCount > 21 Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "外側境界は1～21行で入力してください。"
  ' Unused ledger points are reported with their original rows, never rejected.
  For Each key In coords.keys
    If Not allSeen.Exists(CStr(key)) Then unusedRows.Add CStr(key), coordRows(CStr(key))
    If outerSeen.Exists(CStr(key)) And innerSeen.Exists(CStr(key)) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "外側境界と内側境界で同じ節点を共有できません: " & CStr(key)
  Next key
  usedCount = allSeen.count
  If usedCount < 3 Or usedCount > 60000 Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "境界で使用する座標は3～60000点で入力してください。未使用点は含まれません。"
  If outerIds.count + innerIds.count <> usedCount Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "使用点数と内外境界の節点数が一致しません。"
  ReDim pointData(1 To usedCount, 1 To 3)
  nativeIndex = 0
  For Each key In outerIds
    nativeIndex = nativeIndex + 1: nativeIds.Add CStr(key), nativeIndex
    point = coords(CStr(key))
    pointData(nativeIndex, 1) = nativeIndex: pointData(nativeIndex, 2) = point(0): pointData(nativeIndex, 3) = point(1)
  Next key
  For Each key In innerIds
    nativeIndex = nativeIndex + 1: nativeIds.Add CStr(key), nativeIndex
    point = coords(CStr(key))
    pointData(nativeIndex, 1) = nativeIndex: pointData(nativeIndex, 2) = point(0): pointData(nativeIndex, 3) = point(1)
  Next key
  If nativeIds.count <> usedCount Or nativeIndex <> usedCount Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "外部IDから内部番号への対応数が一致しません。"
  For Each record In boundaries
    tokens = record(1)
    For i = LBound(tokens) To UBound(tokens)
      If Not nativeIds.Exists(CStr(tokens(i))) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "境界の節点番号が内部番号に対応付けられません: " & CStr(tokens(i))
    Next i
  Next record
  noticeText = "登録" & CStr(registeredCount) & "点、使用" & CStr(usedCount) & "点、未使用" & CStr(unusedRows.count) & "点。"
  If unusedRows.count > 0 Then
    noticeText = noticeText & "未使用点は今回の形状生成・解析に使用しません。"
    listRow = 0: unusedText = ""
    For Each key In unusedRows.keys
      listRow = listRow + 1
      If listRow <= 12 Then
        If Len(unusedText) > 0 Then unusedText = unusedText & "、"
        unusedText = unusedText & CStr(key) & "（" & CStr(unusedRows(CStr(key))) & "行目）"
      End If
    Next key
    If unusedRows.count > 12 Then unusedText = unusedText & " ほか" & CStr(unusedRows.count - 12) & "点"
  End If
  ' All input is validated before replacing any legacy data.
  eventsSaved = Application.EnableEvents
  On Error GoTo Failed
  Application.EnableEvents = False
  rowLimit = Application.Max(24 + CLng(val(legacyPage.cells(24, 4).value2)), 24 + innerCount)
  lastRow = legacyPage.cells(legacyPage.rows.count, 1).End(xlUp).row
  If lastRow >= 2 Then legacyPage.range(legacyPage.cells(2, 1), legacyPage.cells(lastRow, 3)).ClearContents
  legacyPage.range(legacyPage.cells(3, 4), legacyPage.cells(rowLimit, 2005)).ClearContents
  legacyPage.range("A1:C1").value2 = Array("No", "X", "Y")
  legacyPage.range(legacyPage.cells(2, 1), legacyPage.cells(usedCount + 1, 3)).value2 = pointData
  legacyPage.cells(1, 7).value2 = "外側境界の節点数": legacyPage.cells(1, 8).value2 = outerIds.count
  legacyPage.cells(1, 9).value2 = "内部境界の節点数": legacyPage.cells(1, 10).value2 = innerIds.count
  legacyPage.cells(1, 11).value2 = "座標総数": legacyPage.cells(1, 12).value2 = usedCount
  legacyPage.cells(1, 13).value2 = "境界平均間隔"
  legacyPage.cells(2, 5).value2 = "境界節点数": legacyPage.cells(2, 6).value2 = "格子間隔"
  legacyPage.cells(3, 4).value2 = outerCount: legacyPage.cells(24, 4).value2 = innerCount
  i = 0: j = 0: spacing = 0#
  For Each record In boundaries
    tokens = record(1): count = UBound(tokens) + 1
    If record(0) = "OUTER" Then i = i + 1: legacyRow = 2 + i Else j = j + 1: legacyRow = 23 + j
    legacyPage.cells(legacyRow, 5).value2 = count: legacyPage.cells(legacyRow, 6).value2 = record(2)
    spacing = spacing + CDbl(record(2))
    ReDim memberData(1 To 1, 1 To count)
    For nativeIndex = 1 To count: memberData(1, nativeIndex) = nativeIds(CStr(tokens(nativeIndex - 1))): Next nativeIndex
    legacyPage.range(legacyPage.cells(legacyRow, 7), legacyPage.cells(legacyRow, 6 + count)).value2 = memberData
  Next record
  legacyPage.cells(1, 14).value2 = spacing / boundaries.count
  ' Ledger counts and the unused list go below the native schema block only.
  ReDim diagnostic(1 To 7 + unusedRows.count, 1 To 2)
  diagnostic(1, 1) = "形状入力診断": diagnostic(1, 2) = "登録台帳と今回の使用点"
  diagnostic(2, 1) = "登録点数": diagnostic(2, 2) = registeredCount
  diagnostic(3, 1) = "使用点数": diagnostic(3, 2) = usedCount
  diagnostic(4, 1) = "未使用点数": diagnostic(4, 2) = unusedRows.count
  diagnostic(5, 1) = "内部番号対応表": diagnostic(5, 2) = nativeIds.count
  diagnostic(6, 1) = "警告": diagnostic(6, 2) = noticeText
  diagnostic(7, 1) = "未使用の点番号": diagnostic(7, 2) = "元の行"
  listRow = 0
  For Each key In unusedRows.keys
    listRow = listRow + 1
    diagnostic(7 + listRow, 1) = CStr(key): diagnostic(7 + listRow, 2) = unusedRows(CStr(key))
  Next key
  diagRow = Application.Max(usedCount + 3, 26 + innerCount)
  legacyPage.range(legacyPage.cells(diagRow, 1), legacyPage.cells(diagRow + UBound(diagnostic, 1) - 1, 2)).value2 = diagnostic
  Call FEMInvalidateInputCache
  Call FEMViewerMarkDirty
  AnalysisOK = False: ResultRevision = 0: P1ResultReady = False
  Application.EnableEvents = eventsSaved
  FEMUiGeometryNotice = noticeText: FEMUiGeometryUnused = unusedText
  Exit Sub
Failed:
  errorNo = Err.number: errorSource = Err.source: errorText = Err.description
  Application.EnableEvents = eventsSaved
  Err.Raise errorNo, errorSource, errorText
End Sub

Private Function UIGeometryInteger(ByVal value As Variant, ByVal label As String) As Long
  If Not IsNumeric(value) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", label & "は正の整数で入力してください。"
  If CDbl(value) < 1# Or CDbl(value) > 60000# Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", label & "は1～60000で入力してください。"
  If CDbl(value) <> Fix(CDbl(value)) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", label & "は整数で入力してください。"
  UIGeometryInteger = CLng(value)
End Function

Public Sub FEMUiGenerateShape()
  On Error GoTo Failed
  If AnalysisRunning Then Exit Sub
  FEMShapeGenerationOK = False
  FEMUiMeshPreviewError = ""
  FEMUiGeometryNotice = "": FEMUiGeometryUnused = ""
  FEMUiBuildGeometry
  Call ボタン19_Click
  If FEMShapeGenerationOK Then FEMUiPreviewMesh
  If Not SuppressUserMessages And Len(FEMUiGeometryNotice) > 0 Then _
    MsgBox FEMUiGeometryNotice & vbCrLf & "未使用点: " & FEMUiGeometryUnused & vbCrLf & "未使用点の一覧は「図形定義」の形状入力診断に出力しました。", vbInformation
  Exit Sub
Failed:
  FEMUiMeshPreviewError = CStr(Err.number) & " " & Err.description
  If Not SuppressUserMessages Then MsgBox "形状入力を確認してください。" & vbCrLf & Err.description, vbExclamation
End Sub
