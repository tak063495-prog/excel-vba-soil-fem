Option Explicit
' Translate the editable shape page into the unchanged native geometry schema.
Public Sub FEMUiBuildGeometry()
  Dim sourcePage As Worksheet, legacyPage As Worksheet
  Dim coords As Object, allSeen As Object, outerSeen As Object, innerSeen As Object, nativeIds As Object
  Dim outerIds As Collection, innerIds As Collection, boundaries As Collection
  Dim r As Long, lastRow As Long, id As Long, key As Variant, point As Variant
  Dim kind As String, sequence As String, tokens As Variant, record As Variant, token As Variant
  Dim spacing As Double, outerCount As Long, innerCount As Long, count As Long
  Dim i As Long, j As Long, legacyRow As Long, nativeIndex As Long, rowLimit As Long
  Dim pointData() As Variant, memberData() As Variant
  Dim eventsSaved As Boolean, errorNo As Long, errorText As String, errorSource As String
  Set sourcePage = ThisWorkbook.Worksheets("形状入力")
  Set legacyPage = ThisWorkbook.Worksheets("図形定義")
  Set coords = CreateObject("Scripting.Dictionary")
  Set allSeen = CreateObject("Scripting.Dictionary")
  Set outerSeen = CreateObject("Scripting.Dictionary"): Set innerSeen = CreateObject("Scripting.Dictionary")
  Set nativeIds = CreateObject("Scripting.Dictionary")
  Set outerIds = New Collection: Set innerIds = New Collection: Set boundaries = New Collection
  lastRow = Application.Max(sourcePage.Cells(sourcePage.rows.count, 2).End(xlUp).row, sourcePage.Cells(sourcePage.rows.count, 3).End(xlUp).row, sourcePage.Cells(sourcePage.rows.count, 4).End(xlUp).row)
  For r = 8 To lastRow
    If Len(Trim$(CStr(sourcePage.Cells(r, 2).value2))) + Len(Trim$(CStr(sourcePage.Cells(r, 3).value2))) + Len(Trim$(CStr(sourcePage.Cells(r, 4).value2))) > 0 Then
      id = UIGeometryInteger(sourcePage.Cells(r, 2).value2, "座標表の節点番号（行" & CStr(r) & "）")
      If coords.Exists(CStr(id)) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "座標表の節点番号が重複しています: " & CStr(id)
      If Not IsNumeric(sourcePage.Cells(r, 3).value2) Or Not IsNumeric(sourcePage.Cells(r, 4).value2) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "X/Y座標は数値で入力してください（行" & CStr(r) & "）。"
      If Len(CStr(sourcePage.Cells(r, 3).value2)) = 0 Or Len(CStr(sourcePage.Cells(r, 4).value2)) = 0 Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "X/Y座標が未入力です（行" & CStr(r) & "）。"
      coords.Add CStr(id), Array(CDbl(sourcePage.Cells(r, 3).value2), CDbl(sourcePage.Cells(r, 4).value2))
    End If
  Next r
  If coords.count < 3 Or coords.count > 60000 Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "形状の節点は3～60000点で入力してください。"
  lastRow = Application.Max(sourcePage.Cells(sourcePage.rows.count, 8).End(xlUp).row, sourcePage.Cells(sourcePage.rows.count, 9).End(xlUp).row, sourcePage.Cells(sourcePage.rows.count, 10).End(xlUp).row)
  For r = 8 To lastRow
    kind = Trim$(CStr(sourcePage.Cells(r, 8).value2))
    sequence = Replace(Replace(Trim$(CStr(sourcePage.Cells(r, 9).value2)), " ", ""), "、", ",")
    If Len(kind) + Len(sequence) + Len(CStr(sourcePage.Cells(r, 10).value2)) > 0 Then
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
      If Not IsNumeric(sourcePage.Cells(r, 10).value2) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "格子間隔は正の数値で入力してください。"
      spacing = CDbl(sourcePage.Cells(r, 10).value2)
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
  For Each key In coords.Keys
    If Not allSeen.Exists(CStr(key)) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "境界の節点順に含まれていない座標があります: " & CStr(key)
    If outerSeen.Exists(CStr(key)) And innerSeen.Exists(CStr(key)) Then Err.Raise vbObjectError + 3890, "FEMUiBuildGeometry", "外側境界と内側境界で同じ節点を共有できません: " & CStr(key)
  Next key
  ReDim pointData(1 To coords.count, 1 To 3)
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
  ' All input is validated before replacing any legacy data.
  eventsSaved = Application.EnableEvents
  On Error GoTo Failed
  Application.EnableEvents = False
  rowLimit = Application.Max(24 + CLng(Val(legacyPage.Cells(24, 4).value2)), 24 + innerCount)
  lastRow = legacyPage.Cells(legacyPage.rows.count, 1).End(xlUp).row
  If lastRow >= 2 Then legacyPage.range(legacyPage.Cells(2, 1), legacyPage.Cells(lastRow, 3)).ClearContents
  legacyPage.range(legacyPage.Cells(3, 4), legacyPage.Cells(rowLimit, 2005)).ClearContents
  legacyPage.range("A1:C1").value2 = Array("No", "X", "Y")
  legacyPage.range(legacyPage.Cells(2, 1), legacyPage.Cells(coords.count + 1, 3)).value2 = pointData
  legacyPage.Cells(1, 7).value2 = "外側境界の節点数": legacyPage.Cells(1, 8).value2 = outerIds.count
  legacyPage.Cells(1, 9).value2 = "内部境界の節点数": legacyPage.Cells(1, 10).value2 = innerIds.count
  legacyPage.Cells(1, 11).value2 = "座標総数": legacyPage.Cells(1, 12).value2 = coords.count
  legacyPage.Cells(1, 13).value2 = "境界平均間隔"
  legacyPage.Cells(2, 5).value2 = "境界節点数": legacyPage.Cells(2, 6).value2 = "格子間隔"
  legacyPage.Cells(3, 4).value2 = outerCount: legacyPage.Cells(24, 4).value2 = innerCount
  i = 0: j = 0: spacing = 0#
  For Each record In boundaries
    tokens = record(1): count = UBound(tokens) + 1
    If record(0) = "OUTER" Then i = i + 1: legacyRow = 2 + i Else j = j + 1: legacyRow = 23 + j
    legacyPage.Cells(legacyRow, 5).value2 = count: legacyPage.Cells(legacyRow, 6).value2 = record(2)
    spacing = spacing + CDbl(record(2))
    ReDim memberData(1 To 1, 1 To count)
    For nativeIndex = 1 To count: memberData(1, nativeIndex) = nativeIds(CStr(tokens(nativeIndex - 1))): Next nativeIndex
    legacyPage.range(legacyPage.Cells(legacyRow, 7), legacyPage.Cells(legacyRow, 6 + count)).value2 = memberData
  Next record
  legacyPage.Cells(1, 14).value2 = spacing / boundaries.count
FEMInvalidateInputCache:   FEMViewerMarkDirty
  AnalysisOK = False: ResultRevision = 0: P1ResultReady = False
  Application.EnableEvents = eventsSaved
  Exit Sub
Failed:
  errorNo = Err.Number: errorSource = Err.source: errorText = Err.Description
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
  FEMUiBuildGeometry
  Call ボタン19_Click
  Exit Sub
Failed:
  MsgBox "形状入力を確認してください。" & vbCrLf & Err.Description, vbExclamation
End Sub
