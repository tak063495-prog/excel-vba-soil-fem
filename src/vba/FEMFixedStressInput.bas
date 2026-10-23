Option Explicit

Public Sub FSCopyTimeOptions(ByVal sourceId As Long, ByVal targetId As Long)
  Dim source As Long, target As Long, r As Long, ws As Worksheet
  source = FSDetailRow(sourceId): If source = 0 Then Exit Sub
  Set ws = ThisWorkbook.Worksheets("水条件")
  For r = 12 To 211
    If Application.CountA(ws.cells(r, 29).Resize(1, 10)) = 0 Then target = r: Exit For
  Next r
  If target = 0 Then FSCoupledInputError "時間詳細の登録は200工程までです。"
  Call FEM24RestoreCells(ws.cells(target, 29).Resize(1, 10), FEM24CaptureCells(ws.cells(source, 29).Resize(1, 10)))
  ws.cells(target, 29).value2 = targetId
End Sub

Public Sub FSDeleteTimeOptions(ByVal id As Long)
  Dim r As Long
  r = FSDetailRow(id)
  If r > 0 Then ThisWorkbook.Worksheets("水条件").cells(r, 29).Resize(1, 10).ClearContents
End Sub

Public Function FSDetailRow(ByVal stageId As Long) As Long
  Dim r As Long, ws As Worksheet
  Set ws = ThisWorkbook.Worksheets("水条件")
  For r = 12 To 211
    If Len(Trim$(CStr(ws.cells(r, 29).value2))) > 0 Then
      If FSNumber(ws.cells(r, 29).value2, "時間詳細の工程ID") = stageId Then
        If FSDetailRow <> 0 Then FSCoupledInputError "時間詳細の工程IDが重複しています。"
        FSDetailRow = r
      End If
    End If
  Next r
End Function

Public Function FSUnitFactor(ByVal unit As String) As Double
  Select Case UCase$(Trim$(unit))
    Case "", "S", "秒": FSUnitFactor = 1#
    Case "MIN", "分": FSUnitFactor = 60#
    Case "H", "時間", "時": FSUnitFactor = 3600#
    Case "D", "DAY", "日": FSUnitFactor = 86400#
    Case Else: FSCoupledInputError "時間単位は秒/分/時間/日です。"
  End Select
End Function

Public Sub FSGetOptions(ByVal stageId As Long, ByRef unit As Double, ByRef waterMode As String, ByRef loadMode As String, ByRef minimum As Double, ByRef maximum As Double, ByRef saveEvery As Double, ByRef series As String)
  Dim r As Long, ws As Worksheet
  Set ws = ThisWorkbook.Worksheets("水条件"): r = FSDetailRow(stageId)
  unit = 1#: waterMode = "HOLD": loadMode = "INSTANT": minimum = 0#: maximum = 0#: saveEvery = 0#: series = ""
  If r = 0 Then Exit Sub
  unit = FSUnitFactor(CStr(ws.cells(r, 30).value2))
  waterMode = UCase$(Trim$(CStr(ws.cells(r, 31).value2)))
  loadMode = UCase$(Trim$(CStr(ws.cells(r, 32).value2)))
  Select Case waterMode
    Case "", "保持": waterMode = "HOLD"
    Case "開始時変更": waterMode = "START"
    Case "線形変化": waterMode = "LINEAR"
  End Select
  Select Case loadMode
    Case "", "瞬時載荷": loadMode = "INSTANT"
    Case "線形載荷": loadMode = "LINEAR"
  End Select
  If waterMode <> "HOLD" And waterMode <> "START" And waterMode <> "LINEAR" Then FSCoupledInputError "水位変化は保持/開始時変更/線形変化です。"
  If loadMode <> "INSTANT" And loadMode <> "LINEAR" Then FSCoupledInputError "載荷方法は瞬時載荷/線形載荷です。"
  minimum = FSNumber(ws.cells(r, 33).value2, "最小刻み", 0#) * unit
  maximum = FSNumber(ws.cells(r, 34).value2, "最大刻み", 0#) * unit
  saveEvery = FSNumber(ws.cells(r, 35).value2, "保存間隔", 0#) * unit
  series = Trim$(CStr(ws.cells(r, 36).value2))
  If minimum < 0# Or maximum < 0# Or saveEvery < 0# Or (maximum > 0# And maximum < minimum) Then FSCoupledInputError "時間刻みの範囲・保存間隔が不正です。"
End Sub

' Profiles are interpolated in time; each profile is itself a spatial polyline.
Public Sub FSCurveAt(ByVal series As String, ByVal time As Double, ByVal unit As Double, ByVal duration As Double, ByRef first As String, ByRef second As String, ByRef fraction As Double, ByRef nextKnot As Double)
  Dim ws As Worksheet, r As Long, count As Long, previousTime As Double, currentTime As Double, previousProfile As String, profile As String, found As Boolean
  Set ws = ThisWorkbook.Worksheets("水条件"): nextKnot = duration
  For r = 420 To 619
    If Trim$(CStr(ws.cells(r, 22).value2)) = series Then
      currentTime = FSNumber(ws.cells(r, 23).value2, "水位時系列の時刻") * unit
      profile = Trim$(CStr(ws.cells(r, 24).value2))
      If Len(profile) = 0 Then FSCoupledInputError "水位時系列の水位線IDがありません。"
      If count = 0 Then
        If currentTime <> 0# Then FSCoupledInputError "水位時系列は工程内時刻0から登録してください。"
      Else
        If currentTime <= previousTime Then FSCoupledInputError "水位時系列は時刻昇順で重複なく登録してください。"
        If time >= previousTime And time <= currentTime Then
          first = previousProfile: second = profile: fraction = (time - previousTime) / (currentTime - previousTime): found = True
        End If
      End If
      If currentTime > time + 0.0000000001 And currentTime < nextKnot Then nextKnot = currentTime
      previousTime = currentTime: previousProfile = profile: count = count + 1
    End If
  Next r
  If count < 2 Or Abs(previousTime - duration) > 0.000000001 * (1# + duration) Or Not found Then FSCoupledInputError "水位時系列は工程時間全体を覆う2点以上が必要です。"
End Sub

Public Sub FSDrainApply(ByVal setId As String, ByRef locked() As Boolean, ByRef pressure() As Double, ByRef source() As Double)
  Dim ws As Worksheet, r As Long, e As Long, side As Long, a As Long, b As Long, na As Long, nb As Long, mat As Long
  Dim kind As String, profile As String, value As Double, gamma As Double, count As Long, seen As Object, key As String
  Dim gp As Long, xi As Double, weight As Double, dx As Double, dy As Double, jac As Double, n0 As Double, n1 As Double
  ReDim source(NumberOfFreeNode - 1)
  If Len(setId) = 0 Then Exit Sub
  Set ws = ThisWorkbook.Worksheets("水条件"): Set seen = CreateObject("Scripting.Dictionary")
  gamma = FSNumber(ws.range("B4").value2, "水の単位体積重量")
  ReDim locked(NumberOfFreeNode - 1): ReDim pressure(NumberOfFreeNode - 1)
  For r = 216 To 415
    If Trim$(CStr(ws.cells(r, 22).value2)) = setId Then
      count = count + 1
      value = FSNumber(ws.cells(r, 23).value2, "排水境界の要素番号")
      If value <> Fix(value) Then FSCoupledInputError "排水境界の要素番号は整数です。"
      e = CLng(value) - 1
      value = FSNumber(ws.cells(r, 24).value2, "排水境界の辺番号")
      If value <> Fix(value) Then FSCoupledInputError "排水境界の辺番号は整数です。"
      side = CLng(value) - 1
      If e < 0 Or e >= NumberOfElement Or side < 0 Or side > 3 Then FSCoupledInputError "排水境界の要素・辺が範囲外です。"
      If Not FEMElementSideValid(e, side) Or Elem(e).IsJoint Then FSCoupledInputError "排水境界は地盤の有効辺に指定してください。"
      If Material(Elem(e).MatNo).kind <> "SOIL" Then FSCoupledInputError "排水境界はSOIL要素に指定してください。"
      key = CStr(e) & ":" & CStr(side)
      If seen.Exists(key) Then FSCoupledInputError "排水セットの同一辺が重複しています。"
      seen.Add key, True
      kind = UCase$(Trim$(CStr(ws.cells(r, 25).value2)))
      If kind <> "HEAD" And kind <> "PRESSURE" And kind <> "NOFLOW" And kind <> "FLUX" Then FSCoupledInputError "排水境界はHEAD/PRESSURE/NOFLOW/FLUXです。"
      If Not P3IsElementActive(e) Then GoTo NextDrain
      If Not GWMaterialEffective(Elem(e).MatNo) Then FSCoupledInputError "時間排水条件は連成するEFFECTIVE地盤へ指定してください。"
      If kind = "NOFLOW" Then
        If FSInternalPermeableFace(e, side) Then FSCoupledInputError "共有する内部透水面へNOFLOWを指定できません。外周面、または遮水JOINT面を指定してください。要素=" & CStr(e + 1) & "、辺=" & CStr(side + 1)
      End If
      a = side: b = (side + 1) Mod 4
      na = GWFlowNode(e, a): nb = GWFlowNode(e, b)
      profile = Trim$(CStr(ws.cells(r, 26).value2))
      If kind = "HEAD" Or kind = "PRESSURE" Then
        For gp = 0 To 1
          If gp = 0 Then mat = a Else mat = b
          If kind = "HEAD" And Len(profile) > 0 Then
            value = gamma * (GWProfileLevel(profile, Elem(e).x(mat)) - Elem(e).y(mat))
          Else
            value = FSNumber(ws.cells(r, 27).value2, "排水境界の指定値")
            If kind = "HEAD" Then value = gamma * (value - Elem(e).y(mat))
          End If
          If value < -0.0000001 Then FSCoupledInputError "負の指定水圧は飽和領域の適用外です。"
          mat = GWFlowNode(e, mat)
          If locked(mat) And Abs(pressure(mat) - value) > 0.000000001 Then FSCoupledInputError "共有圧力節点の排水境界が矛盾しています。"
          locked(mat) = True: pressure(mat) = value
        Next gp
      ElseIf kind = "FLUX" Then
        value = FSNumber(ws.cells(r, 27).value2, "内向流束[m/s]")
        For gp = 0 To 2
          xi = (gp - 1) * Sqr(0.6): weight = 5# / 9#: If gp = 1 Then weight = 8# / 9#
          dx = (xi - 0.5) * Elem(e).x(a) + (xi + 0.5) * Elem(e).x(b) - 2# * xi * Elem(e).x(side + 4)
          dy = (xi - 0.5) * Elem(e).y(a) + (xi + 0.5) * Elem(e).y(b) - 2# * xi * Elem(e).y(side + 4)
          jac = Sqr(dx * dx + dy * dy) * weight * Material(Elem(e).MatNo).thickness
          source(na) = source(na) + value * jac * (1# - xi) / 2#
          source(nb) = source(nb) + value * jac * (1# + xi) / 2#
        Next gp
      End If
NextDrain:
    End If
  Next r
  If count = 0 Then FSCoupledInputError "排水条件セットIDが登録されていません。"
  GWTransientJointConstraints locked, pressure
End Sub

' A physical neighbour is not necessarily a hydraulic neighbour: PRESCRIBED/SOIL
' JOINT faces keep separate hydraulic aliases.  Conversely, a shared SOIL face
' with the same GWFlowNode aliases is an internal permeable face and cannot be
' made NOFLOW in a stage drain set.  The check is deliberately local and also
' respects the collapsed side convention used by T6.
Private Function FSInternalPermeableFace(ByVal e As Long, ByVal side As Long) As Boolean
  Dim q As Long, qs As Long, key As String, qkey As String
  If Not FEMElementSideValid(e, side) Then Exit Function
  If Material(Elem(e).MatNo).kind <> "SOIL" Then Exit Function
  If Not P3IsElementActive(e) Then Exit Function
  If Not GWMaterialEffective(Elem(e).MatNo) Then Exit Function
  key = FSPhysicalFaceKey(e, side)
  For q = 0 To NumberOfElement - 1
    If q <> e And Not Elem(q).IsJoint Then
      If Material(Elem(q).MatNo).kind = "SOIL" And P3IsElementActive(q) Then
        If GWMaterialEffective(Elem(q).MatNo) Then
          For qs = 0 To 3
            If FEMElementSideValid(q, qs) Then
              If FSPhysicalFaceKey(q, qs) = key Then
                qkey = FSFlowFaceKey(q, qs)
                FSInternalPermeableFace = (FSFlowFaceKey(e, side) = qkey)
                Exit Function
              End If
            End If
          Next qs
        End If
      End If
    End If
  Next q
End Function

Private Function FSPhysicalFaceKey(ByVal e As Long, ByVal side As Long) As String
  Dim a As Long, b As Long, c As Long, t As Long
  a = GJTopoNode(Elem(e).node(side)): b = GJTopoNode(Elem(e).node((side + 1) Mod 4)): c = GJTopoNode(Elem(e).node(side + 4))
  If a > b Then t = a: a = b: b = t
  FSPhysicalFaceKey = CStr(a) & ":" & CStr(b) & ":" & CStr(c)
End Function

Private Function FSFlowFaceKey(ByVal e As Long, ByVal side As Long) As String
  Dim a As Long, b As Long, c As Long, t As Long
  a = GWFlowNode(e, side): b = GWFlowNode(e, (side + 1) Mod 4): c = GWFlowNode(e, side + 4)
  If a > b Then t = a: a = b: b = t
  FSFlowFaceKey = CStr(a) & ":" & CStr(b) & ":" & CStr(c)
End Function

Public Function FSLoadingSet(ByVal stageId As Long) As String
  Dim row As Long
  row = FSDetailRow(stageId)
  If row > 0 Then FSLoadingSet = Trim$(CStr(ThisWorkbook.Worksheets("水条件").cells(row, 37).value2))
End Function

Public Sub FSValidateStageEntry(ByVal stageId As Long, ByVal kind As String)
  Dim r As Long, c As Long, ws As Worksheet, text As String, unit As Double
  If FSRequested(stageId) And kind <> "SRM" Then Exit Sub
  If kind = "SRM" And FSRequested(stageId) And Not FSActive Then FSCoupledInputError "時間連成のSRMには、先に受理した時間連成の親状態が必要です。"
  Set ws = ThisWorkbook.Worksheets("水条件"): r = FSPlanRow(stageId)
  If r > 0 Then
    For c = 15 To 19
      If c <> 17 And Len(Trim$(CStr(ws.cells(r, c).value2))) > 0 Then
        If FSNumber(ws.cells(r, c).value2, "時間・連成設定") <> 0# Then FSCoupledInputError "時間・排水・連成反復の指定はFIXED_STRESS施工工程で使用します。SRMは親水圧・時刻を保持します。工程=" & CStr(stageId)
      End If
    Next c
    If Len(Trim$(CStr(ws.cells(r, 17).value2))) > 0 Then FSCoupledInputError "SRMまたは片方向工程では時間排水セットを変更できません。"
  End If
  r = FSDetailRow(stageId): If r = 0 Then Exit Sub
  unit = FSUnitFactor(CStr(ws.cells(r, 30).value2))
  text = UCase$(Trim$(CStr(ws.cells(r, 31).value2)))
  If text <> "" And text <> "HOLD" And text <> "保持" Then FSCoupledInputError "水位の時間変化はFIXED_STRESSのWATER工程へ指定してください。"
  text = UCase$(Trim$(CStr(ws.cells(r, 32).value2)))
  If text <> "" And text <> "INSTANT" And text <> "瞬時載荷" Then FSCoupledInputError "線形載荷はFIXED_STRESSの施工工程へ指定してください。"
  For c = 33 To 35
    If Len(Trim$(CStr(ws.cells(r, c).value2))) > 0 Then
      If FSNumber(ws.cells(r, c).value2, "時間詳細") <> 0# Then FSCoupledInputError "時間詳細はFIXED_STRESSの施工工程へ指定してください。"
    End If
  Next c
  If Len(Trim$(CStr(ws.cells(r, 36).value2))) > 0 Or Len(Trim$(CStr(ws.cells(r, 37).value2))) > 0 Then FSCoupledInputError "時系列・載荷セットはFIXED_STRESSの施工工程へ指定してください。"
End Sub

Public Sub FSValidateTimeDefinitions()
  Dim ws As Worksheet, seen As Object, r As Long, i As Long, id As Double, matches As Long, active As Boolean
  Dim unit As Double, waterMode As String, loadMode As String, minimum As Double, maximum As Double, saveEvery As Double, series As String
  Set ws = ThisWorkbook.Worksheets("水条件"): Set seen = CreateObject("Scripting.Dictionary")
  For r = 12 To 211
    If Application.CountA(ws.cells(r, 29).Resize(1, 10)) > 0 Then
      id = FSNumber(ws.cells(r, 29).value2, "時間詳細の工程ID")
      If id < 1 Or id > 2000000000# Or id <> Fix(id) Then FSCoupledInputError "時間詳細の工程IDは正の整数です。"
      If seen.Exists(CStr(id)) Then FSCoupledInputError "時間詳細の工程IDが重複しています。"
      seen.Add CStr(id), True: matches = 0: active = False
      For i = 1 To P3StageN
        If P3StageId(i) = CLng(id) Then matches = matches + 1: active = P3StageOn(i)
      Next i
      If matches <> 1 Then FSCoupledInputError "時間詳細の工程IDが工程表に一意に対応していません。ID=" & CStr(id)
      If active Then FSGetOptions CLng(id), unit, waterMode, loadMode, minimum, maximum, saveEvery, series
    End If
  Next r
End Sub
