Option Explicit

Public FEM24StageError As String
Private mStageUndo As Variant
Private mWaterUndo As Variant
Private mTimeUndo As Variant
Private mStageUndoValid As Boolean

Private Function FEM24Canon(ByVal s As String) As String
  Select Case UCase$(Trim$(s))
    Case "自重", "GRAVITY": FEM24Canon = "GRAVITY"
    Case "載荷", "LOAD": FEM24Canon = "LOAD"
    Case "除荷", "UNLOAD": FEM24Canon = "UNLOAD"
    Case "盛土", "盛土・追加", "追加", "BIRTH": FEM24Canon = "BIRTH"
    Case "掘削", "掘削・撤去", "撤去", "DEATH": FEM24Canon = "DEATH"
    Case "水位変更", "WATER": FEM24Canon = "WATER"
    Case "時間経過", "WAIT": FEM24Canon = "WAIT"
    Case "SRM照査", "SRM": FEM24Canon = "SRM"
    Case "変位リセット", "RESET_U": FEM24Canon = "RESET_U"
    Case "応力リセット", "RESET_STRESS": FEM24Canon = "RESET_STRESS"
    Case "材料変更", "MATSET": FEM24Canon = "MATSET"
    Case Else: FEM24Canon = ""
  End Select
End Function

Private Function FEM24Label(ByVal s As String) As String
  Select Case s
    Case "GRAVITY": FEM24Label = "自重"
    Case "LOAD": FEM24Label = "載荷"
    Case "UNLOAD": FEM24Label = "除荷"
    Case "BIRTH": FEM24Label = "盛土・追加"
    Case "DEATH": FEM24Label = "掘削・撤去"
    Case "WATER": FEM24Label = "水位変更"
    Case "WAIT": FEM24Label = "時間経過"
    Case "SRM": FEM24Label = "SRM照査"
    Case "RESET_U": FEM24Label = "変位リセット"
    Case "RESET_STRESS": FEM24Label = "応力リセット"
    Case "MATSET": FEM24Label = "材料変更"
    Case Else: FEM24Label = s
  End Select
End Function

Private Function PlanRow(ByVal id As Long, Optional ByVal create As Boolean = False) As Long
  Dim w As Worksheet, r As Long, freeRow As Long
  Set w = ThisWorkbook.Worksheets("水条件")
  For r = 12 To 211
    If Len(CStr(w.cells(r, 12).value2)) > 0 Then
      If CLng(val(CStr(w.cells(r, 12).value2))) = id Then PlanRow = r: Exit Function
    ElseIf Application.CountA(w.range(w.cells(r, 12), w.cells(r, 20))) = 0 And freeRow = 0 Then
      freeRow = r
    End If
  Next r
  If create Then
    If freeRow = 0 Then Err.Raise vbObjectError + 2401, , "工程別の水理設定は200行までです。"
    PlanRow = freeRow: w.cells(freeRow, 12).value2 = id
  End If
End Function

Private Sub Snapshot()
  mStageUndo = FEM24CaptureCells(ThisWorkbook.Worksheets("ステージ").range("A2:F2000"))
  mWaterUndo = FEM24CaptureCells(ThisWorkbook.Worksheets("水条件").range("L12:T211"))
  mTimeUndo = FEM24CaptureCells(ThisWorkbook.Worksheets("水条件").range("AC12:AL211"))
  mStageUndoValid = True
End Sub
Private Sub Restore()
  If Not mStageUndoValid Then Exit Sub
  Call FEM24RestoreCells(ThisWorkbook.Worksheets("ステージ").range("A2:F2000"), mStageUndo)
  Call FEM24RestoreCells(ThisWorkbook.Worksheets("水条件").range("L12:T211"), mWaterUndo)
  Call FEM24RestoreCells(ThisWorkbook.Worksheets("水条件").range("AC12:AL211"), mTimeUndo)
End Sub

Private Function MaxId() As Long
  Dim r As Long, w As Worksheet, x As Double
  Set w = ThisWorkbook.Worksheets("ステージ")
  For r = 2 To 2000
    If Len(Trim$(CStr(w.cells(r, 2).value2))) > 0 Then
      If IsNumeric(w.cells(r, 1).value2) Then
        x = CDbl(w.cells(r, 1).value2)
        If x > 2000000000# Or x < 0 Or x <> Fix(x) Then Err.Raise vbObjectError + 2402, , "工程IDは正の整数です。"
        If x > MaxId Then MaxId = CLng(x)
      End If
    End If
  Next r
End Function

' Legacy blank IDs resolve to the ordinal number of nonempty operations.
' Materialize the same IDs on the first edit, never on a render.
Private Sub NormalizeIds()
  Dim w As Worksheet, r As Long, count As Long, id As Long, seen As Object
  Set w = ThisWorkbook.Worksheets("ステージ"): Set seen = CreateObject("Scripting.Dictionary")
  For r = 2 To 2000
    If Len(Trim$(CStr(w.cells(r, 2).value2))) > 0 Then
      count = count + 1
      If Len(Trim$(CStr(w.cells(r, 1).value2))) = 0 Then
        id = count
      Else
        If Not IsNumeric(w.cells(r, 1).value2) Then Err.Raise vbObjectError + 2403, , "工程IDが数値ではありません。"
        If CDbl(w.cells(r, 1).value2) <> Fix(CDbl(w.cells(r, 1).value2)) Then Err.Raise vbObjectError + 2403, , "工程IDは整数です。"
        id = CLng(w.cells(r, 1).value2)
        If id <= 0 Then id = count
      End If
      If seen.Exists(CStr(id)) Then Err.Raise vbObjectError + 2404, , "工程IDが重複しています。既存のステージ詳細表で番号を整理してください。"
      seen.Add CStr(id), True: w.cells(r, 1).value2 = id
    End If
  Next r
End Sub

Private Sub ValidateStages()
  Dim w As Worksheet, water As Worksheet, r As Long, kind As String, profile As String, k As Long, count As Long
  Set w = ThisWorkbook.Worksheets("ステージ"): Set water = ThisWorkbook.Worksheets("水条件")
  NormalizeIds
  For r = 2 To 2000
    If Len(Trim$(CStr(w.cells(r, 2).value2))) > 0 Then
      kind = FEM24Canon(CStr(w.cells(r, 2).value2))
      If kind = "" Then Err.Raise vbObjectError + 2405, , "操作名が不正です。"
      If Not IsNumeric(w.cells(r, 5).value2) Then Err.Raise vbObjectError + 2406, , "有効欄は0または1です。"
      If CDbl(w.cells(r, 5).value2) <> 0 And CDbl(w.cells(r, 5).value2) <> 1 Then Err.Raise vbObjectError + 2406, , "有効欄は0または1です。"
      If kind = "WATER" And val(CStr(w.cells(r, 5).value2)) <> 0 Then
        profile = GWResolveStageProfile(CLng(w.cells(r, 1).value2), Trim$(CStr(w.cells(r, 3).value2)))
        If profile <> "" Then
        count = 0
        For k = 12 To 211
          If Trim$(CStr(water.cells(k, 1).value2)) = profile Then count = count + 1
        Next k
        If count < 2 Then Err.Raise vbObjectError + 2407, , "水位線は同じIDで2点以上登録してください。"
        End If
      End If
    End If
  Next r
  GWValidateStagePlans
End Sub

Public Sub FEM24StagesRender(ByVal ws As Worksheet)
  Dim src As Worksheet, water As Worksheet, r As Long, outRow As Long, id As Long, kind As String, pr As Long, profile As String, inherited As String, count As Long, lastOut As Long
  Set src = ThisWorkbook.Worksheets("ステージ"): Set water = ThisWorkbook.Worksheets("水条件")
  ws.range("A7:I7").value = Array("工程ID", "操作", "対象材料", "倍率・Fs・変更式", "有効", "工程名", "連携方式", "水位線指定", "採用水位線")
  ws.range("A8:I2006").ClearContents
  inherited = CStr(water.range("B5").value2)
  For r = 2 To 2000
    outRow = r + 6
    If Len(Trim$(CStr(src.cells(r, 2).value2))) > 0 Then
      lastOut = outRow
      count = count + 1: id = CLng(val(CStr(src.cells(r, 1).value2)))
      If id <= 0 Then id = count
      kind = FEM24Canon(CStr(src.cells(r, 2).value2))
      ws.cells(outRow, 1).value2 = id: ws.cells(outRow, 2).value2 = FEM24Label(kind)
      If kind <> "WATER" Then ws.cells(outRow, 3).value2 = src.cells(r, 3).value2
      ws.cells(outRow, 4).value2 = src.cells(r, 4).value2
      ws.cells(outRow, 5).value2 = src.cells(r, 5).value2: ws.cells(outRow, 6).value2 = src.cells(r, 6).value2
      pr = PlanRow(id)
      If pr > 0 Then ws.cells(outRow, 7).value2 = water.cells(pr, 13).value2
      profile = ""
      If kind = "WATER" Then
        profile = Trim$(CStr(src.cells(r, 3).value2))
        If pr > 0 Then
          If Len(Trim$(CStr(water.cells(pr, 14).value2))) > 0 Then profile = Trim$(CStr(water.cells(pr, 14).value2))
        End If
        ws.cells(outRow, 8).value2 = profile
        If profile = "" Then profile = CStr(water.range("B5").value2)
        If val(CStr(src.cells(r, 5).value2)) <> 0 Then inherited = profile
      End If
      If val(CStr(src.cells(r, 5).value2)) = 0 Then
        ws.cells(outRow, 9).value2 = "実行しない"
      ElseIf kind = "WATER" Then
        ws.cells(outRow, 9).value2 = profile
      Else
        ws.cells(outRow, 9).value2 = "継承: " & inherited
      End If
      If val(CStr(src.cells(r, 5).value2)) <> 0 Then
        If UCase$(Trim$(CStr(water.range("B3").value2))) = "NONE" Then ws.cells(outRow, 9).value2 = "水条件なし"
      End If
    End If
  Next r
  ws.range("F8:F2006").WrapText = True
  If lastOut >= 8 Then
    ws.rows("8:" & CStr(lastOut)).AutoFit
    For r = 8 To lastOut
      If ws.rows(r).RowHeight < 28 Then ws.rows(r).RowHeight = 28
    Next r
  End If
End Sub

Public Sub FEM24StagesChanged(ByVal ws As Worksheet, ByVal target As range)
  Dim src As Worksheet, water As Worksheet, hit As range, c As range, rr As Long, id As Long, pr As Long, kind As String, text As String, e As Boolean
  e = Application.EnableEvents: Application.EnableEvents = False: FEM24StageError = ""
  On Error GoTo Failed
  Set src = ThisWorkbook.Worksheets("ステージ"): Set water = ThisWorkbook.Worksheets("水条件")
  Set hit = Intersect(target, ws.range("B8:H2006"))
  If hit Is Nothing Then GoTo Finish
  Snapshot
  NormalizeIds
  ' Ordinary fields first, then water assignments, so pasted rows are deterministic.
  For Each c In hit.cells
    If c.HasFormula Then Err.Raise vbObjectError + 2408, , "工程の入力に計算式は使用できません。"
    rr = c.row - 6
    Select Case c.column
      Case 2
        kind = FEM24Canon(CStr(c.value2))
        If kind = "" Then Err.Raise vbObjectError + 2409, , "操作名を選択してください。削除は上の削除ボタンで行います。"
        src.cells(rr, 2).value2 = kind
        If Len(CStr(src.cells(rr, 1).value2)) = 0 Then src.cells(rr, 1).value2 = MaxId() + 1
        If Len(CStr(src.cells(rr, 5).value2)) = 0 Then src.cells(rr, 5).value2 = 1
      Case 3
        If FEM24Canon(CStr(ws.cells(c.row, 2).value2)) = "WATER" And Len(CStr(c.value2)) > 0 Then Err.Raise vbObjectError + 2410, , "水位線はH列で指定します。C列は材料番号です。"
        src.cells(rr, 3).value2 = c.value2
      Case 4, 5, 6
        src.cells(rr, c.column).value2 = c.value2
    End Select
  Next c
  NormalizeIds
  For Each c In hit.cells
    If c.column = 7 Or c.column = 8 Then
      rr = c.row - 6: kind = FEM24Canon(CStr(src.cells(rr, 2).value2))
      If kind = "" Then Err.Raise vbObjectError + 2411, , "先に操作を選択してください。"
      id = CLng(src.cells(rr, 1).value2): text = Trim$(CStr(c.value2)): pr = PlanRow(id)
      If c.column = 7 Then
        If text <> "" And UCase$(text) <> "ONE_WAY" And UCase$(text) <> "FIXED_STRESS" And UCase$(text) <> "INHERIT" Then Err.Raise vbObjectError + 2412, , "継承/ONE_WAY/FIXED_STRESSを選択してください。"
        If pr = 0 And text <> "" Then pr = PlanRow(id, True)
        If pr > 0 Then water.cells(pr, 13).value2 = UCase$(text)
      Else
        If kind <> "WATER" And text <> "" Then Err.Raise vbObjectError + 2413, , "水位線はWATER工程だけに指定できます。"
        If kind = "WATER" Then
          src.cells(rr, 3).ClearContents
          If pr = 0 And text <> "" Then pr = PlanRow(id, True)
          If pr > 0 Then
            water.cells(pr, 14).value2 = text
            If text <> "" And Len(CStr(water.cells(pr, 13).value2)) = 0 Then water.cells(pr, 13).value2 = "INHERIT"
          End If
        End If
      End If
      If pr > 0 Then
        If Application.CountA(water.range(water.cells(pr, 13), water.cells(pr, 19))) = 0 Then water.range(water.cells(pr, 12), water.cells(pr, 20)).ClearContents
      End If
    End If
  Next c
  ValidateStages
  FEM24Invalidate
  ws.Unprotect: FEM24StagesRender ws: ws.Protect UserInterfaceOnly:=True, DrawingObjects:=False
  GoTo Finish
Failed:
  text = Err.description
  Restore
  ws.Unprotect: FEM24StagesRender ws: ws.Protect UserInterfaceOnly:=True, DrawingObjects:=False
  FEM24StageError = text
  If Not SuppressUserMessages Then MsgBox "入力を元に戻しました。" & vbCrLf & text, vbExclamation
Finish:
  Application.EnableEvents = e
End Sub

Public Sub FEM24StageUndo()
  Dim e As Boolean
  If AnalysisRunning Then Exit Sub
  If Not (ActiveSheet.parent Is ThisWorkbook) Then Exit Sub
  If Not mStageUndoValid Then Exit Sub
  If Not FEM24PageCurrent("施工工程") Then
    FEM24StageError = "元の工程が更新されています。工程画面を開き直して確認してください。"
    If Not SuppressUserMessages Then MsgBox FEM24StageError, vbInformation
    Exit Sub
  End If
  Application.CutCopyMode = False
  e = Application.EnableEvents: Application.EnableEvents = False
  On Error GoTo Failed
  Restore
  mStageUndoValid = False
  FEM24Invalidate
  FEM24RefreshStages
Finish:
  Application.EnableEvents = e
  Exit Sub
Failed:
  FEM24StageError = Err.description
  If Not SuppressUserMessages Then MsgBox FEM24StageError, vbExclamation
  Resume Finish
End Sub
Public Sub FEM24StageAdd(): Mutate "ADD": End Sub
Public Sub FEM24StageCopy(): Mutate "COPY": End Sub
Public Sub FEM24StageDelete(): Mutate "DELETE": End Sub
Public Sub FEM24StageUp(): Mutate "UP": End Sub
Public Sub FEM24StageDown(): Mutate "DOWN": End Sub

Private Sub Mutate(ByVal action As String)
  Dim src As Worksheet, water As Worksheet, r As Long, last As Long, id As Long, k As Long, pr As Long, nr As Long, tmp As Variant, e As Boolean, text As String
  If AnalysisRunning Or ActiveSheet.name <> "施工工程" Then Exit Sub
  If CStr(ActiveSheet.range("C4").value2) <> "工程の編集" Then
    FEM24StageError = "工程の追加・複製・削除・移動は「工程の編集」で対象行を選択してください。"
    If Not SuppressUserMessages Then MsgBox FEM24StageError, vbInformation
    Exit Sub
  End If
  If Not (ActiveSheet.parent Is ThisWorkbook) Then Exit Sub
  If Not FEM24PageCurrent("施工工程") Then
    FEM24StageError = "元の工程が更新されています。工程画面を開き直して確認してください。"
    If Not SuppressUserMessages Then MsgBox FEM24StageError, vbInformation
    Exit Sub
  End If
  Application.CutCopyMode = False
  e = Application.EnableEvents: Application.EnableEvents = False: FEM24StageError = ""
  On Error GoTo Failed
  Set src = ThisWorkbook.Worksheets("ステージ"): Set water = ThisWorkbook.Worksheets("水条件")
  Snapshot
  NormalizeIds
  r = ActiveCell.row - 6: last = src.cells(src.rows.count, 2).End(xlUp).row
  If action <> "ADD" Then
    If r < 2 Or r > last Then Err.Raise vbObjectError + 2414, , "操作する工程のセルを選択してください。"
    If Len(CStr(src.cells(r, 2).value2)) = 0 Then Err.Raise vbObjectError + 2414, , "空行は選択できません。"
  End If
  Select Case action
    Case "ADD", "COPY"
      If last >= 2000 Then Err.Raise vbObjectError + 2415, , "施工工程は1999行までです。"
      id = MaxId() + 1: nr = last + 1
      If action = "COPY" Then
        FSCopyTimeOptions CLng(src.cells(r, 1).value2), id
        Call FEM24RestoreCells(src.cells(nr, 1).Resize(1, 6), FEM24CaptureCells(src.cells(r, 1).Resize(1, 6)))
        pr = PlanRow(CLng(src.cells(r, 1).value2))
        If pr > 0 Then
          k = PlanRow(id, True)
          Call FEM24RestoreCells(water.cells(k, 12).Resize(1, 9), FEM24CaptureCells(water.cells(pr, 12).Resize(1, 9)))
          water.cells(k, 12).value2 = id
        End If
      Else
        src.cells(nr, 2).value2 = "GRAVITY": src.cells(nr, 4).value2 = 1: src.cells(nr, 5).value2 = 1
      End If
      src.cells(nr, 1).value2 = id: r = nr
    Case "DELETE"
      If Application.CountA(src.range("B2:B2000")) <= 1 Then Err.Raise vbObjectError + 2416, , "最後の工程は削除できません。"
      id = CLng(src.cells(r, 1).value2): pr = PlanRow(id)
      FSDeleteTimeOptions id
      If pr > 0 Then water.cells(pr, 12).Resize(1, 9).ClearContents
      If r < last Then Call FEM24RestoreCells(src.cells(r, 1).Resize(last - r, 6), FEM24CaptureCells(src.cells(r + 1, 1).Resize(last - r, 6)))
      src.cells(last, 1).Resize(1, 6).ClearContents
    Case "UP", "DOWN"
      k = r - 1
      If action = "DOWN" Then k = r + 1
      If k < 2 Or k > last Then GoTo Finish
      tmp = FEM24CaptureCells(src.cells(r, 1).Resize(1, 6))
      Call FEM24RestoreCells(src.cells(r, 1).Resize(1, 6), FEM24CaptureCells(src.cells(k, 1).Resize(1, 6)))
      Call FEM24RestoreCells(src.cells(k, 1).Resize(1, 6), tmp)
      r = k
  End Select
  ValidateStages
  FEM24Invalidate
  FEM24RefreshStages
  ActiveSheet.cells(r + 6, 2).Select
  GoTo Finish
Failed:
  text = Err.description: Restore
  FEM24RefreshStages
  FEM24StageError = text
  If Not SuppressUserMessages Then MsgBox "工程を元に戻しました。" & vbCrLf & text, vbExclamation
Finish:
  Application.EnableEvents = e
End Sub
