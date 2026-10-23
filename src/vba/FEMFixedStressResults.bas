Option Explicit
Private mRunFolder As String
Public FSFrameCaption As String

Public Sub FSResultsReset()
  Dim ws As Worksheet
  mRunFolder = "": FSFrameCaption = ""
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("時間履歴")
  If Not ws Is Nothing Then
    If Len(Trim$(FEMReadTextSetting("EXPORT_LOAD", ""))) = 0 Then ws.cells.ClearContents
  End If
  ThisWorkbook.names("_FSTimeCaption").Delete
  ThisWorkbook.names("_FSViewPrefix").Delete
  ThisWorkbook.names("_FSViewStage").Delete
  ThisWorkbook.names("_FSViewSignature").Delete
  On Error GoTo 0
End Sub

Private Sub EnsureRunFolder()
  Dim root As String, stem As String, n As Long
  If Len(mRunFolder) > 0 Then Exit Sub
  root = FEMIoTimeRoot()
  If Len(root) = 0 Then Err.Raise vbObjectError + 8130, , "時刻別結果の保存先を作成できません。"
  stem = root & Application.PathSeparator & "time_" & Format$(Now, "yyyymmdd_hhnnss")
  mRunFolder = stem
  Do While Dir$(mRunFolder, vbDirectory) <> ""
    n = n + 1: mRunFolder = stem & "_" & CStr(n)
  Loop
  MkDir mRunFolder
End Sub

Public Sub FSRecordState(ByVal stageId As Long, ByVal dt As Double, Optional ByVal branch As Boolean = False, Optional ByVal saveField As Boolean = True)
  Dim ws As Worksheet, row As Long, id As String, folder As String, relative As String, kind As String, i As Long
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("時間履歴"): On Error GoTo 0
  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = "時間履歴"
  End If
  id = "C" & Format$(FSAccepted, "000000")
  If branch Then id = "S" & Format$(FSAccepted, "000000") & "_" & CStr(stageId)
  If saveField Then
    EnsureRunFolder
    folder = mRunFolder & Application.PathSeparator & id
    For i = 1 To P3StageN
      If P3StageId(i) = stageId Then kind = P3StageKind(i): Exit For
    Next i
    If Not FEMIoExportTime(folder, stageId, kind, branch) Then Err.Raise vbObjectError + 8130, , "時刻別状態の保存に失敗しました。受理時刻=" & CStr(FSTime)
    relative = folder
    If Left$(folder, Len(ThisWorkbook.path) + 1) = ThisWorkbook.path & Application.PathSeparator Then relative = Mid$(folder, Len(ThisWorkbook.path) + 2)
  End If
  ws.range("A1:M1").value2 = Array("状態ID", "工程ID", "累積時刻[s]", "刻み[s]", "連成反復", "流体残差", "水量誤差", "力学残差", "最大変位[m]", "施工本線/照査枝", "保存先", "親状態ID", "Fs")
  row = ws.cells(ws.rows.count, 1).End(xlUp).row + 1
  ws.cells(row, 1).Resize(1, 13).value2 = Array(id, stageId, FSTime, dt, FSOuterIteration, FSFluidResidual, FSMassError, RelativeResidualFree, P3MaxTrialDisp, IIf(branch, "SRM_BRANCH", "MAINLINE"), relative, IIf(branch, "C" & Format$(FSAccepted, "000000"), ""), IIf(branch, FSS, 1#))
  ws.range("N1:U1").value2 = Array("貯留変化[m3]", "骨格体積変化[m3]", "指定圧流入量[m3]", "指定流束流入量[m3]", "閉合誤差[m3]", "施工持込水量[m3]", "掘削持出水量[m3]", "全水量の基準")
  If Not branch Then
    ws.cells(row, 14).Resize(1, 8).value2 = Array(FSStorageChange, FSSkeletonChange, FSBoundaryVolume, FSSourceVolume, FSVolumeClosure, FSBirthWater, FSDeathWater, IIf(FSVolumeKnown, "間隙率指定", "体積増減は相対値"))
    FSLedgerRecorded
  End If
  ws.columns("A").NumberFormat = "@"
  FSFrameCaption = IIf(branch, "SRM照査枝（水圧保持）", "施工本線") & " / 工程=" & CStr(stageId) & " / t=" & Format$(FSTime, "0.###") & " s（" & Format$(FSTime / 86400#, "0.###") & "日）"
  FSStoreCaption
End Sub

Public Sub FSStoreCaption()
  On Error Resume Next: ThisWorkbook.names("_FSTimeCaption").Delete: On Error GoTo 0
  ThisWorkbook.names.Add name:="_FSTimeCaption", RefersTo:="=""" & Replace$(FSFrameCaption, """", """""") & """", Visible:=False
End Sub

Public Function FSStoredCaption() As String
  Dim text As String
  On Error Resume Next: text = ThisWorkbook.names("_FSTimeCaption").RefersTo: On Error GoTo 0
  If Left$(text, 2) = "=""" Then FSStoredCaption = Replace$(Mid$(text, 3, Len(text) - 3), """""", """")
End Function

Public Sub FSShowTimeFrame()
  Dim ws As Worksheet, id As String, row As Long, match As Long, folder As String, caption As String
  Dim events As Boolean, quiet As Boolean, sheet As Worksheet
  If AnalysisRunning Then Exit Sub
  events = Application.EnableEvents: quiet = SuppressUserMessages
  FEM24Error = ""
  On Error GoTo Failed
  id = Trim$(CStr(ThisWorkbook.Worksheets("結果").range("B8").value2))
  Set ws = ThisWorkbook.Worksheets("時間履歴")
  For row = 2 To ws.cells(ws.rows.count, 1).End(xlUp).row
    If CStr(ws.cells(row, 1).value2) = id And Len(CStr(ws.cells(row, 11).value2)) > 0 Then
      If match > 0 Then FSCoupledInputError "状態IDが重複しています。"
      match = row
    End If
  Next row
  If match = 0 Then FSCoupledInputError "保存先のある状態IDを時間履歴から選択してください。"
  folder = CStr(ws.cells(match, 11).value2)
  If InStr(folder, ":") = 0 And Left$(folder, 2) <> "\\" Then folder = ThisWorkbook.path & Application.PathSeparator & folder
  caption = CStr(ws.cells(match, 10).value2) & " / 工程=" & CStr(ws.cells(match, 2).value2) & " / t=" & CStr(ws.cells(match, 3).value2) & " s / Fs=" & CStr(ws.cells(match, 13).value2)
  Application.EnableEvents = False: SuppressUserMessages = True
  P3LoadTimeFrame folder
  FSFrameCaption = caption: FSStoreCaption
  FEMViewerMarkClean
  FEM24Plot
  If Len(FEM24Error) > 0 Then Err.Raise vbObjectError + 8131, , FEM24Error
Finish:
  Application.EnableEvents = events: SuppressUserMessages = quiet
  Exit Sub
Failed:
  FEM24Error = Err.description
  If Not quiet Then MsgBox FEM24Error, vbExclamation
  Resume Finish
End Sub

Public Sub FSResultsResume(ByVal folder As String, ByVal stageId As Long)
  Dim ws As Worksheet, row As Long, i As Long, index As Long, id As String, count As Double, found As Boolean, path As String
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("時間履歴"): On Error GoTo 0
  If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = "時間履歴"
  End If
  For row = ws.cells(ws.rows.count, 1).End(xlUp).row To 2 Step -1
    id = CStr(ws.cells(row, 1).value2): count = Val(Mid$(Split(id, "_")(0), 2)): index = 0
    For i = 1 To P3StageN
      If P3StageId(i) = Val(CStr(ws.cells(row, 2).value2)) Then index = i: Exit For
    Next i
    If count > FSAccepted Or index > femIoResumeIndex Or index = 0 Then
      ws.rows(row).Delete
    ElseIf id = "C" & Format$(FSAccepted, "000000") Then
      found = True
    End If
  Next row
  If Not found Then
    ws.range("A1:M1").value2 = Array("状態ID", "工程ID", "累積時刻[s]", "刻み[s]", "連成反復", "流体残差", "水量誤差", "力学残差", "最大変位[m]", "施工本線/照査枝", "保存先", "親状態ID", "Fs")
    row = ws.cells(ws.rows.count, 1).End(xlUp).row + 1
    path = folder
    If Left$(folder, Len(ThisWorkbook.path) + 1) = ThisWorkbook.path & Application.PathSeparator Then path = Mid$(folder, Len(ThisWorkbook.path) + 2)
    ws.cells(row, 1).Resize(1, 13).value2 = Array("C" & Format$(FSAccepted, "000000"), stageId, FSTime, 0#, FSOuterIteration, FSFluidResidual, FSMassError, RelativeResidualFree, P3MaxTrialDisp, "MAINLINE", path, "", 1#)
  End If
  FSFrameCaption = "施工本線（保存状態から再開） / 工程=" & CStr(stageId) & " / t=" & CStr(FSTime) & " s"
  FSStoreCaption
End Sub

Public Sub FSStoreViewContext(ByVal prefix As Long, ByVal stageId As Long)
  On Error Resume Next
  ThisWorkbook.names("_FSViewPrefix").Delete: ThisWorkbook.names("_FSViewStage").Delete: ThisWorkbook.names("_FSViewSignature").Delete
  On Error GoTo 0
  ThisWorkbook.names.Add name:="_FSViewPrefix", RefersTo:="=" & CStr(prefix), Visible:=False
  ThisWorkbook.names.Add name:="_FSViewStage", RefersTo:="=" & CStr(stageId), Visible:=False
  ThisWorkbook.names.Add name:="_FSViewSignature", RefersTo:="=""" & FEMIoFrameSignature() & """", Visible:=False
End Sub

Public Function FSViewPrefix() As Long
  Dim prefix As String, signature As String, stage As Long, ws As Worksheet, row As Long, count As Long
  FSViewPrefix = -1
  On Error Resume Next: prefix = ThisWorkbook.names("_FSViewPrefix").RefersTo: On Error GoTo 0
  If Len(prefix) = 0 Then Exit Function
  signature = ThisWorkbook.names("_FSViewSignature").RefersTo
  If signature <> "=""" & FEMIoFrameSignature() & """" Then Err.Raise vbObjectError + 8194, , "時刻別結果と現在の入力が一致しません。受理状態を読み込み直してください。"
  stage = CLng(Val(Mid$(ThisWorkbook.names("_FSViewStage").RefersTo, 2)))
  Set ws = ThisWorkbook.Worksheets("結果節点")
  For row = 2 To ws.cells(ws.rows.count, 1).End(xlUp).row
    If Len(CStr(ws.cells(row, 1).value2)) > 0 Then
      If Val(CStr(ws.cells(row, 1).value2)) <> stage Then Err.Raise vbObjectError + 8194, , "時刻別の節点結果と保存工程が一致しません。"
    End If
  Next row
  Set ws = ThisWorkbook.Worksheets("ステージ")
  For row = 2 To ws.cells(ws.rows.count, 1).End(xlUp).row
    If Len(CStr(ws.cells(row, 2).value2)) > 0 Then count = count + 1
  Next row
  FSViewPrefix = CLng(Val(Mid$(prefix, 2)))
  If FSViewPrefix < 1 Or FSViewPrefix > count Then Err.Raise vbObjectError + 8194, , "時刻別結果の施工順序が一致しません。"
End Function
