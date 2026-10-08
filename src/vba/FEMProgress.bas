Option Explicit

Private stageText As String
Private lastText As String
Private savedStatusBar As Variant
Private savedDisplayStatusBar As Boolean
Private progressActive As Boolean
Private progressPumping As Boolean

Public Sub FEMProgressBegin()
    On Error Resume Next
    If progressActive Then FEMProgressFinish False, vbNullString
    savedStatusBar = Application.StatusBar
    savedDisplayStatusBar = Application.DisplayStatusBar
    progressActive = True
    stageText = vbNullString
    lastText = vbNullString
    Application.DisplayStatusBar = True
    On Error GoTo 0
    FEMProgressPhase "解析準備", "入力データを確認しています"
End Sub

Public Sub FEMProgressPhase(ByVal phaseText As String, Optional ByVal detailText As String = "")
    Dim nextText As String
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    If Not progressActive Or progressPumping Then Exit Sub
    nextText = phaseText & vbLf & detailText
    If nextText = lastText Then Exit Sub
    On Error GoTo UpdateFailed
    ' ステージとFs試行の切替時だけ描画イベントを処理する。
    FEMProgressFlush "2DSoilFEM  " & phaseText & "  " & detailText
    lastText = nextText
    Exit Sub
UpdateFailed:
    errorNumber = Err.Number
    errorSource = Err.source
    errorDescription = Err.Description
    ' 表示の失敗は解析へ影響させないが、ユーザー中断は上位へ渡す。
    If errorNumber = 18 Then Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Sub FEMProgressFlush(ByVal displayText As String)
    Dim previousEnableEvents As Boolean
    Dim eventsSaved As Boolean
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    If progressPumping Then Exit Sub
    progressPumping = True
    On Error GoTo FlushFailed
    previousEnableEvents = Application.EnableEvents
    eventsSaved = True
    ' DoEvents中のワークシートイベントを抑止する。解析再実行は既存AnalysisRunningガードで防ぐ。
    Application.EnableEvents = False
    DoEvents
    Application.StatusBar = displayText
    DoEvents
    GoTo FlushDone
FlushFailed:
    errorNumber = Err.Number
    errorSource = Err.source
    errorDescription = Err.Description
    Resume FlushDone
FlushDone:
    On Error Resume Next
    If eventsSaved Then Application.EnableEvents = previousEnableEvents
    progressPumping = False
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorDescription
End Sub

Public Sub FEMProgressStage(ByVal stageNo As Long, ByVal kindText As String)
    stageText = "ステージ " & CStr(stageNo) & "  " & kindText
    FEMProgressPhase stageText, "処理中"
End Sub

Public Sub FEMProgressFs(ByVal trialNo As Long, ByVal trialFs As Double)
    FEMProgressPhase stageText, "SRM探索  試行 " & CStr(trialNo) & "   Fs = " & Format$(trialFs, "0.000")
End Sub

Public Sub FEMProgressFinish(ByVal succeeded As Boolean, ByVal detailText As String)
    On Error Resume Next
    If Not progressActive Then Exit Sub
    Application.StatusBar = savedStatusBar
    Application.DisplayStatusBar = savedDisplayStatusBar
    progressActive = False
    lastText = vbNullString
End Sub
