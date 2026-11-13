Option Explicit
Private mSampleValid As Boolean, mReady As Boolean, mEnabled As Boolean, mWindow As Long, mLimit As Long, mBase As Long
Private mLast As Long, mPrevious As Double, mResidual As Double, mStoppedAt As Long
Private mAt() As Long, mValue() As Double
Private mReservedSample As Long

' Reset for every increment/retry and each coupled mechanical solve. Trial
' line-search/watchdog states must never enter this accepted-iterate history.
Public Sub NewtonProgressReset(ByVal baseLimit As Long)
    Dim i As Long
    mEnabled = FEMReadSetting("NEWTON_PROGRESS_CONTINUE", 1#) >= 0.5
    mWindow = CLng(FEMReadSetting("NEWTON_PROGRESS_WINDOW", 20#))
    If mWindow < 2 Or mWindow > 1000 Then Err.Raise vbObjectError + 3027, , "Newtonの進捗確認間隔は2～1000です。"
    mBase = baseLimit: mLimit = baseLimit: mLast = -1: mPrevious = 0#: mResidual = 0#: mStoppedAt = -1: mSampleValid = False
    mReservedSample = -1
    ReDim mAt(0 To mWindow): ReDim mValue(0 To mWindow)
    For i = 0 To mWindow: mAt(i) = -1: Next i
    mReady = True
End Sub

Public Function NewtonProgressReduction(ByVal current As Double, ByVal previous As Double) As Boolean
    If Not mReady Or Not mEnabled Then Exit Function
    If Not P2IsFinite(current) Or Not P2IsFinite(previous) Then Exit Function
    If previous <= 0# Or current < 0# Then Exit Function
    ' Even a very slow representable reduction counts. Do not impose a
    ' percentage-rate cutoff while the accepted residual is still falling.
    NewtonProgressReduction = current < previous
End Function

Public Sub NewtonProgressReserveRecovery(ByVal baseLimit As Long, ByVal workIteration As Long, ByVal additionalWork As Long)
    Dim reference As Double, nextLimit As Double, oldLimit As Long
    If Not mReady Then NewtonProgressReset baseLimit
    If Not mEnabled Or Not mSampleValid Then Exit Sub
    If additionalWork < 1 Or additionalWork > 1000 Then Exit Sub
    If mLast < 0 Or mLast = mReservedSample Then Exit Sub
    ' At most the just-completed linear solve may separate this request from
    ' a real accepted state. Rejected look-ahead points are never evidence.
    If workIteration < mLast Then Exit Sub
    If workIteration - mLast > 1 Then Exit Sub
    If CDbl(workIteration) + CDbl(additionalWork) < CDbl(mLimit) Then Exit Sub
    ' A short configured window can contain no predecessor after a bounded
    ' multi-solve correction. Use the last actual accepted predecessor then.
    If Not ProgressReference(reference) Then reference = mPrevious
    If Not NewtonProgressReduction(mResidual, reference) Then Exit Sub
    If Not NewtonProgressReduction(mResidual, mPrevious) Then Exit Sub
    nextLimit = CDbl(workIteration) + CDbl(additionalWork) + CDbl(mWindow)
    If nextLimit >= 2147483000# Then Exit Sub
    oldLimit = mLimit: mLimit = CLng(nextLimit): mReservedSample = mLast
    FEMAppendRunLog "NEWTON_PROGRESS", "RESERVE_RECOVERY", "accepted_iteration=" & CStr(mLast) & ";work_iteration=" & CStr(workIteration) & ";additional_work=" & CStr(additionalWork) & ";old_limit=" & CStr(oldLimit) & ";next_check=" & CStr(mLimit) & ";residual=" & Format$(mResidual, "0.000000000000E+00")
End Sub

Public Sub NewtonProgressRecord(ByVal iteration As Long, ByVal residual As Double)
    Dim slot As Long
    If Not mReady Then Exit Sub
    If iteration <= mLast Then Exit Sub
    If Not P2IsFinite(residual) Or residual < 0# Then mSampleValid = False: Exit Sub
    mSampleValid = True
    mPrevious = mResidual: mResidual = residual: mLast = iteration
    slot = iteration Mod (mWindow + 1): mAt(slot) = iteration: mValue(slot) = residual
End Sub

Private Function ProgressReference(ByRef reference As Double) As Boolean
    Dim i As Long, earliest As Long, first As Long
    If mLast < 1 Then Exit Function
    earliest = mLast - mWindow: If earliest < 0 Then earliest = 0
    first = mLast
    For i = 0 To mWindow
        If mAt(i) >= earliest And mAt(i) < first Then first = mAt(i): reference = mValue(i)
    Next i
    ProgressReference = first < mLast
End Function

Public Function NewtonProgressLimit(ByVal baseLimit As Long, ByVal iteration As Long) As Long
    Dim reference As Double, improving As Boolean, oldLimit As Long
    If Not mReady Then NewtonProgressReset baseLimit
    If Not mEnabled Then NewtonProgressLimit = baseLimit: Exit Function
    If iteration >= mLimit Then
        improving = False
        If mSampleValid And mLast = iteration And ProgressReference(reference) Then improving = NewtonProgressReduction(mResidual, reference) And NewtonProgressReduction(mResidual, mPrevious)
        If improving And iteration < 2147483000 Then
            oldLimit = mLimit: mLimit = iteration + mWindow
            FEMAppendRunLog "NEWTON_PROGRESS", "EXTEND", "iteration=" & CStr(iteration) & ";old_limit=" & CStr(oldLimit) & ";next_check=" & CStr(mLimit) & ";window_reference=" & Format$(reference, "0.000000000000E+00") & ";residual=" & Format$(mResidual, "0.000000000000E+00") & ";sample_valid=" & CStr(mSampleValid)
        ElseIf mStoppedAt <> iteration Then
            mStoppedAt = iteration
            FEMAppendRunLog "NEWTON_PROGRESS", "NO_PROGRESS", "iteration=" & CStr(iteration) & ";window_reference=" & Format$(reference, "0.000000000000E+00") & ";residual=" & Format$(mResidual, "0.000000000000E+00") & ";sample_valid=" & CStr(mSampleValid)
        End If
    End If
    NewtonProgressLimit = mLimit
End Function

Public Sub NewtonProgressEnsureSettings()
    Dim ws As Worksheet, page As Worksheet, keys As Variant, values As Variant, labels As Variant, notes As Variant, i As Long, r As Long, pr As Long, found As Boolean
    keys = Array("NEWTON_PROGRESS_CONTINUE", "NEWTON_PROGRESS_WINDOW")
    values = Array(1#, 20#)
    labels = Array("残差改善中のNewton反復を継続", "延長後の進捗確認間隔")
    notes = Array("1=改善中は固定回数で終了しない / 0=従来の回数上限。収束許容値は変更しない", "2～1000反復。停滞・増大・非有限・手動中断時は延長しない")
    Set ws = ThisWorkbook.Worksheets("設定"): Set page = ThisWorkbook.Worksheets("解析設定")
    ' Excel does not persist UserInterfaceOnly across reopen. This migration
    ' also runs from P0 before the UI installer, including checkpoint resumes.
    If ws.ProtectContents Then ws.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    If page.ProtectContents Then page.Protect Password:="", DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    For i = 0 To UBound(keys)
        found = False
        For r = 5 To ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
            If CStr(ws.Cells(r, 5).Value2) = keys(i) Then found = True: Exit For
        Next r
        If Not found Then
            r = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row + 1
            ws.Cells(r, 1).Value2 = "反復継続": ws.Cells(r, 2).Value2 = labels(i): ws.Cells(r, 3).Value2 = values(i): ws.Cells(r, 4).Value2 = notes(i): ws.Cells(r, 5).Value2 = keys(i)
        End If
        found = False
        For pr = 8 To page.Cells(page.Rows.Count, 6).End(xlUp).Row
            If CStr(page.Cells(pr, 6).Value2) = keys(i) Then found = True: Exit For
        Next pr
        If Not found Then
            pr = page.Cells(page.Rows.Count, 6).End(xlUp).Row + 1: If pr < 50 Then pr = 50
            page.Range("B10:G10").Copy Destination:=page.Cells(pr, 2)
            page.Cells(pr, 2).Value2 = labels(i): page.Cells(pr, 3).Value2 = ws.Cells(r, 3).Value2
            page.Cells(pr, 4).ClearContents: page.Cells(pr, 5).Value2 = notes(i): page.Cells(pr, 6).Value2 = keys(i): page.Cells(pr, 7).Value2 = values(i)
            page.Cells(pr, 3).Validation.Delete: page.Rows(pr).RowHeight = 48
        End If
    Next i
    Application.CutCopyMode = False: FEMInvalidateSettingCache
End Sub
