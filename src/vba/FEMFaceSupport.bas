Option Explicit
' Massless, compression-only normal support attached to a moving vertical face.
' gap = sense*(u_slave-(1+t)*u_upper+t*u_lower), t=(ys-yu)/(yu-yl).
' Its three-node action/reaction conserves both force and moment. No ground fixity.
Private mCount As Long, mIds() As Long, mInputId() As Long, mK() As Double, mSense() As Double
Private mDof() As Long, mC() As Double, mFactorClosed() As Boolean, mFactorKnown As Boolean
Private mReadOK As Boolean, mMapped As Boolean
Public FaceSupportFingerprint As Double, FaceSupportSignature As String

Public Sub FaceSupportInstall()
  Dim ws As Worksheet
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("壁面支持"): On Error GoTo 0
  If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add: ws.name = "壁面支持"
  ws.range("A1:G1").value2 = Array("ID", "盛土節点", "壁上節点", "壁下節点", "剛性[kN/m]", "法線1/-1", "備考")
  ws.columns("A:F").ColumnWidth = 16: ws.columns("G").ColumnWidth = 50
  ws.range("I1").value2 = "既存の鉛直壁へ取り付く無質量・圧縮のみの仮支持。水平法線方向だけを支持し、鉛直方向は自由。3節点は同じX座標。壁の移動・回転を許し、反力を壁へ返す。部材耐力は別途照査。再メッシュ時は節点を再設定。"
  ws.columns("I").ColumnWidth = 70: ws.range("I1").WrapText = True
End Sub

Public Function FaceSupportHasInput() As Boolean
  Dim ws As Worksheet, j As Long, last As Long, r As Long
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("壁面支持"): On Error GoTo 0
  If ws Is Nothing Then Exit Function
  For j = 1 To 6
    r = ws.cells(ws.rows.count, j).End(xlUp).row
    If r > last Then last = r
  Next j
  If last >= 2 Then FaceSupportHasInput = Application.CountA(ws.range("A2:F" & CStr(last))) > 0
End Function

Private Sub SupportFail(ByVal message As String)
  Err.Raise vbObjectError + 3494, "FEMFaceSupport", message
End Sub

Private Function SupportNumber(ByVal value As Variant) As Double
  If IsError(value) Or IsEmpty(value) Then SupportFail "壁面支持の数値が未入力です。"
  If Not IsNumeric(value) Then SupportFail "壁面支持には数値を入力してください。"
  SupportNumber = CDbl(value)
  If Not P2IsFinite(SupportNumber) Then SupportFail "壁面支持には有限値が必要です。"
End Function

Public Sub FaceSupportRead()
  Dim ws As Worksheet, stages As Worksheet, v As Variant, r As Long, j As Long, n As Double
  Dim last As Long, seen As Object, key As String, mode As String
  Dim connected() As Boolean, active() As Boolean, continuum() As Boolean, e As Long, i As Long
  mCount = 0: mReadOK = False: mMapped = False: mFactorKnown = False: FaceSupportFingerprint = 0#: FaceSupportSignature = ""
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("壁面支持"): On Error GoTo 0
  If ws Is Nothing Then Exit Sub
  For j = 1 To 6
    r = ws.cells(ws.rows.count, j).End(xlUp).row
    If r > last Then last = r
  Next j
  If last < 2 Then Exit Sub
  v = ws.range("A2:F" & CStr(last)).value2
  ReDim mIds(0 To last - 2, 0 To 2): ReDim mInputId(0 To last - 2): ReDim mK(0 To last - 2): ReDim mSense(0 To last - 2)
  Set seen = CreateObject("Scripting.Dictionary")
  For r = 1 To UBound(v, 1)
    If Application.CountA(ws.range(ws.cells(r + 1, 1), ws.cells(r + 1, 6))) = 0 Then GoTo NextInput
    n = SupportNumber(v(r, 1))
    If n < 1# Or n > 2147483647# Or n <> Fix(n) Then SupportFail "壁面支持IDは正の整数です。"
    key = CStr(n): If seen.Exists(key) Then SupportFail "壁面支持IDが重複しています。"
    seen.Add key, True
    mInputId(mCount) = CLng(n)
    For j = 0 To 2
      n = SupportNumber(v(r, j + 2))
      If n < 1# Or n > P6InputNodeCount Or n <> Fix(n) Then SupportFail "壁面支持の節点IDが範囲外です。"
      mIds(mCount, j) = CLng(n) - 1
      FaceSupportFingerprint = P6FingerprintStep(FaceSupportFingerprint, n)
    Next j
    If mIds(mCount, 0) = mIds(mCount, 1) Or mIds(mCount, 0) = mIds(mCount, 2) Or mIds(mCount, 1) = mIds(mCount, 2) Then SupportFail "壁面支持の3節点は別の節点が必要です。"
    mK(mCount) = SupportNumber(v(r, 5)): mSense(mCount) = SupportNumber(v(r, 6))
    If mK(mCount) <= 0# Then SupportFail "壁面支持剛性は正数です。"
    If Abs(mSense(mCount)) <> 1# Then SupportFail "壁面支持の法線方向は1または-1です。"
    FaceSupportFingerprint = P6FingerprintStep(FaceSupportFingerprint, P6QuantizeScaled(mK(mCount), 1000000#))
    FaceSupportFingerprint = P6FingerprintStep(FaceSupportFingerprint, mSense(mCount))
    mCount = mCount + 1
NextInput:
  Next r
  If mCount = 0 Then Exit Sub
  If P6ReadSetting("CONSOL_ENABLE", 0#) <> 0# Then SupportFail "壁面支持は静的解析のみです。圧密との併用は未対応です。"
  ReDim connected(0 To P6InputNodeCount - 1): ReDim active(1 To P6InputMaterialCount): ReDim continuum(1 To P6InputMaterialCount)
  For r = 1 To P6InputMaterialCount
    active(P6InputMaterialId(r)) = P6InputMaterialInitialActive(r)
    continuum(P6InputMaterialId(r)) = (P6InputMaterialKind(r) <> "JOINT")
  Next r
  For e = 1 To P6InputElementCount
    ' Supports attach to continuum nodes. JOINT has six actual nodes and two
    ' optional zero fields; neither those fields nor joint-only nodes qualify.
    If active(P6InputElementMaterial(e)) And continuum(P6InputElementMaterial(e)) Then
      For j = 1 To 8: connected(P6InputElementNode(e, j) - 1) = True: Next j
    End If
  Next e
  For i = 0 To mCount - 1
    For j = 0 To 2
      If Not connected(mIds(i, j)) Then SupportFail "壁面支持の全節点を、初期から有効な連続体要素へ接続してください。"
    Next j
  Next i
  ReDim mDof(0 To mCount - 1, 0 To 2): ReDim mC(0 To mCount - 1, 0 To 2): ReDim mFactorClosed(0 To mCount - 1)
  FaceSupportSignature = "|FACE=" & Format$(FaceSupportFingerprint, "0")
  ' Restrict the first release to fixed-topology, small-strain static stages.
  Set stages = ThisWorkbook.Worksheets("ステージ")
  For r = 2 To stages.cells(stages.rows.count, 1).End(xlUp).row
    If Val(CStr(stages.cells(r, 5).value2)) <> 0# Then
      mode = UCase$(Trim$(CStr(stages.cells(r, 2).value2)))
      If mode <> "GRAVITY" And mode <> "LOAD" And mode <> "UNLOAD" And mode <> "APPLY" And mode <> "SRM" Then SupportFail "壁面支持は静的GRAVITY/LOAD/UNLOAD/SRMのみです。"
    End If
  Next r
  Set ws = ThisWorkbook.Worksheets("水条件")
  If Application.CountIf(ws.range("M12:M211"), "FIXED_STRESS") > 0 Then SupportFail "壁面支持と時間連成の併用は未対応です。"
  mReadOK = True
End Sub

Public Sub FaceSupportMap()
  Dim i As Long, j As Long, n As Long, top As Long, lower As Long, t As Double
  mMapped = False
  If mCount > 0 And Not mReadOK Then SupportFail "壁面支持の入力検査が完了していません。"
  For i = 0 To mCount - 1
    top = mIds(i, 1): lower = mIds(i, 2)
    If y(top) <= y(lower) Then SupportFail "壁面支持の上側・下側節点の標高を確認してください。"
    If Abs(x(top) - x(lower)) > 0.00000001 Or Abs(x(mIds(i, 0)) - x(top)) > 0.00000001 Then SupportFail "壁面支持の3節点は同じX座標の鉛直線上に指定してください。"
    t = (y(mIds(i, 0)) - y(top)) / (y(top) - y(lower))
    mC(i, 0) = mSense(i): mC(i, 1) = -mSense(i) * (1# + t): mC(i, 2) = mSense(i) * t
    For j = 0 To 2
      n = FEMFreeNodeMap(mIds(i, j))
      If n < 0 Then SupportFail "壁面支持は有効な連続体節点へ取り付けてください。"
      mDof(i, j) = 2 * P6GetInternalFreeNode(n)
    Next j
  Next i
  mMapped = True
End Sub

Private Function SupportGap(ByVal i As Long) As Double
  Dim j As Long
  For j = 0 To 2: SupportGap = SupportGap + mC(i, j) * TDisp(mDof(i, j)): Next j
End Function

Public Function FaceSupportBandWidth(ByVal original As Long) As Long
  Dim i As Long, j As Long, k As Long, d As Long
  FaceSupportBandWidth = original
  FaceSupportMap
  For i = 0 To mCount - 1
    For j = 0 To 2
      For k = 0 To 2
        d = Abs(mDof(i, j) - mDof(i, k))
        If d > FaceSupportBandWidth Then FaceSupportBandWidth = d
      Next k
    Next j
  Next i
End Function

Public Sub FaceSupportRCMEdges(ByRef adjacency() As Collection)
  Dim i As Long, j As Long, k As Long, a As Long, b As Long, item As Variant, found As Boolean
  For i = 0 To mCount - 1
    For j = 0 To 2
      a = FEMFreeNodeMap(mIds(i, j))
      For k = 0 To 2
        b = FEMFreeNodeMap(mIds(i, k))
        If a <> b Then
          found = False
          For Each item In adjacency(a)
            If CLng(item) = b Then found = True: Exit For
          Next item
          If Not found Then adjacency(a).Add b
        End If
      Next k
    Next j
  Next i
End Sub

Public Function FaceSupportCandidateBand(ByRef ordering() As Long, ByVal original As Long) As Long
  Dim i As Long, j As Long, k As Long, a As Long, b As Long, d As Long
  FaceSupportCandidateBand = original
  For i = 0 To mCount - 1
    For j = 0 To 2
      a = ordering(FEMFreeNodeMap(mIds(i, j)))
      For k = 0 To 2
        b = ordering(FEMFreeNodeMap(mIds(i, k))): d = 2 * Abs(a - b)
        If d > FaceSupportCandidateBand Then FaceSupportCandidateBand = d
      Next k
    Next j
  Next i
End Function

Public Sub FaceSupportAddForce(ByRef internal() As Double)
  Dim i As Long, j As Long, gap As Double
  FaceSupportMap
  For i = 0 To mCount - 1
    gap = SupportGap(i)
    If gap < 0# Then
      For j = 0 To 2: internal(mDof(i, j)) = internal(mDof(i, j)) + mK(i) * gap * mC(i, j): Next j
    End If
  Next i
End Sub

Public Function FaceSupportContactChanged() As Boolean
  Dim i As Long
  If mCount = 0 Then Exit Function
  If Not mFactorKnown Then FaceSupportContactChanged = True: Exit Function
  For i = 0 To mCount - 1
    If (SupportGap(i) <= 0#) <> mFactorClosed(i) Then FaceSupportContactChanged = True: Exit Function
  Next i
End Function

Public Sub FaceSupportAddBand()
  Dim i As Long, j As Long, k As Long, r As Long, c As Long
  FaceSupportMap
  For i = 0 To mCount - 1
    mFactorClosed(i) = (SupportGap(i) <= 0#)
    If mFactorClosed(i) Then
      For j = 0 To 2
        For k = 0 To 2
          r = mDof(i, j): c = mDof(i, k)
          If c >= r Then OriginalMat(r, c - r) = OriginalMat(r, c - r) + mK(i) * mC(i, j) * mC(i, k)
        Next k
      Next j
    End If
  Next i
  mFactorKnown = True
End Sub

Public Sub FaceSupportAugmentCSR()
  Dim extra As Object, cols As Object, r As Long, c As Long, i As Long, j As Long, k As Long, p As Long, n As Long
  Dim key As Variant, rows() As Long, columns() As Long, found As Boolean
  If mCount = 0 Then Exit Sub
  FaceSupportMap
  Set extra = CreateObject("Scripting.Dictionary")
  For i = 0 To mCount - 1
    For j = 0 To 2
      r = mDof(i, j)
      If Not extra.Exists(CStr(r)) Then Set cols = CreateObject("Scripting.Dictionary"): extra.Add CStr(r), cols
      Set cols = extra(CStr(r))
      For k = 0 To 2
        c = mDof(i, k): found = False
        For p = P6CSRRowPtr(r) To P6CSRRowPtr(r + 1) - 1
          If P6CSRColumnIndex(p) = c Then found = True: Exit For
        Next p
        If Not found Then cols(CStr(c)) = c
      Next k
    Next j
  Next i
  n = P6CSRNNZ
  For Each key In extra.keys: Set cols = extra(key): n = n + cols.count: Next key
  ReDim rows(0 To nDof): ReDim columns(0 To n - 1): n = 0
  For r = 0 To lastDof
    rows(r) = n
    For p = P6CSRRowPtr(r) To P6CSRRowPtr(r + 1) - 1: columns(n) = P6CSRColumnIndex(p): n = n + 1: Next p
    If extra.Exists(CStr(r)) Then
      Set cols = extra(CStr(r))
      For Each key In cols.keys: columns(n) = CLng(cols(key)): n = n + 1: Next key
    End If
  Next r
  rows(nDof) = n: P6CSRNNZ = n: P6CSRRowPtr = rows: P6CSRColumnIndex = columns
End Sub

Public Sub FaceSupportAddCSR()
  Dim i As Long, j As Long, k As Long, r As Long, c As Long, p As Long, found As Boolean
  FaceSupportMap
  For i = 0 To mCount - 1
    mFactorClosed(i) = (SupportGap(i) <= 0#)
    If mFactorClosed(i) Then
      For j = 0 To 2
        r = mDof(i, j)
        For k = 0 To 2
          c = mDof(i, k): found = False
          For p = P6CSRRowPtr(r) To P6CSRRowPtr(r + 1) - 1
            If P6CSRColumnIndex(p) = c Then
              P6CSROriginalValues(p) = P6CSROriginalValues(p) + mK(i) * mC(i, j) * mC(i, k): found = True: Exit For
            End If
          Next p
          If Not found Then SupportFail "壁面支持のCSR項がありません。"
        Next k
      Next j
    End If
  Next i
  mFactorKnown = True
End Sub

Private Function SupportResultsReady() As Boolean
  If mCount = 0 Or Not mReadOK Or Not mMapped Or lastDof < 0 Then Exit Function
  On Error GoTo NotReady
  SupportResultsReady = (LBound(TDisp) = 0 And UBound(TDisp) >= lastDof)
NotReady:
End Function

Public Sub FaceSupportWriteResults()
  Dim ws As Worksheet, i As Long, gap As Double
  On Error Resume Next: Set ws = ThisWorkbook.Worksheets("壁面支持結果"): On Error GoTo 0
  If Not SupportResultsReady() Then
    If Not ws Is Nothing Then
      ws.cells.ClearContents
      ws.range("A1").value2 = IIf(mCount = 0, "壁面支持の入力なし", "壁面支持の入力・自由度・変位の準備未完了：結果なし")
    End If
    Exit Sub
  End If
  If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add: ws.name = "壁面支持結果"
  FaceSupportMap
  ws.cells.ClearContents
  ws.range("A1:H1").value2 = Array("支持ID", "盛土節点", "壁上節点", "壁下節点", "法線相対変位[m]", "盛土への支持力X[kN]", "ばね剛性[kN/m]", "接触状態")
  For i = 0 To mCount - 1
    gap = SupportGap(i)
    ws.cells(i + 2, 1).Resize(1, 8).value2 = Array(mInputId(i), mIds(i, 0) + 1, mIds(i, 1) + 1, mIds(i, 2) + 1, gap, -mK(i) * IIf(gap < 0#, gap, 0#) * mSense(i), mK(i), IIf(gap <= 0#, "接触", "離間"))
  Next i
  ws.range("J1").value2 = "仮の無質量・圧縮のみ壁面支持。地盤固定なし、壁に作用反作用を返す。部材耐力の検証は別途必要。"
  ws.columns("A:H").AutoFit: ws.columns("J").ColumnWidth = 70: ws.range("J1").WrapText = True
End Sub
