Option Explicit
' Physical pore water is independent of the mechanical MIXED_UP pressure.
' All stored mechanical Stmat values remain compression-positive TOTAL stress.
Public GWActive As Boolean
Public GWContext As String
Private mOverrideProfile As String
Private mMode As String, mProfile As String, mGamma As Double, mAuto As Boolean
Private mWs As Worksheet, mReady As Boolean, mFlowSolved As Boolean
Private mBasis() As Boolean, mK() As Double, mMatProfile() As String
Private mTarget() As Double, mCurrent() As Double, mCommitted() As Double, mStart() As Double
Private mWaterForce() As Double, mHeads() As Double, mFx() As Double, mFy() As Double, mFlowReaction() As Double
Private mBoundary As Object, mEdges As Object, mOwners As Object, mSecondOwners As Object, mHoles As Object
Private mFaceFlows() As Double, mLocalError As Double, mFluxCorrection As Double
Private mRel As Double, mBalance As Double, mSeconds As Double, mSource As String
Private mInflow As Double, mOutflow As Double, mSig As String
Private mSurfaceRows As Collection
Private mHeadFixed() As Boolean
Private mTemporal As Boolean, mTimeFirst As String, mTimeLast As String, mTimeFraction As Double

Public Function GWEnabled() As Boolean
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("水条件")
  On Error GoTo 0
  If ws Is Nothing Then Exit Function
  GWEnabled = (UCase$(Trim$(CStr(ws.range("B3").value2))) <> "NONE" And Len(Trim$(CStr(ws.range("B3").value2))) > 0)
End Function

Private Sub GWFail(ByVal message As String)
  If AnalysisRunning Then SetAnalysisFailure RESULT_INPUT_ERROR, message, vbObjectError + 7920, -1, -1, CurrentIncrement, CurrentIteration
  Err.Raise vbObjectError + 7920, "FEMWater", message
End Sub

Private Function GWNum(ByVal value As Variant, ByVal field As String) As Double
  If IsError(value) Or IsEmpty(value) Then GWFail field & "を入力してください。"
  If Not IsNumeric(value) Then GWFail field & "は数値で入力してください。"
  GWNum = CDbl(value)
  If Not P2IsFinite(GWNum) Then GWFail field & "は有限値が必要です。"
End Function

Private Function GWId(ByVal value As Variant, ByVal field As String) As Long
  Dim number As Double
  number = GWNum(value, field)
  If number <> Fix(number) Or number < 1# Or number > 2147483647# Then GWFail field & "は正の整数で入力してください。"
  GWId = CLng(number)
End Function

Public Function GWInputSignature() As String
  Dim names As Variant, item As Variant, ws As Worksheet, values As Variant
  Dim i As Long, j As Long, k As Long, h As Double, t As String, lastRow As Long, cols As Long
  h = 17#: names = Array("水条件", "節点データ", "要素データ", "材料データ", "ステージ", "接合")
  For Each item In names
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CStr(item))
    On Error GoTo 0
    If Not ws Is Nothing Then
      Select Case CStr(item)
        Case "水条件": lastRow = 619: cols = 38
        Case "節点データ": lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row: cols = 9
        Case "要素データ": lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row: cols = 10
        Case "材料データ": lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row: cols = 13
        Case "ステージ": lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row: cols = 6
        Case "接合": lastRow = ws.cells(ws.rows.count, 1).End(xlUp).row: cols = 7
      End Select
      If lastRow < 2 Then lastRow = 2
      values = ws.range(ws.cells(1, 1), ws.cells(lastRow, cols)).value2
      For i = 1 To UBound(values, 1)
        For j = 1 To UBound(values, 2)
          If CStr(item) = "水条件" Then
            If Not ((j = 2 And i >= 3 And i <= 6) Or (i >= 12 And i <= 211 And (j <= 3 Or (j >= 6 And j <= 9))) Or (i >= 216 And i <= 415 And j <= 5) Or (i >= 12 And i <= 211 And j >= 12 And j <= 19 And GWPlanRowFilled(values, i)) Or (i >= 12 And i <= 211 And ((j >= 22 And j <= 26) Or j = 28)) Or (i >= 12 And i <= 211 And j >= 29 And j <= 37) Or (i >= 216 And i <= 415 And j >= 22 And j <= 27) Or (i >= 420 And i <= 619 And j >= 22 And j <= 24)) Then GoTo NextSigCell
          Else
            If i = 1 Then GoTo NextSigCell
          End If
          If IsError(values(i, j)) Then
            t = "#ERROR"
          Else
            t = CStr(values(i, j)) & ChrW$(31)
          End If
          For k = 1 To Len(t)
            h = h * 131# + (AscW(mid$(t, k, 1)) And &HFFFF&)
            h = h - Fix(h / 2147483629#) * 2147483629#
          Next k
NextSigCell:
        Next j
      Next i
    End If
  Next item
  GWInputSignature = Format$(h, "0")
End Function

Private Function GWPlanRowFilled(ByRef values As Variant, ByVal row As Long) As Boolean
  Dim j As Long
  For j = 12 To 19
    If IsError(values(row, j)) Then GWPlanRowFilled = True: Exit Function
    If Len(Trim$(CStr(values(row, j)))) > 0 Then GWPlanRowFilled = True: Exit Function
  Next j
End Function

Public Sub GWValidateStagePlans()
  Dim w As Worksheet, stages As Worksheet, seen As Object, r As Long, s As Long, count As Long, id As Long, sid As Long, matches As Long, allMatches As Long, targetRow As Long, col As Long, mode As String, profile As String
  On Error Resume Next: Set w = ThisWorkbook.Worksheets("水条件"): On Error GoTo 0
  If w Is Nothing Then Exit Sub
  Set stages = ThisWorkbook.Worksheets("ステージ")
  Set seen = CreateObject("Scripting.Dictionary")
  For r = 12 To 211
    If Application.CountA(w.range(w.cells(r, 12), w.cells(r, 19))) = 0 Then GoTo NextPlan
    id = GWId(w.cells(r, 12).value2, "工程別水理設定の工程番号")
    If seen.Exists(CStr(id)) Then GWFail "工程別水理設定の工程番号が重複しています。"
    seen.Add CStr(id), True: matches = 0: allMatches = 0: count = 0
    For s = 2 To stages.cells(stages.rows.count, 1).End(xlUp).row
      If Len(Trim$(CStr(stages.cells(s, 2).value2))) > 0 Then
        count = count + 1: sid = CLng(Val(CStr(stages.cells(s, 1).value2))): If sid <= 0 Then sid = count
        If sid = id Then
          allMatches = allMatches + 1
          If Val(CStr(stages.cells(s, 5).value2)) <> 0# Then matches = matches + 1: targetRow = s
        End If
      End If
    Next s
    If allMatches <> 1 Then GWFail "工程別水理設定は有効な工程番号へ一意に対応させてください。工程=" & CStr(id)
    If matches = 0 Then GoTo NextPlan
    If Not GWEnabled() Then GWFail "工程別水理設定を使う場合は、水条件のNONEをHYDROSTATICまたはSTEADYへ変更してください。"
    mode = UCase$(Trim$(CStr(w.cells(r, 13).value2)))
    If mode <> "ONE_WAY" And mode <> "FIXED_STRESS" And mode <> "INHERIT" And mode <> "" Then GWFail "連携方式は継承/ONE_WAY/FIXED_STRESSです。工程=" & CStr(id)
    For col = 15 To 19
      If mode = "ONE_WAY" And Len(Trim$(CStr(w.cells(r, col).value2))) > 0 Then GWFail "片方向工程には時間・排水・連成反復を指定できません。工程=" & CStr(id)
    Next col
    If Len(Trim$(CStr(w.cells(r, 14).value2))) > 0 Then
      If UCase$(Trim$(CStr(stages.cells(targetRow, 2).value2))) <> "WATER" Then GWFail "工程別水位線IDはWATER工程に指定してください。"
      If Len(Trim$(CStr(stages.cells(targetRow, 3).value2))) > 0 And Trim$(CStr(stages.cells(targetRow, 3).value2)) <> Trim$(CStr(w.cells(r, 14).value2)) Then GWFail "ステージC列と工程別水理設定の水位線IDが矛盾しています。"
      profile = Trim$(CStr(w.cells(r, 14).value2)): count = 0
      For s = 12 To 211
        If Trim$(CStr(w.cells(s, 1).value2)) = profile Then count = count + 1
      Next s
      If count < 2 Then GWFail "工程別の水位線IDは、水位線表へ2点以上登録してください。"
    End If
NextPlan:
  Next r
End Sub

Public Function GWResolveStageProfile(ByVal stageId As Long, ByVal profile As String) As String
  Dim w As Worksheet, r As Long
  GWResolveStageProfile = profile: Set w = ThisWorkbook.Worksheets("水条件")
  For r = 12 To 211
    If Val(CStr(w.cells(r, 12).value2)) = stageId Then
      If Len(Trim$(CStr(w.cells(r, 14).value2))) > 0 Then GWResolveStageProfile = Trim$(CStr(w.cells(r, 14).value2))
      Exit Function
    End If
  Next r
End Function

Public Sub GWValidateScope()
  Dim ws As Worksheet, r As Long, kind As String
  GJReady = False: mReady = False
  GWActive = GWEnabled()
  GWValidateStagePlans
  If Not GWActive Then Exit Sub
  If P6ReadSetting("CONSOL_ENABLE", 0#) <> 0# Then GWFail "水条件は指定水圧/定常浸透です。過渡圧密と同時には使えません。"
  Set ws = ThisWorkbook.Worksheets("水条件")
  If CStr(ws.range("B7").value2) <> GWInputSignature() Then GWFail "水条件またはモデルが変更されています。解析設定の「水圧を確認」を実行してください。"
  Set ws = ThisWorkbook.Worksheets("ステージ")
  For r = 2 To ws.cells(ws.rows.count, 1).End(xlUp).row
    If val(CStr(ws.cells(r, 5).value2)) <> 0# Then
      kind = UCase$(Trim$(CStr(ws.cells(r, 2).value2)))
      If kind = "RESET_U" Or kind = "RESET_STRESS" Or kind = "MATSET" Then GWFail "水条件とRESET/MATSETの状態変換は未解禁です。工程=" & CStr(r)
    End If
  Next r
End Sub

Private Function GWLevel(ByVal profile As String, ByVal xv As Double) As Double
  Dim start As Double
  If mTemporal And profile = "*" Then
    start = GWRawLevel(mTimeFirst, xv)
    GWLevel = start + mTimeFraction * (GWRawLevel(mTimeLast, xv) - start)
  Else
    GWLevel = GWRawLevel(profile, xv)
  End If
End Function

Private Function GWRawLevel(ByVal profile As String, ByVal xv As Double) As Double
  Dim r As Long, count As Long, xp As Double, yp As Double, xx As Double, yy As Double, got As Boolean
  If profile = "*" Then profile = mProfile
  For r = 12 To 211
    If Trim$(CStr(mWs.cells(r, 1).value2)) = profile Then
      xx = GWNum(mWs.cells(r, 2).value2, "水位X 行" & CStr(r))
      yy = GWNum(mWs.cells(r, 3).value2, "水位標高 行" & CStr(r))
      If count > 0 Then
        If xx <= xp Then GWFail "水位点は線ごとにX昇順です。重複/逆転 行=" & CStr(r)
        If xv >= xp - 0.000000001 And xv <= xx + 0.000000001 Then
          GWRawLevel = yp + (yy - yp) * (xv - xp) / (xx - xp): got = True
        End If
      End If
      xp = xx: yp = yy: count = count + 1
    End If
  Next r
  If count < 2 Or Not got Then GWFail "水位線「" & profile & "」は対象Xを覆う2点以上が必要です。X=" & CStr(xv)
End Function

Private Function GWPressure(ByVal profile As String, ByVal xv As Double, ByVal yv As Double) As Double
  GWPressure = mGamma * (GWLevel(profile, xv) - yv)
  If GWPressure < 0# Then GWPressure = 0#
End Function

Private Function GWEdgeKey(ByVal e As Long, ByVal side As Long) As String
  Dim a As Long, b As Long, c As Long
  a = GJTopoNode(Elem(e).node(side)): b = GJTopoNode(Elem(e).node((side + 1) Mod 4)): c = GJTopoNode(Elem(e).node(side + 4))
  If a < b Then
    GWEdgeKey = CStr(a) & ":" & CStr(b) & ":" & CStr(c)
  Else
    GWEdgeKey = CStr(b) & ":" & CStr(a) & ":" & CStr(c)
  End If
End Function

Private Sub GWLoadDefinition()
  Dim r As Long, mat As Long, e As Long, side As Long, k As String, spec As String
  Dim seen As Object, kind As String, hydroNodes As Object, j As Long, nodeKey As String, pAtNode As Double
  GJReady = False
  Set mWs = ThisWorkbook.Worksheets("水条件")
  mMode = UCase$(Trim$(CStr(mWs.range("B3").value2)))
  If mMode <> "HYDROSTATIC" And mMode <> "STEADY" Then GWFail "水圧方式はNONE/HYDROSTATIC/STEADYです。"
  mGamma = GWNum(mWs.range("B4").value2, "水の単位体積重量")
  If mGamma <= 0# Then GWFail "水の単位体積重量は正数です。"
  mProfile = Trim$(CStr(mWs.range("B5").value2))
  If Len(mOverrideProfile) > 0 Then mProfile = mOverrideProfile
  kind = UCase$(Trim$(CStr(mWs.range("B6").value2)))
  If kind <> "AUTO" And kind <> "OFF" Then GWFail "外水圧はAUTO/OFFです。個別面指定は下の境界表で行います。"
  mAuto = (kind = "AUTO")
  ReDim mBasis(NumberOfMaterial - 1): ReDim mK(NumberOfElement - 1): ReDim mMatProfile(NumberOfMaterial - 1)
  Set seen = CreateObject("Scripting.Dictionary")
  For r = 12 To 211
    If Len(Trim$(CStr(mWs.cells(r, 6).value2))) > 0 Then
      mat = GWId(mWs.cells(r, 6).value2, "水理材料番号") - 1
      If mat < 0 Or mat >= NumberOfMaterial Then GWFail "水理材料番号が範囲外です。"
      If seen.Exists(CStr(mat)) Then GWFail "水理材料の重複です。"
      seen.Add CStr(mat), True
      kind = UCase$(Trim$(CStr(mWs.cells(r, 7).value2)))
      If kind <> "TOTAL" And kind <> "EFFECTIVE" Then GWFail "材料の応力基準をTOTAL/EFFECTIVEで指定してください。"
      If (Material(mat).kind = "STRUCT" Or Material(mat).kind = "JOINT") And kind <> "TOTAL" Then GWFail "構造材料へ地盤の有効応力を適用できません。"
      If Material(mat).kind = "STRUCT" Or Material(mat).kind = "JOINT" Then
        If GWNum(mWs.cells(r, 8).value2, "構造体のK（初期版は遮水0）") <> 0# Then GWFail "構造体の透水は初期版の適用外です。遮水K=0を明示してください。"
      End If
      mBasis(mat) = (kind = "EFFECTIVE")
      mMatProfile(mat) = Trim$(CStr(mWs.cells(r, 9).value2))
      If mMatProfile(mat) = "" Then mMatProfile(mat) = "*"
      For e = 0 To NumberOfElement - 1
        If P3IsElementActive(e) And Elem(e).MatNo = mat And Material(mat).kind = "SOIL" And mMode = "STEADY" Then
          mK(e) = GWNum(mWs.cells(r, 8).value2, "透水係数K[m/s]")
          If mK(e) <= 0# Then GWFail "地盤の透水係数は正数が必要です。"
        End If
      Next e
    End If
  Next r
  GWContext = "GJPrepare"
  GJPrepare
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Not seen.Exists(CStr(Elem(e).MatNo)) Then GWFail "有効材料すべてに水理材料表の応力基準が必要です。材料=" & CStr(Elem(e).MatNo + 1)
  Next e
  If mMode = "HYDROSTATIC" And Not FSActive Then
    Set hydroNodes = CreateObject("Scripting.Dictionary")
    For e = 0 To NumberOfElement - 1
      If P3IsElementActive(e) And Material(Elem(e).MatNo).kind = "SOIL" Then
        For j = 0 To 7
          nodeKey = CStr(GWFlowNode(e, j))
          pAtNode = GWPressure(mMatProfile(Elem(e).MatNo), Elem(e).x(j), Elem(e).y(j))
          If hydroNodes.Exists(nodeKey) Then
            If Abs(CDbl(hydroNodes(nodeKey)) - pAtNode) > 0.0000001 Then GWFail "連続地盤の共有節点に異なる指定水圧があります。水域の遮水/接続を確認し、定常浸透を設定してください。"
          Else
            hydroNodes.Add nodeKey, pAtNode
          End If
        Next j
      End If
    Next e
  End If
  Set mBoundary = CreateObject("Scripting.Dictionary")
  Set mEdges = CreateObject("Scripting.Dictionary"): Set mOwners = CreateObject("Scripting.Dictionary")
  Set mSecondOwners = CreateObject("Scripting.Dictionary")
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Not Elem(e).IsJoint Then
      For side = 0 To 3
        If FEMElementSideValid(e, side) Then
        k = GWEdgeKey(e, side)
        If mEdges.Exists(k) Then
          mEdges(k) = CLng(mEdges(k)) + 1
          If CLng(mEdges(k)) > 2 Then GWFail "境界辺の接続が非多様体です。"
          mSecondOwners.Add k, Array(e, side)
        Else
          mEdges.Add k, 1: mOwners.Add k, Array(e, side)
        End If
        End If
      Next side
    End If
  Next e
  Set mHoles = GWBoundaryLoops(mEdges, mOwners)
  For r = 216 To 415
    If Len(Trim$(CStr(mWs.cells(r, 1).value2))) > 0 Then
      e = GWId(mWs.cells(r, 1).value2, "境界要素") - 1
      side = GWId(mWs.cells(r, 2).value2, "辺番号") - 1
      If e < 0 Or e >= NumberOfElement Or side < 0 Or side > 3 Then GWFail "境界の要素/辺番号が範囲外です。"
      If Not FEMElementSideValid(e, side) Then GWFail "T6の縮退辺には水境界を指定できません。"
      If Elem(e).IsJoint Then GWFail "水境界はJOINTではなく隣接Q8要素に指定してください。"
      spec = CStr(e) & ":" & CStr(side)
      If mBoundary.Exists(spec) Then GWFail "同一面に水条件が重複しています。"
      kind = UCase$(Trim$(CStr(mWs.cells(r, 3).value2)))
      If kind <> "HEAD" And kind <> "NOFLOW" And kind <> "EXPOSED" And kind <> "DRY" Then GWFail "境界種別はHEAD/NOFLOW/EXPOSED/DRYです。"
      mBoundary.Add spec, r
      If Not P3IsElementActive(e) Then GoTo NextWaterBoundary
      If kind = "NOFLOW" And mMode = "STEADY" Then
        If CLng(mEdges(GWEdgeKey(e, side))) = 2 Then GWFail "共有透水節点を持つ内部辺へ遮水を指定できません。遮水体を領域として定義してください。"
      End If
      If kind = "HEAD" And mMode = "STEADY" And Not FSActive Then
        If GWInternalPermeableFace(e, side) Then GWFail "共有する内部透水面へHEADを指定できません。内部の注排水源は未対応です。外周の水頭境界を指定してください。"
      End If
      If mMode = "HYDROSTATIC" And Not FSActive And Material(Elem(e).MatNo).kind = "SOIL" Then
        If kind = "HEAD" Or kind = "EXPOSED" Then
          spec = Trim$(CStr(mWs.cells(r, 4).value2))
          GWCheckFaceHead e, side, spec, mWs.cells(r, 5).value2
        End If
      End If
      If kind = "EXPOSED" Or kind = "DRY" Then
        If CLng(mEdges(GWEdgeKey(e, side))) <> 1 Then GWFail "接触/内部面への外水圧の二重載荷は禁止です。要素=" & CStr(e + 1)
      End If
NextWaterBoundary:
    End If
  Next r
End Sub

Private Function GWInternalPermeableFace(ByVal e As Long, ByVal side As Long) As Boolean
  Dim key As String, other As Variant
  If mK(e) <= 0# Then Exit Function
  key = GWEdgeKey(e, side)
  If CLng(mEdges(key)) <> 2 Then Exit Function
  other = mOwners(key)
  If CLng(other(0)) = e Then other = mSecondOwners(key)
  If mK(CLng(other(0))) <= 0# Then Exit Function
  ' SOIL/PRESCRIBED JOINT faces are physical neighbours but retain separate
  ' hydraulic DOFs. Only an actually shared flow edge is an internal source.
  GWInternalPermeableFace = (GWFlowEdgeKey(e, side) = GWFlowEdgeKey(CLng(other(0)), CLng(other(1))))
End Function

Private Function GWFlowEdgeKey(ByVal e As Long, ByVal side As Long) As String
  Dim a As Long, b As Long, c As Long, swap As Long
  a = GWFlowNode(e, side): b = GWFlowNode(e, (side + 1) Mod 4)
  c = GWFlowNode(e, side + 4)
  If a > b Then swap = a: a = b: b = swap
  GWFlowEdgeKey = CStr(a) & ":" & CStr(b) & ":" & CStr(c)
End Function

' The difference of two piecewise-linear profiles is linear between their
' combined knots. Endpoints plus both knot sets therefore cover the full face.
Private Sub GWCheckFaceHead(ByVal e As Long, ByVal side As Long, ByVal profile As String, ByVal constantHead As Variant)
  Dim a As Double, b As Double, x As Double, r As Long, key As String, soil As String, water As String
  a = Elem(e).x(side): b = Elem(e).x((side + 1) Mod 4)
  If a > b Then x = a: a = b: b = x
  soil = mMatProfile(Elem(e).MatNo): If soil = "*" Or soil = "" Then soil = mProfile
  water = profile: If water = "*" Then water = mProfile
  GWCheckHeadAt e, profile, constantHead, a
  GWCheckHeadAt e, profile, constantHead, b
  For r = 12 To 211
    key = Trim$(CStr(mWs.cells(r, 1).value2))
    If Len(key) > 0 And (key = soil Or key = water) Then
      x = GWNum(mWs.cells(r, 2).value2, "水位X")
      If x > a And x < b Then GWCheckHeadAt e, profile, constantHead, x
    End If
  Next r
End Sub

Private Sub GWCheckHeadAt(ByVal e As Long, ByVal profile As String, ByVal constantHead As Variant, ByVal x As Double)
  Dim h As Double
  If Len(profile) = 0 Then h = GWNum(constantHead, "固定全水頭") Else h = GWLevel(profile, x)
  If Abs(h - GWLevel(mMatProfile(Elem(e).MatNo), x)) > 0.00000001 Then GWFail "外水位と地盤水圧場に水頭差があります。定常浸透へ切り替え、透水係数と境界条件を入力してください。"
End Sub

Private Sub GWApplyHeads(ByRef fixed() As Boolean, ByRef heads() As Double)
  Dim key As Variant, r As Long, e As Long, side As Long, j As Long, n As Long, node As Long
  Dim profile As String, value As Double, localNodes As Variant
  Dim owner As Variant, kind As String, wetCount As Long, useWater As Boolean
  For Each key In mBoundary.keys
    r = CLng(mBoundary(key))
    If UCase$(Trim$(CStr(mWs.cells(r, 3).value2))) = "HEAD" Then
      e = CLng(mWs.cells(r, 1).value2) - 1: side = CLng(mWs.cells(r, 2).value2) - 1
      If Not P3IsElementActive(e) Then GoTo NextHead
      If mK(e) > 0# Then
        profile = Trim$(CStr(mWs.cells(r, 4).value2))
        localNodes = Array(side, (side + 1) Mod 4, side + 4)
        For j = 0 To 2
          n = CLng(localNodes(j)): node = Elem(e).node(n)
          If Len(profile) > 0 Then
            value = GWLevel(profile, Elem(e).x(n))
          Else
            value = GWNum(mWs.cells(r, 5).value2, "固定全水頭")
          End If
          If fixed(node) And Abs(heads(node) - value) > 0.000000001 Then GWFail "共有節点に異なる固定水頭が指定されています。"
          fixed(node) = True: heads(node) = value
        Next j
      Else
        GWFail "HEAD境界は有効な透水地盤要素に指定してください。"
      End If
    End If
NextHead:
  Next key
  ' Ponded exposed soil surfaces are also hydraulic head boundaries.
  For Each key In mEdges.keys
    If CLng(mEdges(key)) = 1 Then
      owner = mOwners(key): e = CLng(owner(0)): side = CLng(owner(1))
      If mK(e) > 0# Then
        profile = "*": useWater = False: kind = ""
        If mBoundary.Exists(CStr(e) & ":" & CStr(side)) Then
          r = CLng(mBoundary(CStr(e) & ":" & CStr(side)))
          kind = UCase$(Trim$(CStr(mWs.cells(r, 3).value2)))
          If kind = "EXPOSED" Then
            profile = Trim$(CStr(mWs.cells(r, 4).value2)): useWater = True
          End If
        ElseIf mAuto And Not CBool(mHoles(key)) Then
          useWater = (Elem(e).x((side + 1) Mod 4) < Elem(e).x(side) - 0.000000001)
        End If
        If useWater Then
          localNodes = Array(side, (side + 1) Mod 4, side + 4): wetCount = 0
          For j = 0 To 2
            n = CLng(localNodes(j))
            If GWLevel(profile, Elem(e).x(n)) >= Elem(e).y(n) - 0.000000001 Then wetCount = wetCount + 1
          Next j
          If wetCount > 0 And wetCount < 3 Then GWFail "部分水没の浸透境界には自由水面処理が必要です。初期の固定飽和領域では未対応です。"
          If wetCount = 3 Then
            For j = 0 To 2
              n = CLng(localNodes(j)): node = Elem(e).node(n): value = GWLevel(profile, Elem(e).x(n))
              If fixed(node) And Abs(heads(node) - value) > 0.000000001 Then GWFail "外水位と固定水頭が競合しています。"
              fixed(node) = True: heads(node) = value
            Next j
          End If
        End If
      End If
    End If
  Next key
End Sub

Public Sub GWRebuild()
  Dim e As Long, gp As Long, j As Long, px As Double, py As Double, head As Double, p As Double, t As Double
  Dim fixed() As Boolean, prescribed() As Double
  GWActive = GWEnabled()
  If Not GWActive Then mReady = False: Exit Sub
  t = Timer: mReady = False: mFlowSolved = False
  mLocalError = 0#: mFluxCorrection = 0#
  GWContext = "GWLoadDefinition"
  GWLoadDefinition
  ReDim mTarget(0 To 3, 0 To NumberOfElement - 1): ReDim mWaterForce(lastDof)
  If mMode = "STEADY" Then
    ReDim fixed(NumberOfFreeNode - 1): ReDim prescribed(NumberOfFreeNode - 1)
    GWApplyHeads fixed, prescribed
    GJApplyHeads fixed, prescribed
    mHeadFixed = fixed
    GWSeepSolve mK, fixed, prescribed, mHeads, mFx, mFy, mFlowReaction, mRel, mBalance
    GJExpandHeads mHeads
    GWRecoverFaceFlows mK, mHeads, fixed, mFaceFlows, mLocalError, mFluxCorrection
    If mRel > 0.0000000001 Or mBalance > 0.00000001 Then GWFail "定常浸透の真残差/全体水収支が許容値を超えています。"
    For e = 0 To NumberOfElement - 1
      If P3IsElementActive(e) And mK(e) > 0# Then
        For j = 0 To 7
          If mGamma * (mHeads(Elem(e).node(j)) - Elem(e).y(j)) < -0.0000001 Then GWFail "負水圧が発生しました。固定飽和領域の適用外です。自由水面/不飽和解析が必要です。"
        Next j
      End If
    Next e
    mFlowSolved = True
  Else
    mRel = 0#: mBalance = 0#
  End If
  P6EnsureGaussCache
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Material(Elem(e).MatNo).kind = "SOIL" Then
      For gp = 0 To 3
        FEMGaussPhysicalXY e, gp, px, py
        If mMode = "STEADY" Then
          head = 0#
          For j = 0 To 7
            head = head + FEMGaussShapeValue(e, gp, j) * mHeads(Elem(e).node(j))
          Next j
          p = mGamma * (head - py)
          If p < -0.0000001 Then GWFail "積分点の負圧は初期の飽和浸透経路では未対応です。"
        Else
          p = GWPressure(mMatProfile(Elem(e).MatNo), px, py)
        End If
        mTarget(gp, e) = p
      Next gp
    End If
  Next e
  GWContext = "GJSetPressures"
  GJSetPressures mTarget, mGamma, (mMode = "STEADY")
  GWContext = "GWBuildSurfaceForce"
  GWBuildSurfaceForce
  GWContext = "GJAddForces"
  GJAddForces mTarget, mWaterForce, (mMode = "STEADY")
  mSig = GWInputSignature(): mSource = mMode & ":" & mProfile & ":" & mSig & ":" & CStr(P3ActiveSetGen)
  mSeconds = Timer - t: If mSeconds < 0# Then mSeconds = mSeconds + 86400#
  mReady = True
End Sub

Private Sub GWBuildSurfaceForce()
  Dim key As Variant, owner As Variant, e As Long, side As Long, r As Long
  Dim profile As String, kind As String, explicit As Boolean, useFace As Boolean
  Dim a As Long, b As Long, count As Long, px() As Double, py() As Double
  Dim j As Long, n As Long, before(0 To 5) As Double, fx As Double, fy As Double, moment As Double, dx As Double, dy As Double
  Set mSurfaceRows = New Collection
  For Each key In mEdges.keys
    If CLng(mEdges(key)) <> 1 Then GoTo NextFace
    owner = mOwners(key): e = CLng(owner(0)): side = CLng(owner(1))
    a = side: b = (side + 1) Mod 4
    kind = "": profile = "*": explicit = False: useFace = False
    If mBoundary.Exists(CStr(e) & ":" & CStr(side)) Then
      r = CLng(mBoundary(CStr(e) & ":" & CStr(side)))
      kind = UCase$(Trim$(CStr(mWs.cells(r, 3).value2)))
      explicit = (kind = "EXPOSED" Or kind = "DRY" Or kind = "HEAD")
      If kind = "EXPOSED" Then
        profile = Trim$(CStr(mWs.cells(r, 4).value2))
        If Len(profile) = 0 Then GWFail "EXPOSEDの水位線IDが必要です。"
        useFace = True
      End If
    End If
    If mAuto And Not explicit Then
      If CBool(mHoles(key)) Then GWFail "空洞内面はEXPOSED（内水位）またはDRY（乾燥）を指定してください。要素=" & CStr(e + 1) & " 辺=" & CStr(side + 1)
      If Material(Elem(e).MatNo).kind = "SOIL" Then
        If Elem(e).x(b) < Elem(e).x(a) - 0.000000001 Then useFace = True
      Else
        GWFail "構造体の露出面はEXPOSED（水位線）またはDRYを個別指定してください。要素=" & CStr(e + 1) & " 辺=" & CStr(side + 1)
      End If
    End If
    If Not useFace Then GoTo NextFace
    GWFaceProfilePoints profile, px, py
    For j = 0 To 2
      Select Case j
        Case 0: n = a
        Case 1: n = b
        Case 2: n = side + 4
      End Select
      before(2 * j) = mWaterForce(2 * Elem(e).node(n))
      before(2 * j + 1) = mWaterForce(2 * Elem(e).node(n) + 1)
    Next j
    GWWetEdgeLoad e, side, px, py, mGamma, mWaterForce
    fx = 0#: fy = 0#: moment = 0#
    For j = 0 To 2
      Select Case j
        Case 0: n = a
        Case 1: n = b
        Case 2: n = side + 4
      End Select
      dx = mWaterForce(2 * Elem(e).node(n)) - before(2 * j)
      dy = mWaterForce(2 * Elem(e).node(n) + 1) - before(2 * j + 1)
      fx = fx + dx: fy = fy + dy: moment = moment + Elem(e).x(n) * dy - Elem(e).y(n) * dx
    Next j
    mSurfaceRows.Add Array(e + 1, side + 1, IIf(CBool(mHoles(key)), "CAVITY", "OUTER"), profile, fx, fy, moment)
NextFace:
  Next key
End Sub

Public Sub GWResetState()
  mTemporal = False
  mOverrideProfile = ""
  GWActive = GWEnabled()
  If Not GWActive Then mReady = False: Exit Sub
  GWContext = "GWRebuild"
  GWRebuild
  ReDim mCurrent(0 To 3, 0 To NumberOfElement - 1): ReDim mCommitted(0 To 3, 0 To NumberOfElement - 1)
  ReDim mStart(0 To 3, 0 To NumberOfElement - 1)
End Sub

Public Sub GWTransientProfile(ByVal first As String, ByVal last As String, ByVal fraction As Double)
  If fraction < 0# Or fraction > 1# Then GWFail "時間水位の補間率が不正です。"
  mTemporal = True: mTimeFirst = first: mTimeLast = last: mTimeFraction = fraction
  mOverrideProfile = last: mProfile = last
  GWLoadDefinition
  ReDim mWaterForce(lastDof)
  GWBuildSurfaceForce
  GJSetPressures mCurrent, mGamma, False
  GJAddForces mCurrent, mWaterForce, True
End Sub

Public Sub GWTransientTopology()
  If Not FSActive Then GWFail "時間連成の水理接続がありません。"
  GWLoadDefinition
  ReDim mWaterForce(lastDof)
  GWBuildSurfaceForce
  mReady = True
End Sub

Public Sub GWTransientFixed(ByRef locked() As Boolean)
  mHeadFixed = locked
End Sub

Public Sub GWBeginLoad()
  If Not GWActive Then Exit Sub
  If Not mReady Then GWFail "水圧場が未受理です。"
  mStart = mCommitted
End Sub

Public Sub GWTrialFactor(ByVal factor As Double)
  Dim e As Long, g As Long
  If Not GWActive Then Exit Sub
  For e = 0 To NumberOfElement - 1
    For g = 0 To 3
      ' A changed pore pressure invalidates material/force evaluation even
      ' when displacement and total external loads are unchanged.
      If mCurrent(g, e) <> mStart(g, e) + factor * (mTarget(g, e) - mStart(g, e)) Then P3TrialStateValid = False
      mCurrent(g, e) = mStart(g, e) + factor * (mTarget(g, e) - mStart(g, e))
    Next g
  Next e
End Sub

Public Sub GWCommit()
  If GWActive And mReady Then mCommitted = mCurrent
End Sub

Public Sub GWRestore()
  If GWActive And mReady Then mCurrent = mCommitted
End Sub

Public Function GWNeedsRamp() As Boolean
  Dim e As Long, g As Long
  If Not GWActive Or Not mReady Then Exit Function
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) Then
      For g = 0 To 3
        If Abs(mTarget(g, e) - mCommitted(g, e)) > 0.000000000001 Then GWNeedsRamp = True: Exit Function
      Next g
    End If
  Next e
End Function

Public Sub GWAddSurfaceForce(ByRef vector() As Double)
  Dim i As Long
  If Not GWActive Or Not mReady Then Exit Sub
  For i = 0 To lastDof
    vector(i) = vector(i) + mWaterForce(i)
  Next i
End Sub

Public Function GWOldPressure(ByVal e As Long, ByVal gp As Long) As Double
  If Not GWActive Or Not mReady Or e < 0 Then Exit Function
  If mBasis(Elem(e).MatNo) Then GWOldPressure = FSBiotAlpha(e) * mCommitted(gp, e)
End Function

Public Function GWNewPressure(ByVal e As Long, ByVal gp As Long) As Double
  If Not GWActive Or Not mReady Or e < 0 Then Exit Function
  If mBasis(Elem(e).MatNo) Then GWNewPressure = FSBiotAlpha(e) * mCurrent(gp, e)
End Function

Public Sub GWUseFinalPressure()
  If Not GWActive Or Not mReady Then Exit Sub
  mCurrent = mTarget: mCommitted = mTarget
End Sub

' Capture accepted physical water together with the mechanical SRM snapshot.
' Never use the last failed trial's target as the accepted pore pressure.
Public Function GWCaptureSrmState() As Object
  Dim saved As Object, rows As Collection, item As Variant
  Set saved = CreateObject("Scripting.Dictionary")
  saved.Add "active", GWActive: saved.Add "ready", mReady
  If GWActive And mReady Then
    saved.Add "selector", GWTransientState()
    saved.Add "target", mTarget: saved.Add "current", mCurrent
    saved.Add "committed", mCommitted: saved.Add "start", mStart
    saved.Add "force", mWaterForce
    saved.Add "conductivity", mK
    saved.Add "solved", mFlowSolved
    If mFlowSolved Then
      saved.Add "heads", mHeads: saved.Add "fx", mFx: saved.Add "fy", mFy
      saved.Add "reaction", mFlowReaction: saved.Add "fixed", mHeadFixed
      If Not FSActive Then saved.Add "faces", mFaceFlows
    End If
    saved.Add "metrics", Array(mRel, mBalance, mSeconds, mLocalError, mFluxCorrection, mInflow, mOutflow)
    saved.Add "signature", mSig: saved.Add "source", mSource
    Set rows = New Collection
    For Each item In mSurfaceRows: rows.Add item: Next item
    saved.Add "surfaces", rows
  End If
  Set GWCaptureSrmState = saved
End Function

Public Sub GWRestoreSrmState(ByVal saved As Object)
  Dim selector As Variant, metrics As Variant, item As Variant, rows As Collection
  If saved Is Nothing Then GWFail "SRM成功状態の水圧スナップショットがありません。"
  GWActive = CBool(saved("active")): mReady = False: GJReady = False
  If Not GWActive Then Exit Sub
  If Not CBool(saved("ready")) Then GWFail "SRM成功状態の水圧場が未準備です。"
  selector = saved("selector")
  mTemporal = selector(0): mTimeFirst = selector(1): mTimeLast = selector(2): mTimeFraction = selector(3)
  mOverrideProfile = selector(4)
  ' Active elements and node conditions have already been restored by SRM.
  ' Rebuild definitions and JOINT topology/flow aliases for that same active set.
  GWLoadDefinition
  ' Transient conductivity can differ from the one-way input table.
  mK = saved("conductivity")
  mTarget = saved("target"): mCurrent = saved("current")
  mCommitted = saved("committed"): mStart = saved("start")
  mWaterForce = saved("force"): mHeadFixed = selector(6)
  mFlowSolved = CBool(saved("solved"))
  Erase mHeads: Erase mFx: Erase mFy: Erase mFlowReaction: Erase mFaceFlows
  If mFlowSolved Then
    mHeads = saved("heads"): mFx = saved("fx"): mFy = saved("fy")
    mFlowReaction = saved("reaction"): mHeadFixed = saved("fixed")
    If saved.Exists("faces") Then mFaceFlows = saved("faces")
  End If
  metrics = saved("metrics")
  mRel = metrics(0): mBalance = metrics(1): mSeconds = metrics(2)
  mLocalError = metrics(3): mFluxCorrection = metrics(4): mInflow = metrics(5): mOutflow = metrics(6)
  mSig = saved("signature"): mSource = saved("source")
  Set rows = saved("surfaces"): Set mSurfaceRows = New Collection
  For Each item In rows: mSurfaceRows.Add item: Next item
  mReady = True
End Sub

Public Sub GWStage(ByVal profile As String)
  If Not GWEnabled() Then GWFail "WATER工程には水条件の有効化が必要です。"
  ' Stage profile is a separate selector; do not rewrite the input sheet.
  mOverrideProfile = profile
  GWContext = "GWRebuild"
  GWRebuild
End Sub

Public Sub GWWriteResults(ByVal stageNo As Long)
  Dim ws As Worksheet, e As Long, g As Long, row As Long, px As Double, py As Double, p As Double
  If Not GWActive Or Not mReady Then Exit Sub
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("水圧結果")
  On Error GoTo 0
  If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = "水圧結果"
  If stageNo = 0 Then ws.cells.ClearContents
  ws.range("A35:O35").value2 = Array("工程", "要素", "GP", "X[m]", "Y[m]", "水圧[kPa]", "全σX", "有効σX", "全σY", "有効σY", "全σZ", "有効σZ", "qX[m/s]", "qY[m/s]", "水圧元")
  row = ws.cells(ws.rows.count, 1).End(xlUp).row + 1
  If row < 36 Then row = 36
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Not Elem(e).IsJoint Then
      For g = 0 To 3
        FEMGaussPhysicalXY e, g, px, py: p = mCurrent(g, e)
        ws.cells(row, 1).Resize(1, 12).value2 = Array(stageNo, e + 1, g + 1, px, py, p, Elem(e).Stmat(0, g), Elem(e).Stmat(0, g) - FSBiotAlpha(e) * p, Elem(e).Stmat(1, g), Elem(e).Stmat(1, g) - FSBiotAlpha(e) * p, Elem(e).Stmat(16, g), Elem(e).Stmat(16, g) - FSBiotAlpha(e) * p)
        If mFlowSolved And mK(e) > 0# Then ws.cells(row, 13).value2 = mFx(g, e): ws.cells(row, 14).value2 = mFy(g, e)
        ws.cells(row, 15).value2 = mSource
        If FSActive Then
          ws.cells(row, 27).Resize(1, 3).value2 = Array(FSTime, FSReferencePressure(e, g), p - FSReferencePressure(e, g))
        End If
        row = row + 1
      Next g
    End If
  Next e
  If FSActive Then ws.range("AA35:AC35").value2 = Array("累積時刻[s]", "固定基準水圧[kPa]", "過剰水圧[kPa]")
  ws.range("Q1:R5").ClearContents
  ws.range("Q1").value2 = "水理方式": ws.range("R1").value2 = GWViewMode()
  ws.range("Q2").value2 = "真相対残差": ws.range("R2").value2 = mRel
  ws.range("Q3").value2 = "全体水収支": ws.range("R3").value2 = mBalance
  ws.range("Q4").value2 = "水理計算秒": ws.range("R4").value2 = mSeconds
  ws.range("Q5").value2 = "入力署名": ws.range("R5").value2 = mSig
  ws.columns("A:N").ColumnWidth = 12
  ws.columns("O").ColumnWidth = 24
  ws.range("A1").value2 = "水圧・全応力/有効応力"
  ws.range("A2").value2 = "方式": ws.range("B2").value2 = GWViewMode()
  ws.range("D2").value2 = "真相対残差": ws.range("E2").value2 = mRel
  ws.range("G2").value2 = "水収支": ws.range("H2").value2 = mBalance
  ws.range("A4").value2 = "図の下（35行目）に全積分点の数値を表示。点の色は地盤間隙水圧、矢印は等価節点荷重。"
  GWWriteHydraulicTables ws, stageNo
  GWDrawWaterView ws, mCurrent, mWaterForce, "受理水圧 工程=" & CStr(stageNo) & " " & GWViewMode()
End Sub

Public Sub GWPreviewCompute()
  Dim ws As Worksheet, key As Variant, owner As Variant, row As Long, e As Long, s As Long, g As Long
  Dim px As Double, py As Double, normalX As Double, normalY As Double
  If Not GWEnabled() Then GWFail "水条件はNONEです。水圧を確認する場合はHYDROSTATICまたはSTEADYを選んでください。"
  GWValidateStagePlans
  FEMInvalidateInputCache
  GJReady = False: mReady = False
  Call SetNodeData
  Call SetMaterialData
  Call SetElementData
  P3ElementActiveReady = False
  If Not SetElmMat() Then GWFail AnalysisMessage
  GWContext = "P3InitElementActive"
  P3InitElementActive
  GWContext = "GWRebuild"
  GWRebuild
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("水圧確認")
  On Error GoTo 0
  If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = "水圧確認"
  ws.cells.ClearContents
  ws.range("A35:H35").value2 = Array("要素", "GP", "X[m]", "Y[m]", "水圧[kPa]", "方式", "真相対残差", "全体水収支")
  row = 36
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Not Elem(e).IsJoint Then
      For g = 0 To 3
        FEMGaussPhysicalXY e, g, px, py
        ws.cells(row, 1).Resize(1, 8).value2 = Array(e + 1, g + 1, px, py, mTarget(g, e), mMode, mRel, mBalance): row = row + 1
      Next g
    End If
  Next e
  ws.range("Q35:U35").value2 = Array("水圧外力DOF", "Fx[kN]", "Fy[kN]", "節点X", "節点Y")
  For e = 0 To NumberOfElement - 1
    For g = 0 To IIf(Elem(e).IsJoint, 5, 7)
      row = Elem(e).node(g) + 36
      ws.cells(row, 17).Resize(1, 5).value2 = Array(Elem(e).node(g) + 1, mWaterForce(2 * Elem(e).node(g)), mWaterForce(2 * Elem(e).node(g) + 1), Elem(e).x(g), Elem(e).y(g))
    Next g
  Next e
  ws.columns("A:N").ColumnWidth = 12
  ws.range("A1").value2 = "解析前の水圧確認"
  ws.range("A2").value2 = "方式": ws.range("B2").value2 = mMode
  ws.range("D2").value2 = "真相対残差": ws.range("E2").value2 = mRel
  ws.range("G2").value2 = "水収支": ws.range("H2").value2 = mBalance
  ws.range("A4").value2 = "水位線・GP水圧・等価節点水圧荷重を確認。全数値は35行目以降。"
  GWContext = "GWWriteHydraulicTables"
  GWWriteHydraulicTables ws, 0
  GWContext = "GWDrawWaterView"
  GWDrawWaterView ws, mTarget, mWaterForce, "解析前水圧 " & mMode
  ws.Activate
End Sub

Public Sub GWResetElement(ByVal elementId As Long)
  Dim g As Long
  If Not GWActive Or Not mReady Then Exit Sub
  For g = 0 To 3
    mCommitted(g, elementId) = 0#: mCurrent(g, elementId) = 0#
  Next g
End Sub

Public Function GWCharacteristicForce() As Double
  Dim e As Long, g As Long, j As Long, value As Double
  If Not GWActive Or Not mReady Then Exit Function
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And mBasis(Elem(e).MatNo) Then
      For g = 0 To 3
        For j = 0 To 15
          value = value + Abs(mTarget(g, e) * (Elem(e).bMat(0, j, g) + Elem(e).bMat(1, j, g)) * Elem(e).dj(g) * Material(Elem(e).MatNo).thickness)
        Next j
      Next g
    End If
  Next e
  GWCharacteristicForce = value
End Function

Public Function GWWriteCheckpoint(ByVal folder As String) As Boolean
  Dim f As Integer, e As Long, g As Long, decimalMark As String
  GWWriteCheckpoint = False
  If Not GWActive Then GWWriteCheckpoint = True: Exit Function
  On Error GoTo Failed
  f = FreeFile
  Open folder & Application.PathSeparator & "water.csv" For Output As #f
  Print #f, "WATER_V2," & GWInputSignature()
  Print #f, "PROFILE," & mProfile
  Print #f, "FOS," & GWNumberText(P3FosPass) & "," & GWNumberText(P3FosFail) & "," & GWNumberText(P3FosMid) & "," & GWNumberText(P3FosWidth) & "," & IIf(P3FosBracket, "1", "0") & "," & P3FosInterpretation & "," & CStr(P3SrmTrialCount)
  Print #f, "element,gauss,pressure_kPa"
  decimalMark = Application.International(xlDecimalSeparator)
  For e = 0 To NumberOfElement - 1
    For g = 0 To 3
      Print #f, CStr(e + 1) & "," & CStr(g + 1) & "," & Replace$(Format$(mCommitted(g, e), "0.0000000000000000E+00"), decimalMark, ".")
    Next g
  Next e
  Close #f
  GWWriteCheckpoint = True
  Exit Function
Failed:
  On Error Resume Next
  Close #f
End Function

Public Sub GWReadCheckpoint(ByVal folder As String)
  Dim f As Integer, line As String, parts() As String, count As Long, e As Long, g As Long
  Dim loaded() As Double, seen() As Boolean, p As Double, profile As String
  Dim errorNumber As Long, errorText As String, resultSheet As Worksheet, resultRow As Long
  If Not GWActive Then Exit Sub
  On Error GoTo Failed
  f = FreeFile
  Open folder & Application.PathSeparator & "water.csv" For Input As #f
  Line Input #f, line
  If line <> "WATER_V2," & GWInputSignature() Then GWFail "水圧チェックポイントの版/入力署名が一致しません。"
  Line Input #f, line
  If left$(line, 8) <> "PROFILE," Then GWFail "水圧チェックポイントの水位線がありません。"
  profile = mid$(line, 9)
  Line Input #f, line: parts = Split(line, ",")
  If UBound(parts) <> 7 Then GWFail "水圧チェックポイントのFs情報が不完全です。"
  If parts(0) <> "FOS" Then GWFail "水圧チェックポイントのFs情報がありません。"
  P3FosPass = GWParseNumber(parts(1)): P3FosFail = GWParseNumber(parts(2))
  P3FosMid = GWParseNumber(parts(3)): P3FosWidth = GWParseNumber(parts(4))
  If parts(5) <> "0" And parts(5) <> "1" Then GWFail "Fs区間フラグが不正です。"
  P3FosBracket = (parts(5) = "1"): P3FosInterpretation = parts(6)
  P3SrmTrialCount = CLng(GWParseNumber(parts(7)))
  If P3FosPass < 0# Or P3FosFail < 0# Or P3FosWidth < 0# Or P3SrmTrialCount < 0 Then GWFail "Fs情報が不正です。"
  If P3FosBracket Then
    If P3FosFail <= P3FosPass Or Abs(P3FosWidth - (P3FosFail - P3FosPass)) > 0.000000001 Then GWFail "Fs区間の整合性がありません。"
  End If
  Line Input #f, line
  If line <> "element,gauss,pressure_kPa" Then GWFail "水圧チェックポイントの列が不正です。"
  ReDim loaded(0 To 3, 0 To NumberOfElement - 1): ReDim seen(0 To 3, 0 To NumberOfElement - 1)
  Do While Not EOF(f)
    Line Input #f, line: parts = Split(line, ",")
    If UBound(parts) <> 2 Then GWFail "水圧チェックポイントが不完全です。"
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then GWFail "水圧チェックポイントに非数値があります。"
    e = CLng(parts(0)) - 1: g = CLng(parts(1)) - 1
    If e < 0 Or e >= NumberOfElement Or g < 0 Or g > 3 Then GWFail "水圧チェックポイントのIDが範囲外です。"
    If seen(g, e) Then GWFail "水圧チェックポイントに重複があります。"
    p = CDbl(Replace$(parts(2), ".", Application.International(xlDecimalSeparator)))
    If Not P2IsFinite(p) Then GWFail "水圧チェックポイントが非有限値です。"
    loaded(g, e) = p: seen(g, e) = True: count = count + 1
  Loop
  Close #f: f = 0
  If count <> 4 * NumberOfElement Then GWFail "水圧チェックポイントの行が不足しています。"
  If P3SrmTrialCount > 0 Then
    P3SrmNote = "保存水圧・Fs照査を復元。判定=" & P3FosInterpretation & " Fs下限=" & CStr(P3FosPass)
    Set resultSheet = ThisWorkbook.Worksheets("ステージ結果")
    For resultRow = resultSheet.cells(resultSheet.rows.count, 1).End(xlUp).row To 2 Step -1
      If UCase$(CStr(resultSheet.cells(resultRow, 2).value2)) = "SRM" Then
        resultSheet.cells(resultRow, 5).value2 = IIf(P3FosBracket, P3FosMid, P3FosPass)
        resultSheet.cells(resultRow, 9).value2 = P3SrmTrialCount
        resultSheet.cells(resultRow, 10).value2 = P3SrmNote
        Exit For
      End If
    Next resultRow
  End If
  mOverrideProfile = profile
  GWContext = "GWRebuild"
  If Dir$(folder & Application.PathSeparator & "fixed_stress.csv") <> "" Then
    FSActive = True
    GWLoadDefinition
    mReady = True: mFlowSolved = False
    ReDim mWaterForce(lastDof)
    mCurrent = loaded: mCommitted = loaded: mTarget = loaded
    GWBuildSurfaceForce
    FSReadCheckpoint folder
    For e = 0 To NumberOfElement - 1
      For g = 0 To 3
        If Abs(loaded(g, e) - mCurrent(g, e)) > 0.0000000001 * (1# + Abs(loaded(g, e))) Then GWFail "時間連成の節点水圧と積分点水圧が一致しません。"
      Next g
    Next e
    Exit Sub
  End If
  GWRebuild
  For e = 0 To NumberOfElement - 1
    For g = 0 To 3
      If Abs(loaded(g, e) - mTarget(g, e)) > 0.0000000001 * (1# + Abs(mTarget(g, e))) Then GWFail "再生成した水圧場とチェックポイントが一致しません。"
    Next g
  Next e
  mCurrent = loaded: mCommitted = loaded: mStart = loaded
  Exit Sub
Failed:
  errorNumber = Err.number: errorText = Err.description
  On Error Resume Next
  If f > 0 Then Close #f
  On Error GoTo 0
  Err.Raise errorNumber, "GWReadCheckpoint", errorText
End Sub

Public Sub GWClearResults()
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("水圧結果")
  On Error GoTo 0
  If Not ws Is Nothing Then
    ws.cells.ClearContents
    GWClearWaterView ws
    ws.range("A1").value2 = "水圧結果なし：解析が正常終了してから表示してください。"
  End If
End Sub

Private Sub GWWriteHydraulicTables(ByVal ws As Worksheet, ByVal stageNo As Long)
  Dim row As Long, e As Long, side As Long, n As Long, j As Long, item As Variant
  Dim nodeX() As Double, nodeY() As Double
  ws.columns("Q").ColumnWidth = 24: ws.columns("R").ColumnWidth = 18
  ws.range("R5").NumberFormat = "0"
  ws.range("Q6").value2 = "要素水収支誤差": ws.range("R6").value2 = mLocalError
  ws.range("Q7").value2 = "保存流量の補正率": ws.range("R7").value2 = mFluxCorrection
  ws.range("W35:AD35").value2 = Array("工程", "要素", "辺(CCW)", "外面/空洞", "水位線", "Fx[kN]", "Fy[kN]", "M原点[kNm]")
  row = ws.cells(ws.rows.count, 23).End(xlUp).row + 1: If row < 36 Then row = 36
  For Each item In mSurfaceRows
    ws.cells(row, 23).Resize(1, 8).value2 = Array(stageNo, item(0), item(1), item(2), item(3), item(4), item(5), item(6))
    row = row + 1
  Next item
  If Not mFlowSolved Then Exit Sub
  ReDim nodeX(NumberOfFreeNode - 1): ReDim nodeY(NumberOfFreeNode - 1)
  For e = 0 To NumberOfElement - 1
    For j = 0 To IIf(Elem(e).IsJoint, 5, 7)
      nodeX(Elem(e).node(j)) = Elem(e).x(j): nodeY(Elem(e).node(j)) = Elem(e).y(j)
    Next j
  Next e
  ws.range("AF35:AM35").value2 = Array("工程", "水理DOF", "全水頭[m]", "固定水頭", "供給流量[m3/s]", "流量種別", "X[m]", "Y[m]")
  row = ws.cells(ws.rows.count, 32).End(xlUp).row + 1: If row < 36 Then row = 36
  For n = 0 To NumberOfFreeNode - 1
    ws.cells(row, 32).Resize(1, 8).value2 = Array(stageNo, n + 1, mHeads(n), mHeadFixed(n), mFlowReaction(n), "REACTION_INWARD", nodeX(n), nodeY(n))
    row = row + 1
  Next n
  If FSActive Then Exit Sub
  ws.range("AO35:AS35").value2 = Array("工程", "要素", "辺(CCW)", "外向流量[m3/s]", "流量種別")
  row = ws.cells(ws.rows.count, 41).End(xlUp).row + 1: If row < 36 Then row = 36
  For e = 0 To NumberOfElement - 1
    If mK(e) > 0# Then
      For side = 0 To 3
        If FEMElementSideValid(e, side) Then
        ws.cells(row, 41).Resize(1, 5).value2 = Array(stageNo, e + 1, side + 1, mFaceFlows(side, e), "EQUILIBRATED")
        row = row + 1
        End If
      Next side
    End If
  Next e
End Sub

' Rebuild force from accepted stresses; never re-run the constitutive return on resume.
Public Sub GWRebuildAcceptedInternalForce()
  Dim i As Long, e As Long
  For i = 0 To lastDof
    iNForce(i) = 0#
  Next i
  For e = 0 To NumberOfElement - 1
    ' Strain output is derived from the restored displacement; gauss.csv stores stresses only.
    If Not Elem(e).IsJoint Then CalcStrain e, Elem(e).bMat, TDisp, Elem(e).Stmat
    If P3IsElementActive(e) Then
      If Elem(e).IsJoint Then
        P3JointEval e, True
      Else
        CalcForce Elem(e).ElNo, Elem(e).bMat, Elem(e).dj, Elem(e).Stmat
      End If
    End If
  Next e
  If P6MixedUP Then P6AddPressureToInternalForce
  FSForceCorrection iNForce
  For i = 0 To lastDof
    P3FinalInternalForce(i) = iNForce(i)
  Next i
End Sub

Private Function GWNumberText(ByVal number As Double) As String
  GWNumberText = Replace$(Format$(number, "0.0000000000000000E+00"), Application.International(xlDecimalSeparator), ".")
End Function

Private Function GWParseNumber(ByVal text As String) As Double
  GWParseNumber = GWNum(Replace$(text, ".", Application.International(xlDecimalSeparator)), "チェックポイント数値")
End Function

Public Function GWProfileLevel(ByVal profile As String, ByVal x As Double) As Double
  GWProfileLevel = GWLevel(profile, x)
End Function

Public Function GWMaterialProfile(ByVal mat As Long) As String
  GWMaterialProfile = mMatProfile(mat)
End Function

' Read-only drawing access. A water-level profile is an input, never a solved free surface.
Public Function GWViewMode() As String
  If FSActive Then GWViewMode = "FIXED_STRESS": Exit Function
  GWViewMode = mMode
End Function

Public Function GWViewDefaultProfile() As String
  GWViewDefaultProfile = mProfile
End Function

Public Function GWViewHead(ByVal e As Long, ByVal localNode As Long, ByRef head As Double, ByRef prescribed As Boolean) As Boolean
  Dim n As Long
  If Not mReady Or Not mFlowSolved Then Exit Function
  If e < 0 Or e >= NumberOfElement Or localNode < 0 Or localNode > 7 Then Exit Function
  If Not P3IsElementActive(e) Or Elem(e).IsJoint Then Exit Function
  If mK(e) <= 0# Then Exit Function
  n = Elem(e).node(localNode): head = mHeads(n)
  prescribed = mHeadFixed(GWFlowNode(e, localNode))
  GWViewHead = True
End Function

Public Function GWViewFlux(ByVal e As Long, ByVal gp As Long, ByRef qx As Double, ByRef qy As Double) As Boolean
  If Not mReady Or Not mFlowSolved Then Exit Function
  If e < 0 Or e >= NumberOfElement Or gp < 0 Or gp > 3 Then Exit Function
  If Not P3IsElementActive(e) Or Elem(e).IsJoint Then Exit Function
  If mK(e) <= 0# Then Exit Function
  qx = mFx(gp, e): qy = mFy(gp, e): GWViewFlux = True
End Function

Public Function GWJointPressure(ByVal e As Long, ByVal gp As Long) As Double
  If Not GWActive Or Not mReady Then Exit Function
  If Elem(e).IsJoint And P3IsElementActive(e) Then GWJointPressure = mCurrent(gp, e)
End Function

Public Function GWFacePressure(ByVal e As Long, ByVal side As Long, ByVal xi As Double, ByVal startNode As Long) As Double
  Dim a As Long, n As Long, ids As Variant, nn(0 To 2) As Double, x As Double, y As Double, h As Double
  If FSActive Then GWFacePressure = FSFacePressure(e, side, xi, startNode): Exit Function
  ' Opposite owners have duplicated IDs, so compare the hydraulic/topological alias.
  If GJTopoNode(Elem(e).node(side)) <> GJTopoNode(startNode) Then xi = -xi
  ids = Array(side, (side + 1) Mod 4, side + 4)
  nn(0) = 0.5 * xi * (xi - 1#): nn(1) = 0.5 * xi * (xi + 1#): nn(2) = 1# - xi * xi
  For a = 0 To 2
    n = CLng(ids(a)): x = x + nn(a) * Elem(e).x(n): y = y + nn(a) * Elem(e).y(n)
    If mMode = "STEADY" Then h = h + nn(a) * mHeads(Elem(e).node(n))
  Next a
  If mMode <> "STEADY" Then h = GWLevel(mMatProfile(Elem(e).MatNo), x)
  GWFacePressure = mGamma * (h - y)
  If mMode = "STEADY" And GWFacePressure < -0.0000001 Then GWFail "JOINT面の負水圧は固定飽和領域の適用外です。"
  If GWFacePressure < 0# Then GWFacePressure = 0#
End Function

Private Sub GWFaceProfilePoints(ByVal profile As String, ByRef px() As Double, ByRef py() As Double)
  Dim r As Long, i As Long, j As Long, count As Long, firstCount As Long, lastCount As Long
  Dim id As String, x As Double, swap As Double, lower As Double, upper As Double
  Dim firstMin As Double, firstMax As Double, lastMin As Double, lastMax As Double
  If mTemporal And profile = "*" Then
    For r = 12 To 211
      id = Trim$(CStr(mWs.cells(r, 1).value2))
      If id = mTimeFirst Or id = mTimeLast Then
        x = GWNum(mWs.cells(r, 2).value2, "時間水位X")
        If id = mTimeFirst Then
          If firstCount = 0 Then firstMin = x
          firstMax = x: firstCount = firstCount + 1
        End If
        If id = mTimeLast Then
          If lastCount = 0 Then lastMin = x
          lastMax = x: lastCount = lastCount + 1
        End If
      End If
    Next r
    If firstCount < 2 Or lastCount < 2 Then GWFail "時間水位には開始・終了とも2点以上が必要です。"
    lower = firstMin: If lastMin > lower Then lower = lastMin
    upper = firstMax: If lastMax < upper Then upper = lastMax
    If upper <= lower Then GWFail "時間水位線のX範囲が重なりません。"
    ReDim px(0 To 201): px(0) = lower: px(1) = upper: count = 2
    For r = 12 To 211
      id = Trim$(CStr(mWs.cells(r, 1).value2))
      If id = mTimeFirst Or id = mTimeLast Then
        x = GWNum(mWs.cells(r, 2).value2, "時間水位X")
        If x > lower And x < upper Then px(count) = x: count = count + 1
      End If
    Next r
    For i = 1 To count - 1
      x = px(i): j = i - 1
      Do While j >= 0
        If px(j) <= x Then Exit Do
        px(j + 1) = px(j): j = j - 1
      Loop
      px(j + 1) = x
    Next i
    j = 0
    For i = 1 To count - 1
      If px(i) > px(j) Then j = j + 1: px(j) = px(i)
    Next i
    count = j + 1: ReDim Preserve px(count - 1): ReDim py(count - 1)
    For i = 0 To count - 1: py(i) = GWLevel("*", px(i)): Next i
  Else
    If profile = "*" Then profile = mProfile
    For r = 12 To 211
      If Trim$(CStr(mWs.cells(r, 1).value2)) = profile Then count = count + 1
    Next r
    If count < 2 Then GWFail "面の水位線には2点以上が必要です。"
    ReDim px(count - 1): ReDim py(count - 1): count = 0
    For r = 12 To 211
      If Trim$(CStr(mWs.cells(r, 1).value2)) = profile Then
        px(count) = GWNum(mWs.cells(r, 2).value2, "水位X")
        py(count) = GWNum(mWs.cells(r, 3).value2, "水位Y"): count = count + 1
      End If
    Next r
  End If
End Sub

Public Sub GWAddProfileFaceLoad(ByVal e As Long, ByVal side As Long, ByVal profile As String, ByRef force() As Double)
  Dim px() As Double, py() As Double
  GWFaceProfilePoints profile, px, py
  GWWetEdgeLoad e, side, px, py, mGamma, force
End Sub

' A new mechanical mesh invalidates all node alias maps and cached water state.
Public Sub GWInvalidateTopology()
  GWActive = False: GJReady = False: mReady = False: mFlowSolved = False
  FSReset
End Sub

Public Function GWMaterialEffective(ByVal mat As Long) As Boolean
  If mReady Then GWMaterialEffective = mBasis(mat)
End Function

Public Sub GWTransientConductivity(ByRef kx() As Double)
  Dim e As Long
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Material(Elem(e).MatNo).kind = "SOIL" Then mK(e) = kx(Elem(e).MatNo)
  Next e
End Sub

Public Function GWInitialNodePressure(ByVal e As Long, ByVal slot As Long) As Double
  If Not mReady Then GWFail "時間連成の初期水圧場がありません。"
  If mFlowSolved And Not FSActive Then
    GWInitialNodePressure = mGamma * (mHeads(Elem(e).node(slot)) - Elem(e).y(slot))
  Else
    GWInitialNodePressure = mGamma * (GWLevel(mMatProfile(Elem(e).MatNo), Elem(e).x(slot)) - Elem(e).y(slot))
  End If
End Function

Public Sub GWTransientBoundary(ByRef locked() As Boolean, ByRef pressure() As Double)
  Dim n As Long
  ReDim locked(NumberOfFreeNode - 1): ReDim pressure(NumberOfFreeNode - 1)
  GWApplyHeads locked, pressure
  GJApplyHeads locked, pressure
  mHeadFixed = locked
  For n = 0 To NumberOfFreeNode - 1
    If locked(n) Then pressure(n) = mGamma * (pressure(n) - XXX(2 * n + 1))
  Next n
End Sub

Public Sub GWSetTransientPressure(ByRef field() As Double)
  If Not mReady Then GWFail "時間連成の水圧場が未準備です。"
  mCurrent = field: mTarget = field
  FSViewHydraulics mHeads, mFx, mFy, mFlowReaction
  mFlowSolved = True
  ' JOINT pressure is recovered from adjacent physical soil pressure.
  GJSetPressures mCurrent, mGamma, False
  ReDim mWaterForce(lastDof)
  GWBuildSurfaceForce
  GJAddForces mCurrent, mWaterForce, True
  mTarget = mCurrent
  mSource = "FIXED_STRESS:t=" & CStr(FSTime) & ":" & mSig
  mRel = FSFluidResidual: mBalance = FSMassError
End Sub

Public Sub GWTransientForce(ByRef force() As Double)
  force = mWaterForce
End Sub

Public Function GWTransientState() As Variant
  Dim fixed() As Boolean
  ReDim fixed(NumberOfFreeNode - 1)
  On Error Resume Next: fixed = mHeadFixed: On Error GoTo 0
  GWTransientState = Array(mTemporal, mTimeFirst, mTimeLast, mTimeFraction, mOverrideProfile, mProfile, fixed)
End Function
Public Sub GWRestoreTransientState(ByVal saved As Variant)
  mTemporal = saved(0): mTimeFirst = saved(1): mTimeLast = saved(2): mTimeFraction = saved(3)
  mOverrideProfile = saved(4): mProfile = saved(5): mHeadFixed = saved(6)
  GWLoadDefinition
End Sub
Public Sub GWTransientJointConstraints(ByRef locked() As Boolean, ByRef pressure() As Double)
  Dim fixedJoint() As Boolean, headJoint() As Double, n As Long, value As Double
  ReDim fixedJoint(NumberOfFreeNode - 1): ReDim headJoint(NumberOfFreeNode - 1)
  GJApplyHeads fixedJoint, headJoint
  For n = 0 To NumberOfFreeNode - 1
    If fixedJoint(n) Then
      value = mGamma * (headJoint(n) - XXX(2 * n + 1))
      If locked(n) And Abs(pressure(n) - value) > 0.0000001 Then GWFail "排水セットとJOINT指定水位の境界が矛盾しています。"
      locked(n) = True: pressure(n) = value
    End If
  Next n
End Sub
