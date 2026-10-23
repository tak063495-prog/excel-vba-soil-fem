Option Explicit

' Independent of the dormant global CONSOL_* / MIXED_UP implementation.
Public FSActive As Boolean, FSTime As Double, FSAccepted As Long
Public FSFluidResidual As Double, FSMassError As Double, FSOuterIteration As Long
Private mMap() As Long, mMass() As Double, mFlow() As Double, mStab() As Double, mQ() As Double, mGravity() As Double
Private mP() As Double, mOldP() As Double, mReference() As Double, mUsed() As Boolean
Private mPorosity() As Double
Private mAlpha() As Double, mStorage() As Double, mKx() As Double, mKy() As Double
Private mGp() As Double, mOldGp() As Double, mReferenceGp() As Double, mGamma As Double
Private mGeneration As Long, mReady As Boolean
Private mBoundaryFlow() As Double
Private mSource() As Double, mDrainSet As String
Private mPropertiesReady As Boolean
Public FSRunningStage As Long, FSStageElapsed As Double, FSStageDuration As Double, FSResumePending As Boolean
Private mStageBaseSelf() As Double, mStageBaseApplied() As Double, mStageBaseBoundary() As Double
Private mStageEndSelf() As Double, mStageEndApplied() As Double, mStageEndBoundary() As Double
Private mStageStartProfile As String, mStageDt As Double
Private mStageWaterBefore() As Double

Public Sub FSInstallInput()
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets("水条件")
  ws.range("V11:AA11").value2 = Array("材料番号", "Biot係数α", "貯留1/M[1/kPa]", "kx[m/s]", "ky[m/s]", "備考")
  ws.range("O11:P11").value2 = Array("工程時間（詳細の単位）", "初期刻み（詳細の単位）")
  ws.range("AB11").value2 = "間隙率n（施工の水量収支）"
  ws.range("AC11:AL11").value2 = Array("工程ID", "時間単位", "水位変化", "載荷方法", "最小刻み", "最大刻み", "保存間隔", "水位時系列ID", "載荷セットID", "備考")
  ws.range("V215:AB215").value2 = Array("排水セットID", "要素番号", "辺番号", "種別", "水位線ID", "指定値", "備考")
  ws.range("V419:X419").value2 = Array("時系列ID", "工程内時刻", "水位線ID")
  ThisWorkbook.Worksheets("載荷").cells(1, 11).value2 = "載荷セットID（空欄は従来）"
  ws.range("V11:AL11").WrapText = True: ws.range("V11:AL11").Interior.Color = RGB(218, 232, 246)
  ws.range("V215:AB215").WrapText = True: ws.range("V215:AB215").Interior.Color = RGB(218, 232, 246)
End Sub

Public Sub FSCoupledInputError(ByVal message As String)
  SetAnalysisFailure RESULT_INPUT_ERROR, message, vbObjectError + 8120, -1, -1, CurrentIncrement, CurrentIteration
  Err.Raise vbObjectError + 8120, "FEMFixedStress", message
End Sub

Public Function FSPlanRow(ByVal id As Long) As Long
  Dim r As Long, ws As Worksheet
  Set ws = ThisWorkbook.Worksheets("水条件")
  For r = 12 To 211
    If Val(CStr(ws.cells(r, 12).value2)) = id Then FSPlanRow = r: Exit Function
  Next r
End Function

Public Function FSRequested(ByVal id As Long) As Boolean
  Dim r As Long, mode As String
  r = FSPlanRow(id)
  If r > 0 Then mode = UCase$(Trim$(CStr(ThisWorkbook.Worksheets("水条件").cells(r, 13).value2)))
  If mode = "FIXED_STRESS" Then FSRequested = True: Exit Function
  If mode = "" Or mode = "INHERIT" Then FSRequested = FSActive
End Function

Public Function FSResumingStage(ByVal id As Long) As Boolean
  FSResumingStage = FSResumePending And FSRunningStage = id
End Function

Public Function FSNumber(ByVal value As Variant, ByVal label As String, Optional ByVal defaultValue As Double = -1#) As Double
  If IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then
    If defaultValue < 0# Then FSCoupledInputError label & "を入力してください。"
    FSNumber = defaultValue: Exit Function
  End If
  If Not IsNumeric(value) Then FSCoupledInputError label & "は数値です。"
  FSNumber = CDbl(value)
  If Not P2IsFinite(FSNumber) Then FSCoupledInputError label & "は有限値です。"
End Function

Public Sub FSReset()
  FSActive = False: mReady = False: FSTime = 0#: FSAccepted = 0
  mPropertiesReady = False: mDrainSet = ""
  FSResumePending = False: FSRunningStage = 0: FSStageElapsed = 0#: FSStageDuration = 0#
  FSSrmBranch = False
  FSLedgerReset
  mGeneration = -1: FSFluidResidual = 0#: FSMassError = 0#
End Sub

Private Function FSFormat(ByVal value As Double) As String
  FSFormat = Replace$(Format$(value, "0.0000000000000000E+00"), Application.International(xlDecimalSeparator), ".")
End Function

Public Function FSWriteCheckpoint(ByVal folder As String) As Boolean
  Dim f As Integer, i As Long, e As Long, g As Long
  FSWriteCheckpoint = False
  If Not FSActive Then FSWriteCheckpoint = True: Exit Function
  Dim failedNumber As Long
  On Error GoTo Failed
  f = FreeFile
  Open folder & Application.PathSeparator & "fixed_stress.csv" For Output As #f
  Print #f, "FIXED_STRESS_V3," & GWInputSignature() & "," & CStr(NumberOfFreeNode)
  Print #f, "STATE," & FSFormat(FSTime) & "," & CStr(FSAccepted) & "," & FSFormat(FSFluidResidual) & "," & FSFormat(FSMassError) & "," & CStr(FSOuterIteration)
  Print #f, "DRAIN," & mDrainSet
  Print #f, "PHASE," & CStr(FSRunningStage) & "," & FSFormat(FSStageElapsed) & "," & FSFormat(FSStageDuration) & "," & FSFormat(mStageDt) & "," & IIf(FSResumePending, "1", "0")
  Print #f, "START_PROFILE," & mStageStartProfile
  Print #f, "node,used,pressure_kPa,reference_kPa,boundary_inflow_m3s"
  For i = 0 To NumberOfFreeNode - 1
    Print #f, CStr(i) & "," & IIf(mUsed(i), "1", "0") & "," & FSFormat(mOldP(i)) & "," & FSFormat(mReference(i)) & "," & FSFormat(mBoundaryFlow(i))
  Next i
  Print #f, "dof,base_body,base_applied,base_boundary,end_body,end_applied,end_boundary"
  For i = 0 To lastDof
    Print #f, CStr(i) & "," & FSFormat(mStageBaseSelf(i)) & "," & FSFormat(mStageBaseApplied(i)) & "," & FSFormat(mStageBaseBoundary(i)) & "," & FSFormat(mStageEndSelf(i)) & "," & FSFormat(mStageEndApplied(i)) & "," & FSFormat(mStageEndBoundary(i))
  Next i
  Print #f, "element,gauss,reference_pressure_kPa"
  For e = 0 To NumberOfElement - 1
    For g = 0 To 3
      Print #f, CStr(e + 1) & "," & CStr(g + 1) & "," & FSFormat(mReferenceGp(g, e))
    Next g
  Next e
  Close #f
  If Not FSLedgerWrite(folder) Then Exit Function
  FSWriteCheckpoint = True: Exit Function
Failed:
  failedNumber = Err.number
  On Error Resume Next
  If f > 0 Then Close #f
  On Error GoTo 0
  If failedNumber = 18 Then
    P3UserCancel = True: P3SearchFatal = True
    Err.Raise 18
  End If
End Function

Public Sub FSReadCheckpoint(ByVal folder As String)
  Dim f As Integer, line As String, parts() As String, i As Long, value As Double, e As Long, g As Long
  Dim savedTime As Double, savedCount As Long, p() As Double, reference() As Double, expected As String
  Dim savedFluid As Double, savedMass As Double, savedOuter As Long
  Dim locked() As Boolean, boundary() As Double
  If Dir$(folder & Application.PathSeparator & "fixed_stress.csv") = "" Then Exit Sub
  f = FreeFile
  Open folder & Application.PathSeparator & "fixed_stress.csv" For Input As #f
  Line Input #f, line
  expected = "FIXED_STRESS_V3," & GWInputSignature() & "," & CStr(NumberOfFreeNode)
  If line <> expected Then Close #f: FSCoupledInputError "時間連成チェックポイントの版・入力・節点数が一致しません。"
  Line Input #f, line: parts = Split(line, ",")
  If UBound(parts) <> 5 Then Close #f: FSCoupledInputError "時間連成の時刻情報が不完全です。"
  savedTime = FSCheckpointNumber(parts(1)): savedCount = FSCheckpointInteger(parts(2))
  savedFluid = FSCheckpointNumber(parts(3)): savedMass = FSCheckpointNumber(parts(4)): savedOuter = FSCheckpointInteger(parts(5))
  If parts(0) <> "STATE" Or savedTime < 0# Or savedCount < 0 Then Close #f: FSCoupledInputError "時間連成の時刻情報が不正です。"
  FSPrepare True
  Line Input #f, line
  If Left$(line, 6) <> "DRAIN," Then Close #f: FSCoupledInputError "排水状態がありません。"
  mDrainSet = Mid$(line, 7)
  Line Input #f, line: parts = Split(line, ",")
  If UBound(parts) <> 5 Then Close #f: FSCoupledInputError "時間工程の途中状態が不完全です。"
  If parts(0) <> "PHASE" Then Close #f: FSCoupledInputError "時間工程情報がありません。"
  FSRunningStage = FSCheckpointInteger(parts(1)): FSStageElapsed = FSCheckpointNumber(parts(2)): FSStageDuration = FSCheckpointNumber(parts(3)): mStageDt = FSCheckpointNumber(parts(4))
  If parts(5) <> "0" And parts(5) <> "1" Then Close #f: FSCoupledInputError "途中状態フラグが不正です。"
  FSResumePending = (parts(5) = "1")
  If FSRunningStage <= 0 Or FSStageElapsed < 0# Or FSStageDuration < FSStageElapsed Or mStageDt < 0# Then Close #f: FSCoupledInputError "時間工程の途中状態が不正です。"
  Line Input #f, line
  If Left$(line, 14) <> "START_PROFILE," Then Close #f: FSCoupledInputError "工程開始水位がありません。"
  mStageStartProfile = Mid$(line, 15)
  ReDim p(NumberOfFreeNode - 1): ReDim reference(NumberOfFreeNode - 1)
  Line Input #f, line
  If line <> "node,used,pressure_kPa,reference_kPa,boundary_inflow_m3s" Then Close #f: FSCoupledInputError "時間連成チェックポイントの列が不正です。"
  For i = 0 To NumberOfFreeNode - 1
    If EOF(f) Then Close #f: FSCoupledInputError "時間連成チェックポイントが不足しています。"
    Line Input #f, line: parts = Split(line, ",")
    If UBound(parts) <> 4 Then Close #f: FSCoupledInputError "時間連成チェックポイントが不完全です。"
    If parts(0) <> CStr(i) Or parts(1) <> IIf(mUsed(i), "1", "0") Then Close #f: FSCoupledInputError "時間連成の圧力接続が一致しません。"
    If Not IsNumeric(parts(2)) Or Not IsNumeric(parts(3)) Then Close #f: FSCoupledInputError "時間連成チェックポイントに非数値があります。"
    p(i) = FSCheckpointNumber(parts(2)): reference(i) = FSCheckpointNumber(parts(3))
    mBoundaryFlow(i) = FSCheckpointNumber(parts(4))
    If Not P2IsFinite(p(i)) Or Not P2IsFinite(reference(i)) Then Close #f: FSCoupledInputError "時間連成チェックポイントが非有限値です。"
  Next i
  Line Input #f, line
  If line <> "dof,base_body,base_applied,base_boundary,end_body,end_applied,end_boundary" Then Close #f: FSCoupledInputError "時間工程の荷重履歴がありません。"
  ReDim mStageBaseSelf(lastDof): ReDim mStageBaseApplied(lastDof): ReDim mStageBaseBoundary(lastDof)
  ReDim mStageEndSelf(lastDof): ReDim mStageEndApplied(lastDof): ReDim mStageEndBoundary(lastDof)
  For i = 0 To lastDof
    If EOF(f) Then Close #f: FSCoupledInputError "時間工程の荷重履歴が不足しています。"
    Line Input #f, line: parts = Split(line, ",")
    If UBound(parts) <> 6 Then Close #f: FSCoupledInputError "時間工程の荷重履歴が不完全です。"
    If parts(0) <> CStr(i) Then Close #f: FSCoupledInputError "時間工程の荷重自由度が不正です。"
    mStageBaseSelf(i) = FSCheckpointNumber(parts(1)): mStageBaseApplied(i) = FSCheckpointNumber(parts(2)): mStageBaseBoundary(i) = FSCheckpointNumber(parts(3))
    mStageEndSelf(i) = FSCheckpointNumber(parts(4)): mStageEndApplied(i) = FSCheckpointNumber(parts(5)): mStageEndBoundary(i) = FSCheckpointNumber(parts(6))
  Next i
  Line Input #f, line
  If line <> "element,gauss,reference_pressure_kPa" Then Close #f: FSCoupledInputError "過剰水圧の基準場がありません。"
  For e = 0 To NumberOfElement - 1
    For g = 0 To 3
      If EOF(f) Then Close #f: FSCoupledInputError "過剰水圧の基準場が不足しています。"
      Line Input #f, line: parts = Split(line, ",")
      If UBound(parts) <> 2 Then Close #f: FSCoupledInputError "過剰水圧の基準場が不完全です。"
      If parts(0) <> CStr(e + 1) Or parts(1) <> CStr(g + 1) Then Close #f: FSCoupledInputError "過剰水圧の基準場の要素・GPが不正です。"
      mReferenceGp(g, e) = FSCheckpointNumber(parts(2))
    Next g
  Next e
  If Not EOF(f) Then Close #f: FSCoupledInputError "時間連成チェックポイントに余分な行があります。"
  Close #f
  mP = p: mOldP = p: mReference = reference
  FSTime = savedTime: FSAccepted = savedCount
  FSFluidResidual = savedFluid: FSMassError = savedMass: FSOuterIteration = savedOuter
  FSRestoreTimeBoundary
  GWTransientBoundary locked, boundary
  FSDrainApply mDrainSet, locked, boundary, mSource
  GWTransientFixed locked
  FSUpdateField mP, True
  FSLedgerRead folder
  FEMIncrementReadFile folder
End Sub

Private Sub FSRestoreTimeBoundary()
  Dim i As Long, unit As Double, waterMode As String, loadMode As String, minimum As Double, maximum As Double, interval As Double
  Dim series As String, p1 As String, p2 As String, fraction As Double, knot As Double
  For i = 1 To P3StageN
    If P3StageId(i) = FSRunningStage Then
      If P3StageKind(i) = "WATER" Then
        FSGetOptions FSRunningStage, unit, waterMode, loadMode, minimum, maximum, interval, series
        p1 = mStageStartProfile: p2 = p1
        If Len(series) > 0 Then
          FSCurveAt series, FSStageElapsed, unit, FSStageDuration, p1, p2, fraction, knot
        ElseIf waterMode = "START" Or waterMode = "LINEAR" Then
          p2 = GWResolveStageProfile(FSRunningStage, P3StageGroup(i))
          If p2 = "" Then p2 = p1
          If waterMode = "LINEAR" Then
            fraction = FSStageElapsed / FSStageDuration
          Else
            fraction = 1#
          End If
        End If
        GWTransientProfile p1, p2, fraction
        GWTransientConductivity mKx
      End If
      Exit Sub
    End If
  Next i
  FSCoupledInputError "保存された時間工程IDが現在の工程表にありません。"
End Sub

Public Function FSBiotAlpha(ByVal e As Long) As Double
  FSBiotAlpha = 1#
  If mPropertiesReady Then
    If e >= 0 And e < NumberOfElement Then
      If mAlpha(Elem(e).MatNo) > 0# Then FSBiotAlpha = mAlpha(Elem(e).MatNo)
    End If
  End If
End Function

Public Sub FSPrime()
  Dim i As Long, r As Long
  FSValidateTimeDefinitions
  For i = 1 To P3StageN
    If P3StageOn(i) Then
      r = FSPlanRow(P3StageId(i))
      If r > 0 Then
        If UCase$(Trim$(CStr(ThisWorkbook.Worksheets("水条件").cells(r, 13).value2))) = "FIXED_STRESS" Then
          FSLoadMaterials
          Exit Sub
        End If
      End If
    End If
  Next i
End Sub

Private Sub FSLoadMaterials()
  Dim ws As Worksheet, r As Long, mat As Long, seen() As Boolean, e As Long, materialId As Double
  Set ws = ThisWorkbook.Worksheets("水条件")
  mGamma = FSNumber(ws.range("B4").value2, "水の単位体積重量")
  If mGamma <= 0# Then FSCoupledInputError "水の単位体積重量は正数です。"
  ReDim mPorosity(NumberOfMaterial - 1)
  ReDim seen(NumberOfMaterial - 1): ReDim mAlpha(NumberOfMaterial - 1)
  ReDim mStorage(NumberOfMaterial - 1): ReDim mKx(NumberOfMaterial - 1): ReDim mKy(NumberOfMaterial - 1)
  For r = 12 To 211
    If Len(CStr(ws.cells(r, 22).value2)) > 0 Then
      materialId = FSNumber(ws.cells(r, 22).value2, "連成材料番号")
      If materialId <> Fix(materialId) Then FSCoupledInputError "連成材料番号は整数です。"
      If materialId < 1# Or materialId > NumberOfMaterial Then FSCoupledInputError "連成材料番号が範囲外です。"
      mat = CLng(materialId) - 1
      If mat < 0 Or mat >= NumberOfMaterial Then FSCoupledInputError "連成材料番号が範囲外です。"
      If seen(mat) Then FSCoupledInputError "連成材料番号が重複しています。"
      seen(mat) = True
      mPorosity(mat) = FSNumber(ws.cells(r, 28).value2, "間隙率n", 0#)
      If mPorosity(mat) < 0# Or mPorosity(mat) >= 1# Then FSCoupledInputError "間隙率nは0<n<1です。空欄は領域増減のない解析に限ります。"
      mAlpha(mat) = FSNumber(ws.cells(r, 23).value2, "Biot係数")
      mStorage(mat) = FSNumber(ws.cells(r, 24).value2, "貯留1/M [1/kPa]")
      mKx(mat) = FSNumber(ws.cells(r, 25).value2, "kx [m/s]")
      mKy(mat) = FSNumber(ws.cells(r, 26).value2, "ky [m/s]")
      If mAlpha(mat) <= 0# Or mAlpha(mat) > 1# Or mStorage(mat) < 0# Or mKx(mat) <= 0# Or mKy(mat) <= 0# Then FSCoupledInputError "連成物性は0<α<=1、1/M>=0、kx,ky>0です。材料=" & CStr(mat + 1)
    End If
  Next r
  For e = 0 To NumberOfElement - 1
    mat = Elem(e).MatNo
    If Material(mat).kind = "SOIL" Then
      If GWMaterialEffective(mat) Then
        If Not seen(mat) Then FSCoupledInputError "EFFECTIVE地盤材料の連成物性が未設定です。材料=" & CStr(mat + 1)
      ElseIf seen(mat) Then
        FSCoupledInputError "時間連成の地盤材料はEFFECTIVE（排水骨格定数）で指定してください。"
      End If
    End If
  Next e
  mPropertiesReady = True
End Sub

Public Sub FSPrepare(Optional ByVal fromCheckpoint As Boolean = False)
  Dim e As Long, a As Long, b As Long, d As Long, mat As Long, node As Long
  Dim lm() As Double, lh() As Double, ll() As Double, lq() As Double, lg() As Double
  Dim newUsed() As Boolean, initialP() As Double, pressure As Double, volume As Double, initiallyReady As Boolean, g As Long
  If Not GWActive Then FSCoupledInputError "時間連成には初期水圧場が必要です。"
  If P6MixedUP Or P6ConsolActive Then FSCoupledInputError "時間連成と旧MIXED_UP/CONSOLは併用できません。"
  FSLoadMaterials
  initiallyReady = mReady
  ReDim mMap(0 To 3, 0 To NumberOfElement - 1)
  ReDim mMass(0 To 3, 0 To 3, 0 To NumberOfElement - 1)
  ReDim mFlow(0 To 3, 0 To 3, 0 To NumberOfElement - 1)
  ReDim mStab(0 To 3, 0 To 3, 0 To NumberOfElement - 1)
  ReDim mQ(0 To 15, 0 To 3, 0 To NumberOfElement - 1)
  ReDim mGravity(0 To 3, 0 To NumberOfElement - 1)
  ReDim newUsed(NumberOfFreeNode - 1): ReDim initialP(NumberOfFreeNode - 1)
  ReDim mBoundaryFlow(NumberOfFreeNode - 1)
  If Not mReady Then
    ReDim mReferenceGp(0 To 3, 0 To NumberOfElement - 1)
    ReDim mP(NumberOfFreeNode - 1): ReDim mOldP(NumberOfFreeNode - 1)
    ReDim mReference(NumberOfFreeNode - 1): ReDim mUsed(NumberOfFreeNode - 1)
  End If
  For e = 0 To NumberOfElement - 1
    For a = 0 To 3: mMap(a, e) = -1: Next a
    mat = Elem(e).MatNo
    If P3IsElementActive(e) And Material(mat).kind = "SOIL" And mAlpha(mat) > 0# Then
      FSLocalMatrices e, mAlpha(mat), mStorage(mat), mKx(mat), mKy(mat), mGamma, lm, lh, ll, lq, lg
      volume = 0#
      For a = 0 To 3
        For b = 0 To 3: volume = volume + ll(a, b): Next b
      Next a
      volume = volume * Material(mat).young / (3# * (1# - 2# * Material(mat).Poisson)) / (mAlpha(mat) * mAlpha(mat))
      FSLedgerTopology e, True, volume, mPorosity(mat), (Not initiallyReady Or fromCheckpoint)
      For a = 0 To 3
        If a < 3 Or Not FEMElementIsT6(e) Then
          node = GWFlowNode(e, a): mMap(a, e) = node
          pressure = 0#
          If Not fromCheckpoint Then
            If mUsed(node) Then
              pressure = mP(node)
            Else
              pressure = GWInitialNodePressure(e, a)
            End If
          End If
          If pressure < -0.0000001 Then FSCoupledInputError "時間連成の初期領域が飽和していません。要素=" & CStr(e + 1)
          If newUsed(node) And Abs(initialP(node) - pressure) > 0.0000001 Then FSCoupledInputError "共有圧力節点の初期水圧が矛盾しています。"
          initialP(node) = pressure: newUsed(node) = True
          If Not mUsed(node) Then mP(node) = pressure: mReference(node) = pressure
        End If
        mGravity(a, e) = lg(a)
        For b = 0 To 3
          mMass(a, b, e) = lm(a, b): mFlow(a, b, e) = lh(a, b): mStab(a, b, e) = ll(a, b)
        Next b
        For d = 0 To 15: mQ(d, a, e) = lq(d, a): Next d
      Next a
    Else
      FSLedgerTopology e, False, 0#, 0#, (Not initiallyReady Or fromCheckpoint)
      If P3IsElementActive(e) And Material(mat).kind = "SOIL" Then
      For a = 0 To 7
        If GWInitialNodePressure(e, a) > 0.0000001 Then FSCoupledInputError "時間連成から除外するTOTAL地盤は乾燥領域に限ります。湿潤地盤には連成物性が必要です。"
      Next a
      End If
    End If
  Next e
  mUsed = newUsed: mOldP = mP
  mReady = True: FSActive = True: mGeneration = P3ActiveSetGen
  GWTransientConductivity mKx
  FSUpdateField mP, False
  For e = 0 To NumberOfElement - 1
    If Not initiallyReady Or FSLedgerBorn(e) Then
      For g = 0 To 3: mReferenceGp(g, e) = mGp(g, e): Next g
    End If
  Next e
  GWCommit
End Sub

Private Sub FSUpdateField(ByRef pressures() As Double, ByVal commitField As Boolean)
  Dim e As Long, g As Long, a As Long, n As Long, xi As Double, eta As Double, value As Double
  Dim pn(0 To 3) As Double, dx(0 To 3) As Double, dy(0 To 3) As Double
  If commitField Then FSCheckSaturation pressures
  ReDim mGp(0 To 3, 0 To NumberOfElement - 1)
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Material(Elem(e).MatNo).kind = "SOIL" Then
      For g = 0 To 3
        FEMGaussNatural g, xi, eta: FSPressureShape FEMElementIsT6(e), xi, eta, pn, dx, dy
        value = 0#
        For a = 0 To 3
          n = mMap(a, e)
          If n >= 0 Then value = value + pn(a) * pressures(n)
        Next a
        If commitField And value < -0.0000001 Then FSCoupledInputError "負の間隙水圧が発生しました。自由水面・不飽和は適用外です。"
        mGp(g, e) = value
      Next g
    End If
  Next e
  GWSetTransientPressure mGp
  If commitField Then GWCommit
End Sub

Public Function FSNodePressure(ByVal e As Long, ByVal slot As Long) As Double
  Dim a As Long, b As Long
  If slot < 4 Then
    a = slot
    If FEMElementIsT6(e) And slot = 3 Then a = 2
    If mMap(a, e) >= 0 Then FSNodePressure = mP(mMap(a, e))
  Else
    a = slot - 4: b = (a + 1) Mod 4
    If FEMElementIsT6(e) Then
      If a = 2 Then a = 2: b = 2
      If a = 3 Then a = 2
    End If
    If mMap(a, e) >= 0 And mMap(b, e) >= 0 Then FSNodePressure = (mP(mMap(a, e)) + mP(mMap(b, e))) / 2#
  End If
End Function

Public Function FSFacePressure(ByVal e As Long, ByVal side As Long, ByVal xi As Double, ByVal firstNode As Long) As Double
  If GJTopoNode(Elem(e).node(side)) <> GJTopoNode(firstNode) Then xi = -xi
  FSFacePressure = (1# - xi) / 2# * FSNodePressure(e, side) + (1# + xi) / 2# * FSNodePressure(e, (side + 1) Mod 4)
End Function

Public Sub FSViewHydraulics(ByRef heads() As Double, ByRef qx() As Double, ByRef qy() As Double, ByRef reactionFlow() As Double)
  Dim e As Long, slot As Long, g As Long, a As Long, i As Long, n As Long, mat As Long
  Dim xi As Double, eta As Double, j11 As Double, j12 As Double, j21 As Double, j22 As Double, det As Double, gx As Double, gy As Double
  Dim ns(0 To 7) As Double, dx(0 To 7) As Double, dy(0 To 7) As Double
  Dim pn(0 To 3) As Double, px(0 To 3) As Double, py(0 To 3) As Double
  ReDim heads(NumberOfFreeNode - 1): ReDim qx(0 To 3, 0 To NumberOfElement - 1): ReDim qy(0 To 3, 0 To NumberOfElement - 1)
  ReDim reactionFlow(NumberOfFreeNode - 1)
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Material(Elem(e).MatNo).kind = "SOIL" Then
      mat = Elem(e).MatNo
      For slot = 0 To 7
        n = Elem(e).node(slot): heads(n) = Elem(e).y(slot) + FSNodePressure(e, slot) / mGamma
      Next slot
      For g = 0 To 3
        FEMGaussNatural g, xi, eta
        FEMShapeForElement e, xi, eta, ns, dx, dy
        FSPressureShape FEMElementIsT6(e), xi, eta, pn, px, py
        j11 = 0#: j12 = 0#: j21 = 0#: j22 = 0#: gx = 0#: gy = 0#
        For i = 0 To 7
          j11 = j11 + dx(i) * Elem(e).x(i): j12 = j12 + dy(i) * Elem(e).x(i)
          j21 = j21 + dx(i) * Elem(e).y(i): j22 = j22 + dy(i) * Elem(e).y(i)
        Next i
        det = j11 * j22 - j12 * j21
        For a = 0 To 3
          n = mMap(a, e)
          If n >= 0 Then
            gx = gx + (j22 * px(a) - j21 * py(a)) / det * mP(n)
            gy = gy + (-j12 * px(a) + j11 * py(a)) / det * mP(n)
          End If
        Next a
        qx(g, e) = -mKx(mat) * gx / mGamma: qy(g, e) = -mKy(mat) * (gy / mGamma + 1#)
      Next g
    End If
  Next e
  On Error Resume Next
  reactionFlow = mBoundaryFlow
  Err.Clear: On Error GoTo 0
End Sub

' Replace the 2x2 pressure-force quadrature by the same operator as continuity.
' Constitutive history still uses the existing four material points.
Public Sub FSForceCorrection(ByRef internal() As Double)
  Dim e As Long, d As Long, a As Long, g As Long, node As Long, dof As Long, correction As Double, biot As Double
  If Not FSActive Or Not mReady Then Exit Sub
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Material(Elem(e).MatNo).kind = "SOIL" Then
      biot = mAlpha(Elem(e).MatNo)
      For d = 0 To 15
        dof = Elem(e).elnode(d)
        If dof >= 0 Then
          correction = 0#
          For a = 0 To 3
            node = mMap(a, e)
            If node >= 0 Then correction = correction - mQ(d, a, e) * mP(node)
          Next a
          For g = 0 To 3
            correction = correction - (Elem(e).bMat(0, d, g) + Elem(e).bMat(1, d, g)) * biot * mGp(g, e) * Elem(e).dj(g) * Material(Elem(e).MatNo).thickness
          Next g
          internal(dof) = internal(dof) + correction
        End If
      Next d
    End If
  Next e
End Sub

Private Sub FSFluidRhs(ByVal dt As Double, ByRef previousU() As Double, ByRef guessU() As Double, ByRef guessP() As Double, ByRef rhs() As Double)
  Dim e As Long, a As Long, b As Long, d As Long, ia As Long, ib As Long, dof As Long
  ReDim rhs(NumberOfFreeNode - 1)
  For e = 0 To NumberOfElement - 1
    For a = 0 To 3
      ia = mMap(a, e)
      If ia >= 0 Then
        rhs(ia) = rhs(ia) - dt * mGravity(a, e)
        For b = 0 To 3
          ib = mMap(b, e)
          If ib >= 0 Then rhs(ia) = rhs(ia) + mMass(a, b, e) * mOldP(ib) + mStab(a, b, e) * guessP(ib)
        Next b
        For d = 0 To 15
          dof = Elem(e).elnode(d)
          If dof >= 0 Then rhs(ia) = rhs(ia) - mQ(d, a, e) * (guessU(dof) - previousU(dof))
        Next d
      End If
    Next a
  Next e
  For ia = 0 To NumberOfFreeNode - 1: rhs(ia) = rhs(ia) + dt * mSource(ia): Next ia
End Sub

Private Function FSTryStep(ByVal dt As Double, ByRef target() As Double, ByRef locked() As Boolean, ByRef prescribed() As Double, ByVal limit As Long, ByVal tolerance As Double) As Boolean
  Dim blocks() As Double, rhs() As Double, residual() As Double, reactionFlow() As Double
  Dim previousU() As Double, trialU() As Double, guessP() As Double, nextP() As Double, internal() As Double, acceptedRhs() As Double
  Dim e As Long, a As Long, b As Long, i As Long, ia As Long, ib As Long, d As Long, dof As Long
  Dim num As Double, den As Double, value As Double, sum As Double, total As Double, term As Double
  Dim waterBase() As Double, waterTrial() As Double, workingTarget() As Double
  Dim priorU() As Double, pressureChange As Double, displacementChange As Double, pressureNorm As Double, displacementNorm As Double
  On Error GoTo Rejected
  ReDim blocks(0 To 3, 0 To 3, 0 To NumberOfElement - 1)
  For e = 0 To NumberOfElement - 1
    For a = 0 To 3
      For b = 0 To 3: blocks(a, b, e) = mMass(a, b, e) + dt * mFlow(a, b, e) + mStab(a, b, e): Next b
    Next a
  Next e
  GWTransientForce waterBase
  ReDim workingTarget(lastDof)
  previousU = P3CommittedDisp: trialU = previousU: guessP = mOldP
  For FSOuterIteration = 1 To limit
    FSFluidRhs dt, previousU, trialU, guessP, rhs
    If Not FSFlowSolve(blocks, mMap, NumberOfElement, rhs, locked, prescribed, guessP, 0.0000000001, NumberOfFreeNode * 4 + 100, nextP, residual, reactionFlow) Then Exit Function
    priorU = trialU
    mP = nextP: FSUpdateField mP, False
    GWTransientForce waterTrial
    For i = 0 To lastDof: workingTarget(i) = target(i) + waterTrial(i) - waterBase(i): Next i
    If Not P3FixedStressTrial(workingTarget, previousU, trialU, internal, tolerance * 0.01) Then
      If P3UserCancel Then Err.Raise 18
      P6SolverEvent "FIXED_STRESS_REJECT", "reason=MECHANICS;outer=" & CStr(FSOuterIteration) & ";message=" & AnalysisMessage
      Exit Function
    End If
    trialU = TDisp
    ' Original continuity equation, without artificial stabilization.
    num = 0#: den = 0#: sum = 0#: total = 0#
    ReDim acceptedRhs(NumberOfFreeNode - 1)
    For e = 0 To NumberOfElement - 1
      For a = 0 To 3
        ia = mMap(a, e)
        If ia >= 0 Then
          value = dt * mGravity(a, e): den = den + Abs(value)
          For b = 0 To 3
            ib = mMap(b, e)
            If ib >= 0 Then
              term = mMass(a, b, e) * (mP(ib) - mOldP(ib))
              value = value + term: den = den + Abs(term)
              term = dt * mFlow(a, b, e) * mP(ib)
              value = value + term: den = den + Abs(term)
            End If
          Next b
          For d = 0 To 15
            dof = Elem(e).elnode(d)
            If dof >= 0 Then
              term = mQ(d, a, e) * (trialU(dof) - previousU(dof))
              value = value + term: den = den + Abs(term)
            End If
          Next d
          acceptedRhs(ia) = acceptedRhs(ia) + value
        End If
      Next a
    Next e
    For i = 0 To NumberOfFreeNode - 1
      acceptedRhs(i) = acceptedRhs(i) - dt * mSource(i)
      If mUsed(i) And Not locked(i) Then
        num = num + acceptedRhs(i) ^ 2: sum = sum + acceptedRhs(i): total = total + Abs(acceptedRhs(i))
      End If
    Next i
    pressureChange = 0#: pressureNorm = 1#: displacementChange = 0#: displacementNorm = 0.000000000001
    For i = 0 To NumberOfFreeNode - 1
      If mUsed(i) Then
        den = den + Abs(dt * mSource(i))
        If Abs(mP(i) - guessP(i)) > pressureChange Then pressureChange = Abs(mP(i) - guessP(i))
        If Abs(mP(i)) > pressureNorm Then pressureNorm = Abs(mP(i))
      End If
    Next i
    For i = 0 To lastDof
      If Abs(trialU(i) - priorU(i)) > displacementChange Then displacementChange = Abs(trialU(i) - priorU(i))
      If Abs(trialU(i) - previousU(i)) > displacementNorm Then displacementNorm = Abs(trialU(i) - previousU(i))
    Next i
    If den < 1E-20 Then den = 1E-20
    FSFluidResidual = Sqr(num) / den: FSMassError = Abs(sum) / den
    If FSOuterIteration <= 3 Or FSOuterIteration Mod 10 = 0 Then P6SolverEvent "FIXED_STRESS_ITER", "outer=" & CStr(FSOuterIteration) & ";fluid=" & CStr(FSFluidResidual) & ";mass=" & CStr(FSMassError) & ";mechanics=" & CStr(RelativeResidualFree)
    If FSFluidResidual <= tolerance And FSMassError <= tolerance And pressureChange / pressureNorm <= tolerance And displacementChange / displacementNorm <= tolerance And RelativeResidualFree <= P3_ENGINEERING_RESIDUAL Then
      For e = 0 To NumberOfElement - 1
        If Elem(e).IsJoint And P3IsElementActive(e) Then
          For a = 0 To 2
            If Elem(e).Stmat(2, a) > 0.00000001 Then FSCoupledInputError "時間連成でJOINTの開口が発生しました。開口内の流体貯留を含むモデルが必要です。"
          Next a
        End If
      Next e
      FSCheckSaturation mP
      ReDim mBoundaryFlow(NumberOfFreeNode - 1)
      If dt > 0# Then
        For i = 0 To NumberOfFreeNode - 1
          If mUsed(i) And locked(i) Then mBoundaryFlow(i) = acceptedRhs(i) / dt
        Next i
      End If
      FSLedgerStepStart
      For e = 0 To NumberOfElement - 1
        num = 0#: den = 0#
        For a = 0 To 3
          ia = mMap(a, e)
          If ia >= 0 Then
            For b = 0 To 3
              ib = mMap(b, e)
              If ib >= 0 Then num = num + mMass(a, b, e) * (mP(ib) - mOldP(ib))
            Next b
            For d = 0 To 15
              dof = Elem(e).elnode(d)
              If dof >= 0 Then den = den + mQ(d, a, e) * (trialU(dof) - previousU(dof))
            Next d
          End If
        Next a
        FSLedgerIncrement e, num, den
      Next e
      num = 0#: den = 0#
      For i = 0 To NumberOfFreeNode - 1
        If mUsed(i) Then
          If locked(i) Then num = num + acceptedRhs(i)
          den = den + dt * mSource(i)
        End If
      Next i
      FSLedgerClose num, den
      target = workingTarget
      P3FixedStressAccept internal
      P3CommittedNodeCond = NodeCond: P3CommittedBoundaryReady = True
      mOldP = mP: mOldGp = mGp
      FSTime = FSTime + dt: FSAccepted = FSAccepted + 1
      FSUpdateField mP, True
      FSTryStep = True: Exit Function
    End If
    guessP = mP
    If FSOuterIteration Mod 10 = 0 Then DoEvents
  Next FSOuterIteration
  Exit Function
Rejected:
  If Err.number = 18 Then
    P3UserCancel = True: P3SearchFatal = True
    Err.Raise 18
  End If
  P6SolverEvent "FIXED_STRESS_REJECT", "reason=" & Err.description
End Function

Public Function FSRunStage(ByVal stageId As Long, ByVal kind As String) As Boolean
  Dim r As Long, ws As Worksheet, duration As Double, dt As Double, elapsed As Double, take As Double, minimum As Double
  Dim tolerance As Double, limit As Long, target() As Double, locked() As Boolean, prescribed() As Double
  Dim i As Long, factor As Double, baseSelf() As Double, baseApplied() As Double, finalSelf() As Double, finalApplied() As Double
  Dim baseBoundary() As Double, finalBoundary() As Double, acceptedBoundary() As Double
  Dim unit As Double, waterMode As String, loadMode As String, maximum As Double, saveEvery As Double, series As String
  Dim firstProfile As String, lastProfile As String, p1 As String, p2 As String, fraction As Double, nextKnot As Double, saveTarget As Double
  Dim instantPending As Boolean, isInstant As Boolean, endTime As Double, continuing As Boolean
  Dim acceptedWater As Variant, acceptedLedger As Variant, gaugeChecked As Boolean, failureReason As String
  Dim currentWater() As Double
  Dim failedNumber As Long, failedSource As String, failedDescription As String
  On Error GoTo RunFailed
  Set ws = ThisWorkbook.Worksheets("水条件"): r = FSPlanRow(stageId)
  If r = 0 And (kind = "WAIT" Or kind = "WATER") Then FSCoupledInputError "時間経過・水理変動には時間設定が必要です。工程=" & CStr(stageId)
  FSGetOptions stageId, unit, waterMode, loadMode, minimum, maximum, saveEvery, series
  If r > 0 Then
    duration = FSNumber(ws.cells(r, 15).value2, "工程時間", 0#) * unit
    dt = FSNumber(ws.cells(r, 16).value2, "初期刻み", duration / unit) * unit
    factor = FSNumber(ws.cells(r, 18).value2, "連成反復上限", 100#)
    If factor <> Fix(factor) Or factor > 100000# Then FSCoupledInputError "連成反復上限は100000以下の正の整数です。"
    limit = CLng(factor)
    tolerance = FSNumber(ws.cells(r, 19).value2, "連成許容値", 0.00000001)
  Else
    limit = 100: tolerance = 0.00000001
  End If
  If duration < 0# Or dt < 0# Or (duration > 0# And dt <= 0#) Or limit < 1 Or tolerance <= 0# Or tolerance > 0.000001 Then FSCoupledInputError "工程時間・刻み・連成収束設定が不正です。"
  If kind = "WAIT" And duration <= 0# Then FSCoupledInputError "時間経過には正の工程継続時間が必要です。"
  If loadMode = "LINEAR" And duration <= 0# Then FSCoupledInputError "線形載荷には正の工程時間が必要です。"
  If waterMode = "LINEAR" And duration <= 0# Then FSCoupledInputError "線形水位変化には正の工程時間が必要です。"
  If kind <> "LOAD" And kind <> "APPLY" And kind <> "UNLOAD" And Len(FSLoadingSet(stageId)) > 0 Then FSCoupledInputError "載荷セットはLOAD/APPLY/UNLOADへ指定してください。"
  If kind <> "WATER" And (waterMode <> "HOLD" Or Len(series) > 0) Then FSCoupledInputError "境界水位の変更はWATER（水理変動）工程へ指定してください。"
  If kind = "WATER" And FSDetailRow(stageId) = 0 Then waterMode = "START"
  If Not mReady Or mGeneration <> P3ActiveSetGen Then FSPrepare
  If r > 0 Then
    If Len(Trim$(CStr(ws.cells(r, 17).value2))) > 0 Then mDrainSet = Trim$(CStr(ws.cells(r, 17).value2))
  End If
  continuing = FSResumingStage(stageId)
  firstProfile = GWViewDefaultProfile()
  If continuing Then firstProfile = mStageStartProfile
  lastProfile = firstProfile
  If kind = "WATER" Then lastProfile = GWResolveStageProfile(stageId, P3StageGroup(P3ActivePrefix))
  If lastProfile = "" Then lastProfile = firstProfile
  ReDim target(lastDof)
  baseSelf = P3LockedSelfWeight: baseApplied = P3LockedApplied
  finalSelf = P3SelfWeightForce: finalApplied = P3AppliedForce
  baseBoundary = P3CommittedBoundaryDisp: finalBoundary = P3BoundaryDisp: acceptedBoundary = baseBoundary
  If continuing Then
    baseSelf = mStageBaseSelf: baseApplied = mStageBaseApplied: baseBoundary = mStageBaseBoundary
    finalSelf = mStageEndSelf: finalApplied = mStageEndApplied: finalBoundary = mStageEndBoundary
    elapsed = FSStageElapsed: If mStageDt > 0# Then dt = mStageDt
  Else
    ' Interpolate only body/applied loads. Hydraulic tractions always belong
    ' to the actual time/pressure, independently of construction load ramps.
    GWTransientForce currentWater
    For i = 0 To lastDof
      baseSelf(i) = baseSelf(i) - mStageWaterBefore(i)
      finalSelf(i) = finalSelf(i) - currentWater(i)
    Next i
    FSRunningStage = stageId: FSStageElapsed = 0#: FSStageDuration = duration
    mStageStartProfile = firstProfile
    For i = 0 To lastDof
      If NodeCond(i) <> 0 And P3CommittedNodeCond(i) = 0 Then baseBoundary(i) = P3CommittedDisp(i)
    Next i
    mStageBaseSelf = baseSelf: mStageBaseApplied = baseApplied: mStageBaseBoundary = baseBoundary
    mStageEndSelf = finalSelf: mStageEndApplied = finalApplied: mStageEndBoundary = finalBoundary
  End If
  If minimum = 0# Then minimum = dt / 1024#
  If maximum = 0# Then maximum = dt
  If maximum > 0# And dt > maximum Then dt = maximum
  If minimum > dt And duration > 0# Then FSCoupledInputError "最小刻みは初期刻み以下にしてください。"
  instantPending = (duration = 0# Or ((kind = "APPLY" Or kind = "LOAD" Or kind = "UNLOAD" Or kind = "BIRTH" Or kind = "DEATH") And loadMode = "INSTANT"))
  If continuing Then instantPending = False
  If saveEvery > 0# Then saveTarget = saveEvery Else saveTarget = duration
  If continuing And saveEvery > 0# Then saveTarget = (Int(elapsed / saveEvery) + 1#) * saveEvery
  acceptedWater = GWTransientState()
  acceptedLedger = FSLedgerMemory()
  Do
    take = dt: isInstant = instantPending
    If isInstant Then
      take = 0#
    Else
      If take > duration - elapsed Then take = duration - elapsed
      If take > saveTarget - elapsed And saveTarget > elapsed Then take = saveTarget - elapsed
      If Len(series) > 0 Then
        FSCurveAt series, elapsed, unit, duration, p1, p2, fraction, nextKnot
        If take > nextKnot - elapsed Then take = nextKnot - elapsed
      End If
    End If
    endTime = elapsed + take
    p1 = firstProfile: p2 = firstProfile: fraction = 0#
    If kind = "WATER" Then
      If Len(series) > 0 Then
        FSCurveAt series, endTime, unit, duration, p1, p2, fraction, nextKnot
      ElseIf waterMode = "START" Then
        p1 = lastProfile: p2 = lastProfile
      ElseIf waterMode = "LINEAR" Then
        p2 = lastProfile: fraction = endTime / duration
      End If
      GWTransientProfile p1, p2, fraction
      GWTransientConductivity mKx
    End If
    GWTransientBoundary locked, prescribed
    FSDrainApply mDrainSet, locked, prescribed, mSource
    For i = 0 To NumberOfFreeNode - 1
      If Not mUsed(i) Then
        locked(i) = True: prescribed(i) = 0#
      End If
    Next i
    GWTransientFixed locked
    If Not gaugeChecked Then
      FSCheckPressureGauge mMap, mMass, mQ, locked
      gaugeChecked = True
    End If
    factor = 1#: If loadMode = "LINEAR" Then factor = endTime / duration
    GWTransientForce currentWater
    For i = 0 To lastDof
      target(i) = baseSelf(i) + factor * (finalSelf(i) - baseSelf(i)) + baseApplied(i) + factor * (finalApplied(i) - baseApplied(i)) + currentWater(i)
      P3BoundaryDisp(i) = baseBoundary(i) + factor * (finalBoundary(i) - baseBoundary(i))
    Next i
    If FSTryStep(take, target, locked, prescribed, limit, tolerance) Then
      elapsed = endTime: instantPending = False: acceptedBoundary = P3BoundaryDisp
      acceptedWater = GWTransientState()
      acceptedLedger = FSLedgerMemory()
      FSStageElapsed = elapsed: FSResumePending = (elapsed < duration): mStageDt = dt
      For i = 0 To lastDof
        P3LockedApplied(i) = baseApplied(i) + factor * (finalApplied(i) - baseApplied(i))
        P3LockedSelfWeight(i) = target(i) - P3LockedApplied(i)
        P3CommittedBoundaryDisp(i) = P3BoundaryDisp(i)
      Next i
      FSWriteAccepted stageId, take, (isInstant Or saveEvery = 0# Or elapsed >= saveTarget - 0.000000001 Or elapsed >= duration - 0.000000001)
      acceptedLedger = FSLedgerMemory()
      If saveEvery > 0# And elapsed >= saveTarget - 0.000000001 Then saveTarget = saveTarget + saveEvery
      If duration = 0# Or elapsed >= duration - 0.000000001 Then Exit Do
    Else
      If P3UserCancel Then Err.Raise 18
      failureReason = AnalysisMessage
      P3FixedStressRollback
      GWRestoreTransientState acceptedWater
      FSLedgerRestore acceptedLedger
      GWTransientConductivity mKx
      mP = mOldP: FSUpdateField mP, False: GWRestore: P3BoundaryDisp = acceptedBoundary
      If take <= minimum Or isInstant Or P3UserCancel Then
        SetAnalysisFailure RESULT_NONCONVERGED, "時間連成が収束しません。最後の受理時刻=" & CStr(FSTime) & " / " & failureReason, vbObjectError + 8121, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
        Exit Function
      End If
      dt = take / 2#
    End If
  Loop
  FSResumePending = False
  ' A resumed LINEAR stage started from accepted intermediate loads. Its
  ' completed end target must become the inherited target of the next stage.
  P3AppliedForce = finalApplied
  P3AssembleSelfWeightOnly
  FSRunStage = True
  Exit Function
RunFailed:
  failedNumber = Err.number: failedSource = Err.source: failedDescription = Err.description
  If failedNumber = 18 Then
    P3UserCancel = True: P3SearchFatal = True
    ' Restore the last accepted state; cancellation is never a cutback retry.
    On Error Resume Next
    P3FixedStressRollback
    GWRestoreTransientState acceptedWater
    FSLedgerRestore acceptedLedger
    GWTransientConductivity mKx
    mP = mOldP: FSUpdateField mP, False: GWRestore
    P3BoundaryDisp = acceptedBoundary
    On Error GoTo 0
  End If
  Err.Raise failedNumber, failedSource, failedDescription
End Function

Private Sub FSWriteAccepted(ByVal stageId As Long, ByVal dt As Double, ByVal saveField As Boolean)
  FSRecordState stageId, dt, False, saveField
  P6SolverEvent "FIXED_STRESS_ACCEPT", "stage=" & CStr(stageId) & ";time_s=" & CStr(FSTime) & ";dt=" & CStr(dt) & ";outer=" & CStr(FSOuterIteration) & ";fluid=" & CStr(FSFluidResidual)
End Sub

Public Function FSCheckpointNumber(ByVal text As String) As Double
  If Not IsNumeric(text) Then FSCoupledInputError "時間連成チェックポイントに非数値があります。"
  FSCheckpointNumber = Val(text)
  If Not P2IsFinite(FSCheckpointNumber) Then FSCoupledInputError "時間連成チェックポイントが非有限値です。"
End Function
Private Function FSCheckpointInteger(ByVal text As String) As Long
  Dim value As Double
  value = FSCheckpointNumber(text)
  If value <> Fix(value) Or value < 0# Or value > 2147483647# Then FSCoupledInputError "時間連成チェックポイントの整数が不正です。"
  FSCheckpointInteger = CLng(value)
End Function

Public Function FSReferencePressure(ByVal e As Long, ByVal gp As Long) As Double
  If FSActive And mReady Then FSReferencePressure = mReferenceGp(gp, e)
End Function

Public Sub FSInitializeBirth(ByVal materialId As Long)
  Dim e As Long, g As Long, pressure As Double, component As Variant
  For e = 0 To NumberOfElement - 1
    If P3IsElementActive(e) And Elem(e).MatNo = materialId - 1 And Material(Elem(e).MatNo).kind = "SOIL" Then
      If mAlpha(Elem(e).MatNo) > 0# And FSLedgerBorn(e) Then
        For g = 0 To 3
          pressure = mAlpha(Elem(e).MatNo) * mGp(g, e)
          For Each component In Array(0, 1, 3, 4, 16)
            Elem(e).Stmat(CLng(component), g) = pressure
            Elem(e).mStmat(CLng(component), g) = pressure
          Next component
        Next g
      End If
    End If
  Next e
  P3CommitMaterialState
End Sub

Public Function FSHydraulicMemory() As Variant
  FSHydraulicMemory = Array(mP, mReference, mUsed, GWTransientState(), FSLedgerMemory(), FSRunningStage, FSStageElapsed, FSStageDuration, FSResumePending, mStageStartProfile, mStageDt, FSFluidResidual, FSMassError, FSOuterIteration, mReferenceGp)
End Function
Public Sub FSRestoreHydraulicMemory(ByVal saved As Variant)
  mP = saved(0): mOldP = mP: mReference = saved(1): mUsed = saved(2): mReferenceGp = saved(14)
  GWRestoreTransientState saved(3)
  FSPrepare True
  FSLedgerRestore saved(4)
  FSRunningStage = saved(5): FSStageElapsed = saved(6): FSStageDuration = saved(7): FSResumePending = saved(8)
  mStageStartProfile = saved(9): mStageDt = saved(10): FSFluidResidual = saved(11): FSMassError = saved(12): FSOuterIteration = saved(13)
  FSUpdateField mP, True
End Sub

Private Sub FSCheckSaturation(ByRef pressure() As Double)
  Dim n As Long
  For n = 0 To NumberOfFreeNode - 1
    If mUsed(n) And pressure(n) < -0.0000001 Then FSCoupledInputError "受理候補に負の水圧があります。自由水面・不飽和は適用外です。節点=" & CStr(n + 1) & " p[kPa]=" & CStr(pressure(n))
  Next n
End Sub

Public Sub FSCaptureStageWater()
  GWTransientForce mStageWaterBefore
End Sub
