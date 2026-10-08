Option Explicit
Private mAccelAttemptFailure As String
Private mLineSearchFailure As String, mLineSearchTries As Long, mLineSearchMaterialRejects As Long
Private mLineSearchBestQ As Double


' P2 constitutive update, P3 Newton / stages / SRM.
Private mAAReady As Boolean, mAAFactor As Long
Private mAAX() As Double, mAAF() As Double
Private V3CostSearchEnabled As Boolean, V3BracketStep As Long
Private V3LastPassCostSec As Double, V3LastFailCostSec As Double
Private V3PassCostSamples As Long, V3FailCostSamples As Long
Private V3TrialStartedAt As Double
Public Function P3MaterialIsJoint(ByVal materialIndex As Long) As Boolean
  P3MaterialIsJoint = (P6MaterialKindOf(materialIndex) = "JOINT")
End Function

Public Function P3ElementIsJoint(ByVal elementId As Long) As Boolean
  P3ElementIsJoint = False
  If elementId < 0 Or elementId > NumberOfElement - 1 Then Exit Function
  P3ElementIsJoint = Elem(elementId).IsJoint
End Function

Private Function P3LastOriginIndex(ByVal beforeIndex As Long) As Long
  Dim i As Long, lastI As Long
  P3LastOriginIndex = 0
  lastI = beforeIndex
  If lastI > P3StageN Then lastI = P3StageN
  For i = 1 To lastI
    If P3StageOn(i) Then
      If P3StageKind(i) = "RESET_U" Or P3StageKind(i) = "RESET_STRESS" Or P3StageKind(i) = "MATSET" Then
        P3LastOriginIndex = i
      End If
    End If
  Next i
End Function

' Copy current active set, NodeCond, boundary disp and locked loads. Called after RESET_U / RESET_STRESS / MATSET.
Private Sub P3CaptureReplayOrigin()
  Dim i As Long, boundN As Long, boundF As Long, boundLsw As Long, boundLap As Long
  P3OriginSnapReady = False
  If NumberOfElement <= 0 Or lastDof < 0 Then Exit Sub
  If Not P3ElementActiveReady Then P3InitElementActive
  ReDim P3OriginElementActive(0 To NumberOfElement - 1)
  For i = 0 To NumberOfElement - 1
    P3OriginElementActive(i) = P3ElementActive(i)
  Next i
  ReDim P3OriginNodeCond(lastDof)
  ReDim P3OriginBoundaryDisp(lastDof)
  ReDim P3OriginAppliedForce(lastDof)
  ReDim P3OriginLockedSW(lastDof)
  ReDim P3OriginLockedAP(lastDof)
  boundN = -1: boundF = -1: boundLsw = -1: boundLap = -1
  On Error Resume Next
  boundN = UBound(P3BoundaryDisp)
  boundF = UBound(P3AppliedForce)
  boundLsw = UBound(P3LockedSelfWeight)
  boundLap = UBound(P3LockedApplied)
  Err.Clear
  On Error GoTo 0
  For i = 0 To lastDof
    P3OriginNodeCond(i) = NodeCond(i)
    If boundN = lastDof Then P3OriginBoundaryDisp(i) = P3BoundaryDisp(i)
    If boundF = lastDof Then P3OriginAppliedForce(i) = P3AppliedForce(i)
    If boundLsw = lastDof Then P3OriginLockedSW(i) = P3LockedSelfWeight(i)
    If boundLap = lastDof Then P3OriginLockedAP(i) = P3LockedApplied(i)
  Next i
  P3OriginHasGravity = P3PlanHasGravity
  P3OriginHasApply = P3PlanHasApply
  P3OriginSnapReady = True
End Sub

' Restore the captured origin, bump ActiveSetGen, then reassemble self-weight. No-op if OriginSnapReady is False.
Private Sub P3RestoreReplayOrigin()
  Dim i As Long
  If Not P3OriginSnapReady Then Exit Sub
  If NumberOfElement <= 0 Or lastDof < 0 Then Exit Sub
  If UBound(P3OriginElementActive) <> NumberOfElement - 1 Then Exit Sub
  If UBound(P3OriginNodeCond) <> lastDof Then Exit Sub
  ReDim P3ElementActive(0 To NumberOfElement - 1)
  For i = 0 To NumberOfElement - 1
    P3ElementActive(i) = P3OriginElementActive(i)
  Next i
  P3ElementActiveReady = True
  P3ActiveSetGen = P3ActiveSetGen + 1
  If P3ActiveSetGen <= 0 Then P3ActiveSetGen = 1
  P3LoadIndexReady = False
  P6ScatterReady = False
  ReDim P3LockedSelfWeight(lastDof)
  ReDim P3LockedApplied(lastDof)
  ReDim P3AppliedForce(lastDof)
  For i = 0 To lastDof
    NodeCond(i) = P3OriginNodeCond(i)
    P3BoundaryDisp(i) = P3OriginBoundaryDisp(i)
    P3AppliedForce(i) = P3OriginAppliedForce(i)
    P3LockedSelfWeight(i) = P3OriginLockedSW(i)
    P3LockedApplied(i) = P3OriginLockedAP(i)
  Next i
  P3PlanHasGravity = P3OriginHasGravity
  P3PlanHasApply = P3OriginHasApply
  P3HoldOrphanDofs
  P3AssembleSelfWeightOnly
End Sub

' Zero displacements and plastic history for an SRM trial. Must NOT reset CurrentStrengthFactor / FSS.
Private Sub P3ZeroTrialKinematics()
  Dim i As Long, k As Long
  If lastDof >= 0 Then
    ReDim P3CommittedDisp(lastDof)
    For i = 0 To lastDof
      TDisp(i) = 0#
      UDisp(i) = 0#
      Disp(i) = 0#
      P3CommittedDisp(i) = 0#
    Next i
    ReDim P3LockedSelfWeight(lastDof)
    ReDim P3LockedApplied(lastDof)
  End If
  P3GravityCommitted = False
  For k = 0 To NumberOfElement - 1
    P3ResetElementHistory k
  Next k
  P3HoldOrphanDofs
  P3AssembleSelfWeightOnly
End Sub

Private Sub P3ResetKinematicState(ByVal moveGeom As Boolean)
  Dim i As Long, j As Long, k As Long, nd As Long
  If moveGeom Then
    For k = 0 To NumberOfElement - 1
      For j = 0 To 7
        nd = Elem(k).node(j)
        If nd >= 0 Then
          Elem(k).x(j) = XXX(2 * nd) + TDisp(2 * nd)
          Elem(k).y(j) = XXX(2 * nd + 1) + TDisp(2 * nd + 1)
        End If
      Next j
    Next k
    If lastDof >= 0 Then
      For i = 0 To lastDof
        XXX(i) = XXX(i) + TDisp(i)
      Next i
    End If
    P6InvalidateTangentGeneration
    If Not SetElmMat() Then Err.Raise vbObjectError + 3241, "FEM.P3ResetKinematicState", "RESET_STRESSの剛性再作成に失敗しました。"
    SetTotalMat
  End If
  P3ZeroTrialKinematics
  If Not P3SetMaterialForStrengthFactor(1#) Then Err.Raise vbObjectError + 3242, "FEM.P3ResetKinematicState", "強度をFs=1へ戻せませんでした。"
  FSS = 1#
  P3RecountActivePlasticPoints
End Sub

Private Function P3KindIsJoint(ByVal kindText As String) As Boolean
  P3KindIsJoint = (UCase$(Trim$(kindText)) = "JOINT")
End Function

Private Function P3ParseStrictDouble(ByVal rawText As String, ByRef value As Double, ByRef failMessage As String) As Boolean
  Dim t As String
  P3ParseStrictDouble = False
  value = 0#
  t = Trim$(rawText)
  If Len(t) = 0 Then
    failMessage = "数値が空です。"
    Exit Function
  End If
  If Not IsNumeric(t) Then
    failMessage = "数値ではありません。値=" & t
    Exit Function
  End If
  On Error GoTo ParseFail
  value = CDbl(t)
  If Not P2IsFinite(value) Then
    failMessage = "数値が有限ではありません。値=" & t
    Exit Function
  End If
  P3ParseStrictDouble = True
  Exit Function
ParseFail:
  failMessage = "数値の変換に失敗しました。値=" & t
End Function

Public Function P3ValidateJointMaterial(ByRef mat As Material_Data, ByRef failMessage As String) As Boolean
  P3ValidateJointMaterial = False
  If mat.kn <= 0# Then
    failMessage = "ジョイントのKnは正で指定してください。"
    Exit Function
  End If
  If mat.ks <= 0# Then
    failMessage = "ジョイントのKsは正で指定してください。"
    Exit Function
  End If
  If mat.cohesion < 0# Then
    failMessage = "ジョイントの粘着力は0以上で指定してください。"
    Exit Function
  End If
  If mat.fai < 0# Then
    failMessage = "ジョイントの摩擦角は0度以上で指定してください。"
    Exit Function
  End If
  If mat.fai >= 89.999 Then
    failMessage = "ジョイントの摩擦角が90度以上または特異点に近すぎます。"
    Exit Function
  End If
  If mat.psai < 0# Then
    failMessage = "ジョイントのダイレタンシー角は0度以上で指定してください。"
    Exit Function
  End If
  If mat.psai >= 89.999 Then
    failMessage = "ジョイントのダイレタンシー角が90度以上または特異点に近すぎます。"
    Exit Function
  End If
  If mat.thickness <= 0# Then
    failMessage = "ジョイントの厚さは正で指定してください。"
    Exit Function
  End If
  If mat.weight < 0# Then
    failMessage = "ジョイントの単位体積重量は0以上で指定してください。"
    Exit Function
  End If
  P3ValidateJointMaterial = True
End Function

Public Function P3ValidateSolidMaterial(ByRef mat As Material_Data, ByRef failMessage As String) As Boolean
  P3ValidateSolidMaterial = False
  If Not P2MaterialConstantsAreValid(mat.Young, mat.Poisson, mat.fai, mat.cohesion, mat.psai, P2_DEFAULT_TOLERANCE, failMessage) Then Exit Function
  If mat.thickness <= 0# Then
    failMessage = "板厚は正で指定してください。"
    Exit Function
  End If
  If mat.weight < 0# Then
    failMessage = "単位体積重量は0以上で指定してください。"
    Exit Function
  End If
  P3ValidateSolidMaterial = True
End Function

Private Function P3ValidateAssignedMaterial(ByRef mat As Material_Data, ByRef failMessage As String) As Boolean
  If P3KindIsJoint(mat.kind) Then
    P3ValidateAssignedMaterial = P3ValidateJointMaterial(mat, failMessage)
  Else
    P3ValidateAssignedMaterial = P3ValidateSolidMaterial(mat, failMessage)
  End If
End Function

Private Function P3ParseMatSetToken(ByVal rawText As String, ByVal keyName As String, ByRef found As Boolean, ByRef value As Double, ByRef failMessage As String) As Boolean
  Dim parts() As String, i As Long, token As String, p As Long, lhs As String
  found = False
  value = 0#
  P3ParseMatSetToken = True
  If Len(rawText) = 0 Then Exit Function
  parts = Split(rawText, ";")
  For i = LBound(parts) To UBound(parts)
    token = Trim$(parts(i))
    p = InStr(1, token, "=")
    If p > 1 Then
      lhs = LCase$(Trim$(left$(token, p - 1)))
      If lhs = LCase$(keyName) Then
        found = True
        If Not P3ParseStrictDouble(mid$(token, p + 1), value, failMessage) Then
          failMessage = keyName & "=" & failMessage
          P3ParseMatSetToken = False
        End If
        Exit Function
      End If
    End If
  Next i
End Function

Private Function P3ApplyMatSet(ByVal stageIndex As Long) As Boolean
  Dim targetId As Long, sourceId As Long, idx As Long, srcIdx As Long
  Dim paramText As String, found As Boolean, value As Double, piValue As Double
  Dim shearModulus As Double, lameLambda As Double
  Dim matNote As String, failMessage As String
  Dim destIsJoint As Boolean, srcIsJoint As Boolean
  P3ApplyMatSet = False
  matNote = ""
  P3StateNote = ""
  targetId = P3StageMaterialId(stageIndex)
  If targetId < 1 Or targetId > NumberOfMaterial Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "MATSETの材料番号が範囲外です。ステージ=" & CStr(P3StageId(stageIndex)), vbObjectError + 3243, -1, -1, CurrentIncrement, CurrentIteration
    Exit Function
  End If
  idx = targetId - 1
  paramText = Trim$(P3StageParam(stageIndex))
  If Len(paramText) = 0 Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "MATSETのパラメータが空です。コピー元材料番号または c=;phi= を指定してください。", vbObjectError + 3244, -1, -1, CurrentIncrement, CurrentIteration
    Exit Function
  End If
  If IsNumeric(paramText) Then
    If Not P3ParseStrictDouble(paramText, value, failMessage) Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSETのコピー元材料が数値ではありません。 " & failMessage, vbObjectError + 3245, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    If value <> CLng(value) Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSETのコピー元は整数の材料番号で指定してください。値=" & CStr(value), vbObjectError + 3245, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    sourceId = CLng(value)
    If sourceId < 1 Or sourceId > NumberOfMaterial Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSETのコピー元材料が範囲外です。値=" & CStr(sourceId), vbObjectError + 3245, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    srcIdx = sourceId - 1
    destIsJoint = P3KindIsJoint(Material(idx).kind)
    srcIsJoint = P3KindIsJoint(P3SheetMaterial(srcIdx).kind)
    If destIsJoint Xor srcIsJoint Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSETはJOINTと非JOINTを相互にコピーできません。先=" & CStr(targetId) & " 元=" & CStr(sourceId), vbObjectError + 3247, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    Material(idx) = P3SheetMaterial(srcIdx)
  Else
    If Not P3ParseMatSetToken(paramText, "c", found, value, failMessage) Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    If found Then Material(idx).cohesion = value
    If Not P3ParseMatSetToken(paramText, "phi", found, value, failMessage) Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    If Not found Then
      If Not P3ParseMatSetToken(paramText, "fai", found, value, failMessage) Then
        SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
        Exit Function
      End If
    End If
    If found Then Material(idx).fai = value
    If Not P3ParseMatSetToken(paramText, "psi", found, value, failMessage) Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    If Not found Then
      If Not P3ParseMatSetToken(paramText, "psai", found, value, failMessage) Then
        SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
        Exit Function
      End If
    End If
    If found Then Material(idx).psai = value
    If Not P3ParseMatSetToken(paramText, "E", found, value, failMessage) Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    If Not found Then
      If Not P3ParseMatSetToken(paramText, "young", found, value, failMessage) Then
        SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
        Exit Function
      End If
    End If
    If found Then
      If P3KindIsJoint(Material(idx).kind) Then
        matNote = "JOINTにE=は無視。kn=を使ってください。"
      Else
        Material(idx).Young = value
      End If
    End If
    If Not P3ParseMatSetToken(paramText, "nu", found, value, failMessage) Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    If Not found Then
      If Not P3ParseMatSetToken(paramText, "poisson", found, value, failMessage) Then
        SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
        Exit Function
      End If
    End If
    If found Then Material(idx).Poisson = value
    If Not P3ParseMatSetToken(paramText, "gamma", found, value, failMessage) Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    If Not found Then
      If Not P3ParseMatSetToken(paramText, "weight", found, value, failMessage) Then
        SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
        Exit Function
      End If
    End If
    If found Then Material(idx).weight = value
    If Not P3ParseMatSetToken(paramText, "ks", found, value, failMessage) Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    If found Then Material(idx).ks = value
    If Not P3ParseMatSetToken(paramText, "kn", found, value, failMessage) Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "MATSET: " & failMessage, vbObjectError + 3248, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    If found Then
      If P3KindIsJoint(Material(idx).kind) Then
        Material(idx).kn = value
      Else
        matNote = "固体にkn=は無視。E=を使ってください。"
      End If
    End If
  End If
  If P3KindIsJoint(Material(idx).kind) Then
    If Material(idx).ks <= 0# Then Material(idx).ks = 0.1 * Material(idx).kn
  End If
  If Not P3ValidateAssignedMaterial(Material(idx), failMessage) Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "MATSETの材料定数が範囲外です。材料=" & CStr(targetId) & " " & failMessage, vbObjectError + 3249, -1, -1, CurrentIncrement, CurrentIteration
    Exit Function
  End If
  piValue = 3.14159265358979
  If P3KindIsJoint(Material(idx).kind) Then
    Material(idx).ElasticD00 = Material(idx).kn
    Material(idx).ElasticD01 = 0#
    Material(idx).ElasticD22 = Material(idx).ks
  Else
    shearModulus = Material(idx).Young / (2# * (1# + Material(idx).Poisson))
    lameLambda = Material(idx).Young * Material(idx).Poisson / ((1# + Material(idx).Poisson) * (1# - 2# * Material(idx).Poisson))
    Material(idx).ElasticD00 = lameLambda + 2# * shearModulus
    Material(idx).ElasticD01 = lameLambda
    Material(idx).ElasticD22 = shearModulus
  End If
  Material(idx).SinFriction = Sin(Material(idx).fai * piValue / 180#)
  Material(idx).CosFriction = Cos(Material(idx).fai * piValue / 180#)
  Material(idx).SinDilation = Sin(Material(idx).psai * piValue / 180#)
  Material(idx).CosDilation = Cos(Material(idx).psai * piValue / 180#)
  Material(idx).MaterialCacheReady = True
  P3OriginalMaterial(idx) = Material(idx)
  P6InvalidateTangentGeneration
  If Not SetElmMat() Then
    SetAnalysisFailure RESULT_MATERIAL_ERROR, "MATSET後の剛性再作成に失敗しました。", vbObjectError + 3246, -1, -1, CurrentIncrement, CurrentIteration
    Exit Function
  End If
  SetTotalMat
  P3AssembleSelfWeightOnly
  P3RecountActivePlasticPoints
  If Len(matNote) > 0 Then P3StateNote = matNote
  P3ApplyMatSet = True
End Function

Public Sub P3JointBuild(ByVal elementId As Long)
  P3JointEval elementId, False
End Sub

Public Sub P3JointEval(ByVal elementId As Long, ByVal addForce As Boolean)
  Dim a As Long, i As Long, j As Long, gp As Long, materialIndex As Long
  Dim knValue As Double, ksValue As Double, jac As Double, dVolume As Double
  Dim dXdxi As Double, dYdxi As Double, tx As Double, ty As Double, nxVal As Double, nyVal As Double
  Dim gapValue As Double, slipValue As Double, fnValue As Double, fsValue As Double
  Dim tauMax As Double, dnn As Double, dss As Double, scaleValue As Double
  Dim umx As Double, umy As Double, usx As Double, usy As Double
  Dim xiValue As Double, n0 As Double, n1 As Double, n2 As Double
  Dim dn0 As Double, dn1 As Double, dn2 As Double
  Dim shapeN(0 To 2) As Double, dN(0 To 2) As Double
  Dim bn(0 To 11) As Double, bs(0 To 11) As Double
  Dim xiG(0 To 2) As Double, wG(0 To 2) As Double
  Dim dofId As Long, piValue As Double, tanPhi As Double
  If elementId < 0 Or elementId > NumberOfElement - 1 Then Exit Sub
  materialIndex = Elem(elementId).MatNo
  If materialIndex < 0 Or materialIndex > NumberOfMaterial - 1 Then Exit Sub
  Elem(elementId).IsJoint = True
  For i = 0 To 5
    If Elem(elementId).node(i) < 0 Then Exit Sub
    Elem(elementId).ElNode(2 * i) = 2 * Elem(elementId).node(i)
    Elem(elementId).ElNode(2 * i + 1) = 2 * Elem(elementId).node(i) + 1
  Next i
  For i = 12 To 15
    Elem(elementId).ElNode(i) = -1
  Next i
  For i = 0 To 15
    For j = 0 To 15
      Elem(elementId).kmat(i, j) = 0#
    Next j
  Next i
  Elem(elementId).ElmWeight = 0#
  Elem(elementId).HourglassModeCount = 0
  Elem(elementId).HourglassScale = 0#
  knValue = Material(materialIndex).kn
  If knValue <= 0# Then knValue = Material(materialIndex).Young
  ksValue = Material(materialIndex).ks
  If ksValue <= 0# Then ksValue = 0.1 * knValue
  piValue = 3.14159265358979
  tanPhi = Tan(Material(materialIndex).fai * piValue / 180#)
  xiG(0) = -0.774596669241483: xiG(1) = 0#: xiG(2) = 0.774596669241483
  wG(0) = 0.555555555555556: wG(1) = 0.888888888888889: wG(2) = 0.555555555555556
  For gp = 0 To 2
    xiValue = xiG(gp)
    shapeN(0) = 0.5 * xiValue * (xiValue - 1#)
    shapeN(1) = 0.5 * xiValue * (xiValue + 1#)
    shapeN(2) = 1# - xiValue * xiValue
    dN(0) = xiValue - 0.5
    dN(1) = xiValue + 0.5
    dN(2) = -2# * xiValue
    dXdxi = 0#: dYdxi = 0#
    For a = 0 To 2
      dXdxi = dXdxi + dN(a) * Elem(elementId).x(a)
      dYdxi = dYdxi + dN(a) * Elem(elementId).y(a)
    Next a
    jac = Sqr(dXdxi * dXdxi + dYdxi * dYdxi)
    If jac < 0.000000000001 Then
      Elem(elementId).dj(gp) = 0#
      Elem(elementId).JointContact(gp) = 1
      GoTo NextJointGp
    End If
    tx = dXdxi / jac
    ty = dYdxi / jac
    nxVal = ty
    nyVal = -tx
    gapValue = 0#: slipValue = 0#
    For a = 0 To 2
      umx = TDisp(Elem(elementId).ElNode(2 * a))
      umy = TDisp(Elem(elementId).ElNode(2 * a + 1))
      usx = TDisp(Elem(elementId).ElNode(6 + 2 * a))
      usy = TDisp(Elem(elementId).ElNode(7 + 2 * a))
      gapValue = gapValue + shapeN(a) * ((usx - umx) * nxVal + (usy - umy) * nyVal)
      slipValue = slipValue + shapeN(a) * ((usx - umx) * tx + (usy - umy) * ty)
    Next a
    For i = 0 To 11
      bn(i) = 0#: bs(i) = 0#
    Next i
    For a = 0 To 2
      bn(2 * a) = -shapeN(a) * nxVal
      bn(2 * a + 1) = -shapeN(a) * nyVal
      bs(2 * a) = -shapeN(a) * tx
      bs(2 * a + 1) = -shapeN(a) * ty
      bn(6 + 2 * a) = shapeN(a) * nxVal
      bn(6 + 2 * a + 1) = shapeN(a) * nyVal
      bs(6 + 2 * a) = shapeN(a) * tx
      bs(6 + 2 * a + 1) = shapeN(a) * ty
    Next a
    dnn = knValue
    dss = ksValue
    fnValue = knValue * gapValue
    fsValue = ksValue * slipValue
    ' gap>0 はマスター外向き法線に対する開口（引張）。AllowTension=False なら剥離。
    ' fn=kn*gap なので圧縮は fn<0。ここを引張側と取り違えると接合が常にゼロ剛性になる。
    If (Not Material(materialIndex).AllowTension) And (gapValue > 0.000000000001) Then
      dnn = 0#: dss = 0#: fnValue = 0#: fsValue = 0#
      Elem(elementId).JointContact(gp) = 1
    Else
      tauMax = Material(materialIndex).cohesion
      If fnValue < 0# Then tauMax = tauMax + (-fnValue) * tanPhi
      If tauMax < 0# Then tauMax = 0#
      If Abs(fsValue) > tauMax Then
        If fsValue >= 0# Then fsValue = tauMax Else fsValue = -tauMax
        dss = 0#
        Elem(elementId).JointContact(gp) = 2
      Else
        Elem(elementId).JointContact(gp) = 0
      End If
    End If
    dVolume = wG(gp) * jac * Material(materialIndex).thickness
    If dVolume < 0# Then dVolume = -dVolume
    Elem(elementId).dj(gp) = dVolume
    Elem(elementId).Stmat(0, gp) = -fnValue
    Elem(elementId).Stmat(1, gp) = fsValue
    Elem(elementId).Stmat(2, gp) = gapValue
    Elem(elementId).Stmat(10, gp) = slipValue
    For i = 0 To 11
      For j = 0 To 11
        Elem(elementId).kmat(i, j) = Elem(elementId).kmat(i, j) + dVolume * (dnn * bn(i) * bn(j) + dss * bs(i) * bs(j))
      Next j
    Next i
    If addForce Then
      For i = 0 To 11
        dofId = Elem(elementId).ElNode(i)
        If dofId >= 0 And dofId <= lastDof Then
          iNForce(dofId) = iNForce(dofId) + dVolume * (fnValue * bn(i) + fsValue * bs(i))
        End If
      Next i
    End If
NextJointGp:
  Next gp
End Sub

Private Sub P3ClearTransientFailure()
  AnalysisOK = False
  ResultStatus = RESULT_NOT_RUN
  AnalysisMessage = vbNullString
  AnalysisErrorNumber = 0
  FailureElement = -1
  FailureGaussPoint = -1
  FailureIncrement = -1
  FailureIteration = -1
  ResultRevision = 0
  P3CurrentElementZeroIncrement = False
End Sub

Private Function P3IsFatalFailure() As Boolean
  If P3UserCancel Then
    P3IsFatalFailure = True
    Exit Function
  End If
  If P3SearchFatal Then
    P3IsFatalFailure = True
    Exit Function
  End If
  If AnalysisErrorNumber = 18 Then
    P3IsFatalFailure = True
    Exit Function
  End If
  If ResultStatus = RESULT_INPUT_ERROR Or ResultStatus = RESULT_CAPACITY_ERROR Or ResultStatus = RESULT_RUNTIME_ERROR Then
    P3IsFatalFailure = True
    Exit Function
  End If
  If ResultStatus = RESULT_MATERIAL_ERROR Then
    P3IsFatalFailure = True
  End If
End Function

Private Function P3HasLaterEnabledSrm(ByVal afterIndex As Long) As Boolean
  Dim i As Long
  P3HasLaterEnabledSrm = False
  For i = afterIndex + 1 To P3StageN
    If P3StageOn(i) Then
      If P3StageKind(i) = "SRM" Then
        P3HasLaterEnabledSrm = True
        Exit Function
      End If
    End If
  Next i
End Function

Private Function P3CanDeferGravityToSrm(ByVal stageIndex As Long) As Boolean
  P3CanDeferGravityToSrm = False
  If P3ReplayQuiet Then Exit Function
  If P3SrmTrialRunning Then Exit Function
  If P3UserCancel Or P3SearchFatal Then Exit Function
  If AnalysisErrorNumber = 18 Then Exit Function
  If ResultStatus <> RESULT_NONCONVERGED And ResultStatus <> RESULT_GLOBAL_SINGULAR Then Exit Function
  If Not P3HasLaterEnabledSrm(stageIndex) Then Exit Function
  P3CanDeferGravityToSrm = True
End Function

Private Sub P3TraceIteration(ByVal eventName As String)
  If femIoMode = "ITER" Then FEMIoLog eventName, "ITER", "giter=" & CStr(P3GlobalIterationCount)
End Sub

Private Sub P3NoteSolveAge()
  If P3TangentJustRebuilt Then
    P3TangentAge = 0
    P3TangentJustRebuilt = False
    P6TanSolve0 = P6TanSolve0 + 1
  Else
    P3TangentAge = P3TangentAge + 1
    If P3TangentAge = 1 Then
      P6TanSolve1 = P6TanSolve1 + 1
    ElseIf P3TangentAge = 2 Then
      P6TanSolve2 = P6TanSolve2 + 1
    Else
      P6TanSolve3 = P6TanSolve3 + 1
    End If
  End If
End Sub

Private Sub P3PerfTrace(ByVal eventName As String, ByVal stepSize As Double, ByVal consecutiveCutback As Long, ByVal duMax As Double, ByVal lambdaNow As Double, ByVal lambdaTarget As Double)
  Dim stageNo As Long
  Dim kindText As String
  ' During an SRM trial the inner replay is GRAVITY/APPLY. Log the outer SRM stage.
  If P3SrmTrialRunning And P3SrmLogStageNo > 0 Then
    stageNo = P3SrmLogStageNo
    kindText = P3SrmLogKind
    If Len(kindText) = 0 Then kindText = "SRM"
  Else
    stageNo = P3RunLogStageNo
    kindText = P3RunLogKind
  End If
  P6PerfTrace eventName, stepSize, consecutiveCutback, duMax, lambdaNow, lambdaTarget, stageNo, kindText
  If eventName = "CUTBACK" Then P6SolverEvent "CUTBACK", "step=" & Format$(stepSize, "0.000E+00")
End Sub

Private Sub P3CaptureOriginalMaterial()
  Dim i As Long
  ReDim P3OriginalMaterial(NumberOfMaterial - 1)
  ReDim P3SheetMaterial(NumberOfMaterial - 1)
  For i = 0 To NumberOfMaterial - 1
    P3SheetMaterial(i) = Material(i)
    P3OriginalMaterial(i) = Material(i)
  Next i
End Sub

Public Function P3SetMaterialForStrengthFactor(ByVal strengthFactor As Double) As Boolean
  Dim i As Long, piValue As Double
  Dim phiDeg As Double, psiDeg As Double, origPhi As Double, origPsi As Double
  Dim constFail As String
  P3SetMaterialForStrengthFactor = False
  If strengthFactor <= 0# Then
    SetAnalysisFailure RESULT_INPUT_ERROR, "P3の強度低減係数は正で指定してください。", vbObjectError + 3210, -1, -1, CurrentIncrement, CurrentIteration
    Exit Function
  End If
  piValue = 3.14159265358979
  P3EnsurePolicyCache
  For i = 0 To NumberOfMaterial - 1
    Material(i).Young = P3OriginalMaterial(i).Young
    Material(i).Poisson = P3OriginalMaterial(i).Poisson
    Material(i).thickness = P3OriginalMaterial(i).thickness
    Material(i).weight = P3OriginalMaterial(i).weight
    Material(i).kind = P3OriginalMaterial(i).kind
    Material(i).kn = P3OriginalMaterial(i).kn
    Material(i).ks = P3OriginalMaterial(i).ks
    Material(i).AllowTension = P3OriginalMaterial(i).AllowTension
    Material(i).ElasticD00 = P3OriginalMaterial(i).ElasticD00
    Material(i).ElasticD01 = P3OriginalMaterial(i).ElasticD01
    Material(i).ElasticD22 = P3OriginalMaterial(i).ElasticD22
    origPhi = P3OriginalMaterial(i).fai
    origPsi = P3OriginalMaterial(i).psai
    If Not P3OriginalMaterial(i).ReduceStrength Then
      phiDeg = origPhi
      psiDeg = origPsi
      Material(i).cohesion = P3OriginalMaterial(i).cohesion
    Else
      phiDeg = Atn(Tan(origPhi * piValue / 180#) / strengthFactor) * 180# / piValue
      If P3SrmPsiMode = 1 Then
        psiDeg = Atn(Tan(origPsi * piValue / 180#) / strengthFactor) * 180# / piValue
      Else
        psiDeg = origPsi
      End If
      If P3FlowPolicyMode <> 1 And P3SrmPsiMode <> 2 Then
        If psiDeg > phiDeg Then psiDeg = phiDeg
      End If
      Material(i).cohesion = P3OriginalMaterial(i).cohesion / strengthFactor
    End If
    ' Davis equivalent c*,phi* is for continuum MC only. JOINT keeps its own Coulomb c,phi.
    If P3FlowPolicyIsDavis() And (Not P3KindIsJoint(Material(i).kind)) Then P3ApplyDavisEquivalent phiDeg, psiDeg, Material(i).cohesion
    Material(i).fai = phiDeg
    Material(i).psai = psiDeg
    Material(i).ReduceStrength = P3OriginalMaterial(i).ReduceStrength
    Material(i).InitialActive = P3OriginalMaterial(i).InitialActive
    Material(i).SinFriction = Sin(Material(i).fai * piValue / 180#)
    Material(i).CosFriction = Cos(Material(i).fai * piValue / 180#)
    Material(i).SinDilation = Sin(Material(i).psai * piValue / 180#)
    Material(i).CosDilation = Cos(Material(i).psai * piValue / 180#)
    If Material(i).kind = "JOINT" Then
      Material(i).ElasticD00 = Material(i).kn
      Material(i).ElasticD01 = 0#
      Material(i).ElasticD22 = Material(i).ks
    End If
    Material(i).MaterialCacheReady = True
    If Material(i).kind <> "JOINT" Then
      If Not P2MaterialConstantsAreValid(Material(i).Young, Material(i).Poisson, Material(i).fai, Material(i).cohesion, Material(i).psai, P2_DEFAULT_TOLERANCE, constFail) Then
        SetAnalysisFailure RESULT_INPUT_ERROR, "強度低減後の材料定数が範囲外です。材料=" & CStr(i + 1), vbObjectError + 3210, -1, -1, CurrentIncrement, CurrentIteration
        Exit Function
      End If
    End If
   Next i
   If Abs(strengthFactor - P3CurrentStrengthFactor) > 0.000000000001 Then P6InvalidateStrengthGeneration
   P3CurrentStrengthFactor = strengthFactor
   P3SetMaterialForStrengthFactor = True
End Function

Public Sub P3EnsurePolicyCache()
  Dim flowText As String, psiText As String, weightText As String
  If P3PolicyCacheReady Then Exit Sub
  flowText = UCase$(Trim$(P6ReadTextSetting("FLOW_POLICY", "INCONSISTENT")))
  If flowText = "DAVIS" Then
    P3FlowPolicyMode = 1
  Else
    P3FlowPolicyMode = 0
  End If
  psiText = UCase$(Trim$(P6ReadTextSetting("SRM_PSI_POLICY", "CAP")))
  If psiText = "REDUCE" Then
    P3SrmPsiMode = 1
  ElseIf psiText = "KEEP" Then
    P3SrmPsiMode = 2
  Else
    P3SrmPsiMode = 0
  End If
  weightText = UCase$(Trim$(P6ReadTextSetting("WEIGHT_MODE", "TOTAL")))
  If weightText = "TOTAL" Then
    P3WeightMode = 0
  Else
    P3WeightMode = -1
  End If
  P3PolicyCacheReady = True
End Sub

Public Function P3FlowPolicyText() As String
  P3EnsurePolicyCache
  If P3FlowPolicyMode = 1 Then
    P3FlowPolicyText = "DAVIS"
  Else
    P3FlowPolicyText = "INCONSISTENT"
  End If
End Function

Private Function P3FlowPolicyIsDavis() As Boolean
  P3EnsurePolicyCache
  P3FlowPolicyIsDavis = (P3FlowPolicyMode = 1)
End Function

Public Function P3FlowPolicyIsInconsistent() As Boolean
  P3EnsurePolicyCache
  P3FlowPolicyIsInconsistent = (P3FlowPolicyMode <> 1)
End Function

Private Sub P3ApplyDavisEquivalent(ByRef phiDeg As Double, ByRef psiDeg As Double, ByRef cohesion As Double)
  Dim piValue As Double, phiRad As Double, psiRad As Double
  Dim denom As Double, ratio As Double, tanPhi As Double
  If Abs(phiDeg - psiDeg) <= 0.000000000001 Then Exit Sub
  piValue = 3.14159265358979
  phiRad = phiDeg * piValue / 180#
  psiRad = psiDeg * piValue / 180#
  denom = 1# - Sin(phiRad) * Sin(psiRad)
  If denom <= 0.000000000001 Then
    psiDeg = phiDeg
    Exit Sub
  End If
  ratio = Cos(psiRad) * Cos(phiRad) / denom
  If ratio < 0# Then ratio = 0#
  cohesion = cohesion * ratio
  tanPhi = Tan(phiRad)
  phiDeg = Atn(tanPhi * ratio) * 180# / piValue
  psiDeg = phiDeg
End Sub

Private Sub P3SymmetrizeKmat(ByRef kmat() As Double)
  Dim i As Long, j As Long, avg As Double
  For i = 0 To 15
    For j = i + 1 To 15
      avg = 0.5 * (kmat(i, j) + kmat(j, i))
      kmat(i, j) = avg
      kmat(j, i) = avg
    Next j
  Next i
End Sub

Public Sub P3EnsureElementCounterClockwise(ByRef elementData As Element_Data)
  Dim areaValue As Double
  Dim swapNode As Long, swapX As Double, swapY As Double
  areaValue = (elementData.x(0) * elementData.y(1) - elementData.x(1) * elementData.y(0)) _
            + (elementData.x(1) * elementData.y(2) - elementData.x(2) * elementData.y(1)) _
            + (elementData.x(2) * elementData.y(3) - elementData.x(3) * elementData.y(2)) _
            + (elementData.x(3) * elementData.y(0) - elementData.x(0) * elementData.y(3))
  If areaValue >= 0# Then Exit Sub
  swapNode = elementData.node(1): elementData.node(1) = elementData.node(3): elementData.node(3) = swapNode
  swapX = elementData.x(1): elementData.x(1) = elementData.x(3): elementData.x(3) = swapX
  swapY = elementData.y(1): elementData.y(1) = elementData.y(3): elementData.y(3) = swapY
  swapNode = elementData.node(4): elementData.node(4) = elementData.node(7): elementData.node(7) = swapNode
  swapX = elementData.x(4): elementData.x(4) = elementData.x(7): elementData.x(7) = swapX
  swapY = elementData.y(4): elementData.y(4) = elementData.y(7): elementData.y(7) = swapY
  swapNode = elementData.node(5): elementData.node(5) = elementData.node(6): elementData.node(6) = swapNode
  swapX = elementData.x(5): elementData.x(5) = elementData.x(6): elementData.x(6) = swapX
  swapY = elementData.y(5): elementData.y(5) = elementData.y(6): elementData.y(6) = swapY
End Sub

Public Function P3IsElementActive(ByVal elementId As Long) As Boolean
  If Not P3ElementActiveReady Then
    P3IsElementActive = True
    Exit Function
  End If
  If elementId < LBound(P3ElementActive) Or elementId > UBound(P3ElementActive) Then
    P3IsElementActive = True
    Exit Function
  End If
  P3IsElementActive = P3ElementActive(elementId)
End Function

Private Sub P3InitElementActive()
  Dim i As Long, materialIndex As Long
  If NumberOfElement <= 0 Then
    P3ElementActiveReady = False
    Exit Sub
  End If
  ReDim P3ElementActive(0 To NumberOfElement - 1)
  For i = 0 To NumberOfElement - 1
    materialIndex = Elem(i).MatNo
    If materialIndex >= 0 And materialIndex <= UBound(Material) Then
      P3ElementActive(i) = Material(materialIndex).InitialActive
    Else
      P3ElementActive(i) = True
    End If
  Next i
  P3ElementActiveReady = True
  P3ActiveSetGen = P3ActiveSetGen + 1
  If P3ActiveSetGen <= 0 Then P3ActiveSetGen = 1
  P3LoadIndexReady = False
  P6ScatterReady = False
End Sub

' Live BIRTH zeros history. Resume skip must pass resetHistoryOnBirth:=False so gauss.csv is kept.
Private Sub P3SetMaterialElementsActive(ByVal materialId As Long, ByVal isActive As Boolean, Optional ByVal resetHistoryOnBirth As Boolean = True)
  Dim i As Long
  If materialId <= 0 Or NumberOfElement <= 0 Then Exit Sub
  If Not P3ElementActiveReady Then P3InitElementActive
  For i = 0 To NumberOfElement - 1
    If Elem(i).MatNo = materialId - 1 Then
      If isActive And Not P3ElementActive(i) Then
        If resetHistoryOnBirth Then P3ResetElementHistory i
      End If
      P3ElementActive(i) = isActive
    End If
  Next i
  P3RecountActivePlasticPoints
End Sub

Private Sub P3ResetElementHistory(ByVal elementId As Long)
  Dim i As Long, j As Long, im As Long
  If elementId < 0 Or elementId > NumberOfElement - 1 Then Exit Sub
  For im = 0 To 3
    For i = 0 To 16
      Elem(elementId).Stmat(i, im) = 0#
      Elem(elementId).mStmat(i, im) = 0#
    Next i
    For i = 0 To 3
      Elem(elementId).P2PlasticStrain(i, im) = 0#
      P3CommittedPlasticStrain(i, im, elementId) = 0#
    Next i
    Elem(elementId).P2PlasticMultiplier(im) = 0#
    Elem(elementId).P2YieldFunction(im) = 0#
    Elem(elementId).P2Yielded(im) = False
    Elem(elementId).P2LastSuccessfulSubsteps(im) = 1
    Elem(elementId).P2EasySubstepStreak(im) = 0
    P3CommittedPlasticMultiplier(im, elementId) = 0#
    P3CommittedYieldFunction(im, elementId) = 0#
    P3CommittedYielded(im, elementId) = False
    For j = 0 To 2
      Elem(elementId).P2PrincipalStress(j, im) = 0#
      P3CommittedPrincipalStress(j, im, elementId) = 0#
    Next j
  Next im
  P3ResetElementElasticTangent elementId
  On Error Resume Next
  If UBound(P3CommittedSpmatValid, 2) >= elementId Then
    For im = 0 To 3
      P3CommittedSpmatValid(im, elementId) = Elem(elementId).SpmatValid(im)
      For i = 0 To 2
        For j = 0 To 15
          P3CommittedSpmat(i, j, im, elementId) = Elem(elementId).Spmat(i, j, im)
        Next j
        For j = 0 To 2
          P3CommittedDpmat(i, j, im, elementId) = Elem(elementId).Dpmat(i, j, im)
        Next j
      Next i
    Next im
  End If
  Err.Clear
  PlasticLogResetElement elementId
End Sub

Private Sub P3EnsureMaterialElasticCache(ByVal materialIndex As Long)
  Dim shearModulus As Double, lameLambda As Double
  If materialIndex < 0 Or materialIndex > NumberOfMaterial - 1 Then Exit Sub
  If Material(materialIndex).MaterialCacheReady Then Exit Sub
  If Material(materialIndex).kind = "JOINT" Then
    Material(materialIndex).ElasticD00 = Material(materialIndex).kn
    Material(materialIndex).ElasticD01 = 0#
    Material(materialIndex).ElasticD22 = Material(materialIndex).ks
    Material(materialIndex).MaterialCacheReady = True
    Exit Sub
  End If
  shearModulus = Material(materialIndex).Young / (2# * (1# + Material(materialIndex).Poisson))
  lameLambda = Material(materialIndex).Young * Material(materialIndex).Poisson / ((1# + Material(materialIndex).Poisson) * (1# - 2# * Material(materialIndex).Poisson))
  Material(materialIndex).ElasticD00 = lameLambda + 2# * shearModulus
  Material(materialIndex).ElasticD01 = lameLambda
  Material(materialIndex).ElasticD22 = shearModulus
  Material(materialIndex).MaterialCacheReady = True
End Sub

Private Sub P3ResetElementElasticTangent(ByVal elementId As Long)
  Dim i As Long, j As Long, im As Long, materialIndex As Long
  If elementId < 0 Or elementId > NumberOfElement - 1 Then Exit Sub
  If Elem(elementId).IsJoint Then
    P3JointBuild elementId
    Exit Sub
  End If
  materialIndex = Elem(elementId).MatNo
  If materialIndex < 0 Or materialIndex > NumberOfMaterial - 1 Then Exit Sub
  P3EnsureMaterialElasticCache materialIndex
  Elem(elementId).Dmat(0, 0) = Material(materialIndex).ElasticD00
  Elem(elementId).Dmat(0, 1) = Material(materialIndex).ElasticD01
  Elem(elementId).Dmat(0, 2) = 0#
  Elem(elementId).Dmat(1, 0) = Material(materialIndex).ElasticD01
  Elem(elementId).Dmat(1, 1) = Material(materialIndex).ElasticD00
  Elem(elementId).Dmat(1, 2) = 0#
  Elem(elementId).Dmat(2, 0) = 0#
  Elem(elementId).Dmat(2, 1) = 0#
  Elem(elementId).Dmat(2, 2) = Material(materialIndex).ElasticD22
  For im = 0 To 3
    For i = 0 To 2
      For j = 0 To 2
        Elem(elementId).Dpmat(i, j, im) = Elem(elementId).Dmat(i, j)
      Next j
    Next i
  Next im
  SetSmat Elem(elementId).Dmat, Elem(elementId).Bmat, Elem(elementId).Smat
  SetSmat Elem(elementId).Dmat, Elem(elementId).Bmat, Elem(elementId).Spmat
  For im = 0 To 3
    Elem(elementId).SpmatValid(im) = True
  Next im
  P6EnsureGaussCache
  SetElmStiffness Material(materialIndex).thickness, Material(materialIndex).weight, Elem(elementId).Bmat, Elem(elementId).Smat, Elem(elementId).kmat, Elem(elementId).ElmWeight, Elem(elementId).dj, P6GaussN
  P6AddQ8HourglassStabilization Elem(elementId), Material(materialIndex).thickness, True
  Elem(elementId).TangentDirty = False
End Sub

Private Sub P3ApplyBirthDeathUpTo(ByVal lastIndex As Long)
  Dim i As Long, materialId As Long
  Dim lastI As Long
  lastI = lastIndex
  If lastI > P3StageN Then lastI = P3StageN
  For i = 1 To lastI
    If Not P3StageOn(i) Then GoTo NextBirthDeath
    If P3StageKind(i) = "BIRTH" Then
      materialId = 0
      If IsNumeric(P3StageGroup(i)) Then materialId = CLng(Val(P3StageGroup(i)))
      If materialId > 0 Then P3SetMaterialElementsActive materialId, True
    ElseIf P3StageKind(i) = "DEATH" Then
      materialId = 0
      If IsNumeric(P3StageGroup(i)) Then materialId = CLng(Val(P3StageGroup(i)))
      If materialId > 0 Then P3SetMaterialElementsActive materialId, False
    End If
NextBirthDeath:
  Next i
  P3ActiveSetGen = P3ActiveSetGen + 1
  If P3ActiveSetGen <= 0 Then P3ActiveSetGen = 1
  P3LoadIndexReady = False
  P6ScatterReady = False
End Sub

Public Sub P3AssembleSelfWeightOnly()
  Dim i As Long, j As Long, a As Long, b As Long, materialIndex As Long, dofId As Long
  Dim thickness As Double, gammaValue As Double, detJ As Double, dVolume As Double
  Dim dXdxi As Double, dYdxi As Double, dXdeta As Double, dYdeta As Double
  Dim xiValue As Double, etaValue As Double, shapeValue As Double
  Dim gaussXi(0 To 1) As Double, gaussW(0 To 1) As Double
  Dim shapeN(0 To 7) As Double, dNxi(0 To 7) As Double, dNeta(0 To 7) As Double
  gaussXi(0) = -0.5773502692: gaussXi(1) = 0.5773502692
  gaussW(0) = 1#: gaussW(1) = 1#
  If lastDof < 0 Then Exit Sub
  ReDim P3SelfWeightForce(lastDof)
  TotalSelfWeight = 0#
  For i = 0 To NumberOfElement - 1
    If Not P3IsElementActive(i) Then GoTo NextSelfWeightElement
    If Elem(i).IsJoint Then GoTo NextSelfWeightElement
    materialIndex = Elem(i).MatNo
    thickness = Material(materialIndex).thickness
    gammaValue = Material(materialIndex).weight
    For a = 0 To 1
      For b = 0 To 1
        xiValue = gaussXi(a)
        etaValue = gaussXi(b)
        P6Q8Shape xiValue, etaValue, shapeN, dNxi, dNeta
        dXdxi = 0#: dYdxi = 0#: dXdeta = 0#: dYdeta = 0#
        For j = 0 To 7
          dXdxi = dXdxi + dNxi(j) * Elem(i).x(j)
          dYdxi = dYdxi + dNxi(j) * Elem(i).y(j)
          dXdeta = dXdeta + dNeta(j) * Elem(i).x(j)
          dYdeta = dYdeta + dNeta(j) * Elem(i).y(j)
        Next j
        detJ = dXdxi * dYdeta - dYdxi * dXdeta
        dVolume = detJ * gaussW(a) * gaussW(b) * thickness
        TotalSelfWeight = TotalSelfWeight + gammaValue * dVolume
        For j = 0 To 7
          shapeValue = gammaValue * shapeN(j) * dVolume
          dofId = 2 * Elem(i).node(j)
          If dofId >= 0 And dofId < nDof Then
            P3SelfWeightForce(dofId) = P3SelfWeightForce(dofId) + shapeValue * nh
            P3SelfWeightForce(dofId + 1) = P3SelfWeightForce(dofId + 1) + shapeValue * nv
          End If
        Next j
      Next b
    Next a
NextSelfWeightElement:
  Next i
End Sub

Private Sub P3CaptureSupportBoundary()
  Dim i As Long
  If lastDof < 0 Then Exit Sub
  ReDim P3SupportNodeCond(lastDof)
  ReDim P3SupportBoundaryDisp(lastDof)
  ReDim P3OrphanHeld(lastDof)
  For i = 0 To lastDof
    P3SupportNodeCond(i) = NodeCond(i)
    P3SupportBoundaryDisp(i) = P3BoundaryDisp(i)
    P3OrphanHeld(i) = False
  Next i
  P3SupportCondReady = True
End Sub

Private Sub P3RestoreSupportBoundary()
  Dim i As Long
  If lastDof < 0 Then Exit Sub
  If Not P3SupportCondReady Then Exit Sub
  For i = 0 To lastDof
    NodeCond(i) = P3SupportNodeCond(i)
    P3BoundaryDisp(i) = P3SupportBoundaryDisp(i)
  Next i
  P3BumpConstraintGeneration
End Sub

Public Sub P3HoldOrphanDofs()
  Dim elementId As Long, j As Long, orig As Long, freeNode As Long, internalNode As Long, dofId As Long
  Dim connected() As Boolean
  If NumberOfFreeNode <= 0 Then Exit Sub
  ReDim connected(0 To NumberOfFreeNode - 1)
  If lastDof >= 0 Then
    If Not P3SupportCondReady Then
      P3CaptureSupportBoundary
    ElseIf (Not P3OrphanHeldAllocated()) Then
      ReDim P3OrphanHeld(lastDof)
    End If
  End If
  For elementId = 0 To NumberOfElement - 1
    If P3IsElementActive(elementId) Then
      For j = 0 To 7
        freeNode = Elem(elementId).node(j)
        If freeNode >= 0 And freeNode <= UBound(connected) Then connected(freeNode) = True
      Next j
    End If
  Next elementId
  For orig = 0 To NumberOfNode - 1
    If orig < LBound(FEMFreeNodeMap) Or orig > UBound(FEMFreeNodeMap) Then GoTo NextOrphanNode
    freeNode = FEMFreeNodeMap(orig)
    If freeNode < 0 Then GoTo NextOrphanNode
    internalNode = P6GetInternalFreeNode(freeNode)
    If internalNode < 0 Or internalNode >= NumberOfFreeNode Then GoTo NextOrphanNode
    If Not connected(internalNode) Then
      For dofId = internalNode * 2 To internalNode * 2 + 1
        If dofId >= 0 And dofId <= lastDof Then
          NodeCond(dofId) = 1
          P3BoundaryDisp(dofId) = TDisp(dofId)
          If P3OrphanHeldAllocated() Then P3OrphanHeld(dofId) = True
        End If
      Next dofId
    Else
      For dofId = internalNode * 2 To internalNode * 2 + 1
        If dofId >= 0 And dofId <= lastDof Then
          If P3OrphanHeldAllocated() Then
            If P3OrphanHeld(dofId) Then
              NodeCond(dofId) = P3SupportNodeCond(dofId)
              P3BoundaryDisp(dofId) = P3SupportBoundaryDisp(dofId)
              P3OrphanHeld(dofId) = False
            End If
          End If
        End If
      Next dofId
    End If
NextOrphanNode:
  Next orig
  P3BumpConstraintGeneration
End Sub

Private Function P3OrphanHeldAllocated() As Boolean
  On Error GoTo NotAllocated
  If lastDof < 0 Then Exit Function
  P3OrphanHeldAllocated = (UBound(P3OrphanHeld) >= lastDof)
  Exit Function
NotAllocated:
  P3OrphanHeldAllocated = False
End Function

Public Sub P3BumpConstraintGeneration()
  P3ConstraintGeneration = P3ConstraintGeneration + 1
  If P3ConstraintGeneration <= 0 Then P3ConstraintGeneration = 1
End Sub

Private Function P3JointContactChangedSinceFactor() As Boolean
  Dim i As Long, k As Long, gp As Long
  P3JointContactChangedSinceFactor = False
  If Not P3HasJointElements Then Exit Function
  If Not P6FactorReady Then
    P3JointContactChangedSinceFactor = True
    Exit Function
  End If
  If P6FactorJointContactN < 1 Then
    P3JointContactChangedSinceFactor = True
    Exit Function
  End If
  For i = 0 To P6FactorJointContactN - 1
    k = P6FactorJointElem(i)
    If k < 0 Or k > NumberOfElement - 1 Then
      P3JointContactChangedSinceFactor = True
      Exit Function
    End If
    For gp = 0 To 2
      If Elem(k).JointContact(gp) <> P6FactorJointContact(gp, i) Then
        P3JointContactChangedSinceFactor = True
        Exit Function
      End If
    Next gp
  Next i
End Function

Public Function P3FactorPlasticXorCount() As Long
  Dim k As Long, gp As Long, nXor As Long, bitValue As Long, oldOn As Boolean, nowOn As Boolean
  P3FactorPlasticXorCount = 0
  If NumberOfElement < 1 Then Exit Function
  If P6FactorYieldMaskN <> NumberOfElement Then
    P3FactorPlasticXorCount = P3ActivePlasticPointCount + 1
    Exit Function
  End If
  nXor = 0
  For k = 0 To NumberOfElement - 1
    If P3IsElementActive(k) Then
      bitValue = 1
      For gp = 0 To 3
        nowOn = Elem(k).P2Yielded(gp)
        oldOn = ((P6FactorYieldMask(k) And bitValue) <> 0)
        If nowOn Xor oldOn Then nXor = nXor + 1
        bitValue = bitValue * 2
      Next gp
    End If
  Next k
  P3FactorPlasticXorCount = nXor
End Function

Private Function P3AnyTangentDirty() As Boolean
  Dim k As Long
  P3AnyTangentDirty = False
  If NumberOfElement < 1 Then Exit Function
  On Error GoTo DirtySkip
  For k = 0 To NumberOfElement - 1
    If P3IsElementActive(k) Then
      If Elem(k).TangentDirty Then
        P3AnyTangentDirty = True
        Exit Function
      End If
    End If
  Next k
  Exit Function
DirtySkip:
  Err.Clear
  P3AnyTangentDirty = True
End Function

Private Function GetNewtonPolicy(ByVal analysisKind As String) As NewtonPolicy
  Select Case UCase$(Trim$(analysisKind))
    Case "GRAVITY"
      GetNewtonPolicy = NP_NORMAL
    Case "SRM"
      GetNewtonPolicy = NP_ROBUST_SRM
    Case Else
      GetNewtonPolicy = NP_NORMAL
  End Select
End Function

Private Sub P3ApplyNewtonPolicy(ByVal analysisKind As String)
  Dim enabledText As String
  If P3SrmTrialRunning Then
    P3NewtonPolicy = GetNewtonPolicy("SRM")
  Else
    P3NewtonPolicy = GetNewtonPolicy(analysisKind)
  End If
  P3Ls50Enabled = (P3NewtonPolicy = NP_ROBUST_SRM)
  P6LimitMonitorOn = P3Ls50Enabled
  If P3Ls50Enabled Then
    enabledText = "1"
  Else
    enabledText = "0"
  End If
  If P3NewtonPolicy = NP_ROBUST_SRM Then
    FEMAppendRunLog "NEWTON", "POLICY", "NewtonPolicy=ROBUST_SRM;LS50Enabled=" & enabledText
  Else
    FEMAppendRunLog "NEWTON", "POLICY", "NewtonPolicy=NORMAL;LS50Enabled=" & enabledText
  End If
End Sub

Private Function CanReuseFactorAfterLineSearch(ByVal policy As NewtonPolicy, ByVal alpha As Double, ByVal qLS As Double, ByVal residualGrowing As Boolean, ByVal justCutback As Boolean, ByVal activeSetForced As Boolean, ByVal factorValid As Boolean) As Boolean
  CanReuseFactorAfterLineSearch = False
  If policy <> NP_ROBUST_SRM Then
    P6Ls50GatePolicy = P6Ls50GatePolicy + 1
    Exit Function
  End If
  If Not P3Ls50Enabled Then
    P6Ls50GateDisabled = P6Ls50GateDisabled + 1
    Exit Function
  End If
  If Not factorValid Then
    P6Ls50GateFactor = P6Ls50GateFactor + 1
    Exit Function
  End If
  If justCutback Then
    P6Ls50GateCutback = P6Ls50GateCutback + 1
    Exit Function
  End If
  If activeSetForced Then
    P6Ls50GateActive = P6Ls50GateActive + 1
    Exit Function
  End If
  If residualGrowing Then
    P6Ls50GateResidual = P6Ls50GateResidual + 1
    Exit Function
  End If
  If Abs(alpha - 0.5) > 0.000000000001 Then
    P6Ls50GateAlpha = P6Ls50GateAlpha + 1
    Exit Function
  End If
  If qLS > 0.7 Then
    P6Ls50GateQls = P6Ls50GateQls + 1
    Exit Function
  End If
  CanReuseFactorAfterLineSearch = True
End Function

Private Function NextFsByBracket(ByVal fsLower As Double, ByVal fsUpper As Double, ByVal targetTol As Double) As Double
  Dim width As Double, mid As Double, frac As Double
  Dim passCost As Double, failCost As Double, totalCost As Double
  width = fsUpper - fsLower
  If width <= targetTol Then
    NextFsByBracket = fsLower
    Exit Function
  End If
  mid = 0.5 * (fsLower + fsUpper)
  If Not V3CostSearchEnabled Then
    NextFsByBracket = mid
    Exit Function
  End If
  V3BracketStep = V3BracketStep + 1
  If (V3BracketStep Mod 2) = 1 Then
    NextFsByBracket = mid
    Exit Function
  End If
  If width <= 2# * targetTol Then
    NextFsByBracket = mid
    Exit Function
  End If
  If V3PassCostSamples < 1 Or V3FailCostSamples < 1 Then
    NextFsByBracket = mid
    Exit Function
  End If
  passCost = V3LastPassCostSec
  failCost = V3LastFailCostSec
  totalCost = passCost + failCost
  If passCost <= 0# Or failCost <= 0# Or totalCost <= 0# Then
    NextFsByBracket = mid
    Exit Function
  End If
  ' Spend more interval on the cheaper side, while staying guarded.
  frac = passCost / totalCost
  If frac < 0.35 Then frac = 0.35
  If frac > 0.65 Then frac = 0.65
  NextFsByBracket = fsLower + frac * width
  If NextFsByBracket <= fsLower Or NextFsByBracket >= fsUpper Then NextFsByBracket = mid
  If AccelTrace Then P6SolverEvent "SRM_COST_SEARCH", "lower=" & Format$(fsLower, "0.000000000000") & ";upper=" & Format$(fsUpper, "0.000000000000") & ";next=" & Format$(NextFsByBracket, "0.000000000000") & ";pass_cost=" & Format$(passCost, "0.000") & ";fail_cost=" & Format$(failCost, "0.000")
End Function

Private Sub P3PublishFos(ByVal fsPass As Double, ByVal fsFail As Double, ByVal haveFail As Boolean)
  P3FosPass = fsPass
  P3FosBracket = haveFail
  If haveFail And fsFail > fsPass Then
    P3FosFail = fsFail
    P3FosWidth = fsFail - fsPass
    P3FosMid = 0.5 * (fsPass + fsFail)
    P3SrmNote = "FOS_PASS=" & Format$(fsPass, "0.000")
    P3SrmNote = P3SrmNote & " FOS_FAIL=" & Format$(fsFail, "0.000")
    P3SrmNote = P3SrmNote & " FOS_MID=" & Format$(P3FosMid, "0.000")
    P3SrmNote = P3SrmNote & " FOS_WIDTH=" & Format$(P3FosWidth, "0.000")
    P3SrmNote = P3SrmNote & " FOS=" & Format$(P3FosMid, "0.000") & " ± " & Format$(P3FosWidth * 0.5, "0.000")
  Else
    P3FosFail = 0#
    P3FosMid = fsPass
    P3FosWidth = 0#
    P3FosBracket = False
    P3SrmNote = "FOS_PASS=" & Format$(fsPass, "0.000") & " FOS_FAIL=NA FOS_MID=" & Format$(fsPass, "0.000") & " FOS_WIDTH=NA"
  End If
  P3SrmNote = P3SrmNote & " 試行=" & CStr(P3SrmTrialCount)
End Sub

Private Function P3FosDisplayFs() As Double
  If P3FosBracket Then
    P3FosDisplayFs = P3FosMid
  ElseIf P3FosPass > 0# Then
    P3FosDisplayFs = P3FosPass
  Else
    P3FosDisplayFs = FSS
  End If
End Function

Private Sub InvalidateTangent(ByVal reason As String)
  P6FactorReady = False
  P3ForceTangentRebuild = True
  P3ForceRebuildReason = reason
  P3Ls50Pending = False
  P3Ls50Watch = False
  P3LastRebuildWhy = reason
  AccelRecordReason reason
  P3AccelResetAA
  Select Case reason
    Case "FIRST"
      P6PerfRebuildFirst = P6PerfRebuildFirst + 1
    Case "RESID"
      P6PerfRebuildResidual = P6PerfRebuildResidual + 1
    Case "STAG", "LS50_STAGNATION"
      P6PerfRebuildStale = P6PerfRebuildStale + 1
    Case "ACTIVE"
      P6PerfRebuildActiveSet = P6PerfRebuildActiveSet + 1
    Case "FACINV"
      P6PerfRebuildFactorInvalid = P6PerfRebuildFactorInvalid + 1
    Case Else
      P6PerfRebuildForce = P6PerfRebuildForce + 1
      Select Case reason
        Case "STAGE_START"
          P6PerfRebuildForceStart = P6PerfRebuildForceStart + 1
        Case "CUTBACK"
          P6PerfRebuildForceCutback = P6PerfRebuildForceCutback + 1
        Case "LINESEARCH_CUT"
          P6PerfRebuildForceLsCut = P6PerfRebuildForceLsCut + 1
        Case "LINESEARCH_RETRY"
          P6PerfRebuildForceLsRetry = P6PerfRebuildForceLsRetry + 1
        Case Else
          P6PerfRebuildForceLsOk = P6PerfRebuildForceLsOk + 1
      End Select
  End Select
End Sub

Private Function P3ShouldRebuildTangent(ByVal localIteration As Long, ByVal residualNow As Double) As Boolean
  Dim plasticJump As Long, plasticRef As Long, xorLimit As Long, factorXor As Long
  Dim ls50 As Boolean, watch As Boolean, heldReason As String
  Dim v2Eligible As Boolean, v2Path As Long, gateReason As String
  P3ShouldRebuildTangent = True
  v2Eligible = (localIteration = 0 And AccelV2a And Not AccelBaselineRetry And AccelIncrementEligible And Not P3AfterCutback)
  v2Path = AdaptiveV2aPath()
  ls50 = P3Ls50Pending
  watch = P3Ls50Watch
  P3Ls50Pending = False
  P3Ls50Watch = False
  If P3ForceTangentRebuild Then
    If v2Eligible Then PolicyLogGate v2Path, "FORCE_REBUILD"
    heldReason = P3ForceRebuildReason
    P3ForceTangentRebuild = False
    P3ForceRebuildReason = vbNullString
    InvalidateTangent heldReason
    P3ForceTangentRebuild = False
    Exit Function
  End If
  If Not P6CanReuseAssembledFactor() Then
    If v2Eligible Then PolicyLogGate v2Path, "FACTOR_INVALID"
    InvalidateTangent "FACINV"
    P3ForceTangentRebuild = False
    Exit Function
  End If
  factorXor = P3FactorPlasticXorCount()
  If AccelFirstReusePending And localIteration > 0 Then
    AccelFirstReusePending = False
    If P3PrevResidualNorm <= 0# Or residualNow > 0.7 * P3PrevResidualNorm Then
      InvalidateTangent "V2A_POOR_REDUCTION"
      P3ForceTangentRebuild = False
      Exit Function
    End If
  End If
  If localIteration <= 0 And StepRecoveryFreshLU() Then
    If v2Eligible Then
      PolicyLogGate v2Path, "STEP_PROBE_FRESH_LU"
      StepRecoveryNoteFreshLUSkip
    End If
    InvalidateTangent "STEP_PROBE_FRESH_LU"
    P3ForceTangentRebuild = False
    Exit Function
  End If
  If localIteration <= 0 Then
    If AccelV2a And Not AccelBaselineRetry And AccelIncrementEligible And Not P3AfterCutback And (P3ActivePlasticPointCount <> 0 Or P3AnyTangentDirty()) Then
      If AdaptiveAllowMethod(2) Then
      If factorXor = 0 Then
        If AccelSmallTangentChange(gateReason) Then
        AccelFirstReusePending = True
        AccelFirstReuseCount = AccelFirstReuseCount + 1
        If AccelV2aPreviousIterations > 5 Then AccelV2aLongReuses = AccelV2aLongReuses + 1
        AdaptiveUseMethod 2
        AdaptiveV2aBeginUse
        P3ShouldRebuildTangent = False
        If AccelTrace Then P6SolverEvent "V2A_FIRST_REUSE", "approximate_jacobian=True;previous_iterations=" & CStr(AccelV2aPreviousIterations) & ";predictor_history=" & CStr(AccelHistoryReady) & ";long_increment=" & CStr(AccelV2aPreviousIterations > 5)
        Exit Function
        Else
          PolicyLogGate v2Path, gateReason
        End If
      Else
        PolicyLogGate v2Path, "PLASTIC_XOR"
      End If
      End If
    ElseIf v2Eligible Then
      PolicyLogGate v2Path, "ELASTIC_REUSE"
    End If
    If P3ActivePlasticPointCount <> 0 Or factorXor <> 0 Or P3JointContactChangedSinceFactor() Or P3AnyTangentDirty() Then
      InvalidateTangent "FIRST"
      P3ForceTangentRebuild = False
      Exit Function
    End If
    P3ShouldRebuildTangent = False
    Exit Function
  End If
  If factorXor > 0 Then
    plasticRef = P3ActivePlasticPointCount
    If P6FactorActivePlastic > plasticRef Then plasticRef = P6FactorActivePlastic
    xorLimit = plasticRef \ 10
    If xorLimit < 2 Then xorLimit = 2
    If factorXor > xorLimit Then
      InvalidateTangent "ACTIVE"
      P3ForceTangentRebuild = False
      Exit Function
    End If
  End If
  plasticJump = P3ActivePlasticPointCount - P6FactorActivePlastic
  If plasticJump < 0 Then plasticJump = -plasticJump
  If plasticJump > (P6FactorActivePlastic \ 10) + 1 Then
    InvalidateTangent "ACTIVE"
    P3ForceTangentRebuild = False
    Exit Function
  End If
  If watch Then
    If P3PrevResidualNorm > 0# Then
      If residualNow > P3PrevResidualNorm Then
        P6Ls50Worse = P6Ls50Worse + 1
        P6Ls50Event "LS50_REBUILD", "post;why=QNEXT"
        InvalidateTangent "LS50_QNEXT_BAD"
        P3ForceTangentRebuild = False
        Exit Function
      End If
      If residualNow > 0.7 * P3PrevResidualNorm Then
        P6Ls50Resid = P6Ls50Resid + 1
        P6Ls50Event "LS50_REBUILD", "post;why=RESID"
        InvalidateTangent "LS50"
        P3ForceTangentRebuild = False
        Exit Function
      End If
    End If
    If P3TangentStaleIters >= 3 Then
      P3TangentStaleIters = 0
      P6Ls50Stag = P6Ls50Stag + 1
      P6Ls50Event "LS50_REBUILD", "post;why=STAG"
      InvalidateTangent "LS50_STAGNATION"
      P3ForceTangentRebuild = False
      Exit Function
    End If
    P6Ls50Event "LS50_REUSE", "post;why=RECOVERED"
  End If
  If ls50 Then
    P6Ls50Tried = P6Ls50Tried + 1
    If P3PrevResidualNorm > 0# Then
      If residualNow > P3PrevResidualNorm Then
        P6Ls50Worse = P6Ls50Worse + 1
        P6Ls50Event "LS50_REBUILD", "pre;why=QNEXT"
        InvalidateTangent "LS50_QNEXT_BAD"
        P3ForceTangentRebuild = False
        Exit Function
      End If
      If residualNow > 0.7 * P3PrevResidualNorm Then
        P6Ls50Resid = P6Ls50Resid + 1
        P6Ls50Event "LS50_REBUILD", "pre;why=RESID"
        InvalidateTangent "LS50"
        P3ForceTangentRebuild = False
        Exit Function
      End If
    End If
    P3TangentStaleIters = P3TangentStaleIters + 1
    If P3TangentStaleIters >= 3 Then
      P3TangentStaleIters = 0
      P6Ls50Stag = P6Ls50Stag + 1
      P6Ls50Event "LS50_REBUILD", "pre;why=STAG"
      InvalidateTangent "LS50_STAGNATION"
      P3ForceTangentRebuild = False
      Exit Function
    End If
    P6Ls50Ok = P6Ls50Ok + 1
    P6Ls50Event "LS50_REUSE", "pre"
    P6Ls50Saved = P6Ls50Saved + 1
    P3Ls50Watch = True
    P3ShouldRebuildTangent = False
    Exit Function
  End If
  If P3PrevResidualNorm > 0# Then
    If residualNow > 0.7 * P3PrevResidualNorm Then
      If Not AccelCostAllowsReuse(residualNow / P3PrevResidualNorm) Then
        InvalidateTangent "RESID"
        P3ForceTangentRebuild = False
        Exit Function
      End If
    End If
  End If
  P3TangentStaleIters = P3TangentStaleIters + 1
  If P3TangentStaleIters >= 3 Then
    P3TangentStaleIters = 0
    InvalidateTangent "STAG"
    P3ForceTangentRebuild = False
    Exit Function
  End If
  P3ShouldRebuildTangent = False
End Function

Private Sub P3RequestTangentRebuild(ByVal reasonText As String)
  ' Flag only. Counting and P6FactorReady=False happen when the next Newton rebuilds.
  ' An LS50 accept can still clear the flag and keep the current factor.
  P3ForceTangentRebuild = True
  P3ForceRebuildReason = reasonText
  P3Ls50Pending = False
End Sub

Private Function P3ActiveSetForcesRebuild() As Boolean
  Dim plasticJump As Long, plasticRef As Long, xorLimit As Long, factorXor As Long
  P3ActiveSetForcesRebuild = False
  factorXor = P3FactorPlasticXorCount()
  If factorXor > 0 Then
    plasticRef = P3ActivePlasticPointCount
    If P6FactorActivePlastic > plasticRef Then plasticRef = P6FactorActivePlastic
    xorLimit = plasticRef \ 10
    If xorLimit < 2 Then xorLimit = 2
    If factorXor > xorLimit Then
      P3ActiveSetForcesRebuild = True
      Exit Function
    End If
  End If
  plasticJump = P3ActivePlasticPointCount - P6FactorActivePlastic
  If plasticJump < 0 Then plasticJump = -plasticJump
  If plasticJump > (P6FactorActivePlastic \ 10) + 1 Then P3ActiveSetForcesRebuild = True
  If P3JointContactChangedSinceFactor() Then P3ActiveSetForcesRebuild = True
End Function

Private Sub P3AcceptDampedLineSearch(ByVal alpha As Double, ByVal residualBefore As Double)
  Dim q As Double
  Dim growing As Boolean
  P6LsNoteAccept alpha, True
  P6LsNoteQ residualBefore, ResidualNormFree
  q = 2#
  If residualBefore > 0# Then q = ResidualNormFree / residualBefore
  growing = (residualBefore > 0# And ResidualNormFree > residualBefore)
  If CanReuseFactorAfterLineSearch(P3NewtonPolicy, alpha, q, growing, P3AfterCutback, P3ActiveSetForcesRebuild(), P6CanReuseAssembledFactor()) Then
    P3ForceTangentRebuild = False
    P3ForceRebuildReason = vbNullString
    P3Ls50Pending = True
  Else
    P3RequestTangentRebuild "LINESEARCH_OK"
  End If
End Sub

Private Function P3RebuildTangentFromSpmat() As Boolean
  Dim k As Long, materialIndex As Long, im As Long, i As Long, j As Long
  Dim hasSpmat As Boolean
  Dim t0 As Double
  Dim errNum As Long, errSrc As String, errDesc As String
  P3RebuildTangentFromSpmat = False
  FEMLastProc = "P3RebuildTangentFromSpmat"
  t0 = Timer
  On Error GoTo TangentFail
  P6EnsureGaussCache
  For k = 0 To NumberOfElement - 1
    If P3IsElementActive(k) Then
      materialIndex = Elem(k).MatNo
      If Elem(k).IsJoint Then
        P3JointEval k, False
        If P3FlowPolicyIsInconsistent() Then P3SymmetrizeKmat Elem(k).kmat
        Elem(k).TangentDirty = False
        P3TangentDirtyRebuildCount = P3TangentDirtyRebuildCount + 1
        GoTo NextTangentElement
      End If
      If Not Elem(k).TangentDirty Then GoTo NextTangentElement
      If materialIndex >= 0 And materialIndex <= UBound(Material) Then
        For im = 0 To 3
          If Not Elem(k).SpmatValid(im) Then
            For i = 0 To 2
              For j = 0 To 15
                Elem(k).Spmat(i, j, im) = Elem(k).Smat(i, j, im)
              Next j
            Next i
          End If
        Next im
        SetElmStiffness Material(materialIndex).thickness, Material(materialIndex).weight, Elem(k).Bmat, Elem(k).Spmat, Elem(k).kmat, Elem(k).ElmWeight, Elem(k).dj, P6GaussN
        P6AddQ8HourglassStabilization Elem(k), Material(materialIndex).thickness, True
        If P3FlowPolicyIsInconsistent() Then
          P3SymmetrizeKmat Elem(k).kmat
        ElseIf P3KmatIsUnsymmetric(Elem(k).kmat) Then
          SetAnalysisFailure RESULT_MATERIAL_ERROR, "BAND対称ソルバは非対称接線に未対応です（非関連流れ φ≠ψ の塑性接線）。FLOW_POLICY=INCONSISTENT または DAVIS を指定してください。要素=" & CStr(k + 1), vbObjectError + 3240, k + 1, -1, CurrentIncrement, CurrentIteration
          GoTo TangentDone
        End If
        Elem(k).TangentDirty = False
        P3TangentDirtyRebuildCount = P3TangentDirtyRebuildCount + 1
      End If
    End If
NextTangentElement:
  Next k
  P6FactorReady = False
  SetTotalMat
  P3LastTangentPlasticCount = P3ActivePlasticPointCount
  P3TangentStaleIters = 0
  P3RebuildTangentFromSpmat = True
  P3TangentJustRebuilt = True
  GoTo TangentDone
TangentFail:
  errNum = Err.Number
  errSrc = Err.source
  errDesc = Err.Description
  Resume TangentDone
TangentDone:
  On Error GoTo 0
  P6ProfTangentCount = P6ProfTangentCount + 1
  P6ProfTangentMs = P6ProfTangentMs + P6ElapsedMs(t0)
  If errNum <> 0 Then Err.Raise errNum, errSrc, errDesc
End Function

Private Sub P3CommitLoadLock(ByVal lockSelfWeight As Boolean, ByVal lockApplied As Boolean)
  Dim i As Long, swReady As Boolean, apReady As Boolean
  If lastDof < 0 Then Exit Sub
  swReady = False
  apReady = False
  On Error Resume Next
  swReady = (UBound(P3LockedSelfWeight) = lastDof)
  apReady = (UBound(P3LockedApplied) = lastDof)
  Err.Clear
  On Error GoTo 0
  If lockSelfWeight And Not swReady Then ReDim P3LockedSelfWeight(lastDof)
  If lockApplied And Not apReady Then ReDim P3LockedApplied(lastDof)
  ReDim P3CommittedNodeCond(lastDof)
  ReDim P3CommittedBoundaryDisp(lastDof)
  For i = 0 To lastDof
    If lockSelfWeight Then P3LockedSelfWeight(i) = P3SelfWeightForce(i)
    If lockApplied Then P3LockedApplied(i) = P3AppliedForce(i)
    P3CommittedNodeCond(i) = NodeCond(i)
    P3CommittedBoundaryDisp(i) = P3BoundaryDisp(i)
  Next i
  P3CommittedBoundaryReady = True
End Sub

Private Sub P3RefreshActiveStiffness()
  P6InvalidateTangentGeneration
  SetTotalMat
End Sub

Private Sub P3ApplyPrefixState()
  Dim i As Long
  P3InitElementActive
  P3ApplyBirthDeathUpTo P3ActivePrefix
  If lastDof >= 0 Then
    If Not P3SupportCondReady Then
      P3CaptureSupportBoundary
    Else
      P3RestoreSupportBoundary
    End If
  End If
  P3OverlayLoadingSheet
  P3HoldOrphanDofs
  P3AssembleSelfWeightOnly
End Sub

Private Function P3KmatIsUnsymmetric(ByRef kmat() As Double) As Boolean
  Dim i As Long, j As Long, gap As Double, scaleValue As Double
  P3KmatIsUnsymmetric = False
  For i = 0 To 15
    For j = i + 1 To 15
      gap = Abs(kmat(i, j) - kmat(j, i))
      scaleValue = 1# + Abs(kmat(i, j)) + Abs(kmat(j, i))
      If gap > 0.00000001 * scaleValue Then
        P3KmatIsUnsymmetric = True
        Exit Function
      End If
    Next j
  Next i
End Function

Private Sub P3WriteStageSnapshot()
  Dim outputMode As String
  If P3ReplayQuiet Then Exit Sub
  outputMode = UCase$(Trim$(P6ReadTextSetting("OUTPUT_STAGE_MODE", "FINAL")))
  If outputMode <> "ALL" Then Exit Sub
  P3PrepareOutputResidual
  SaveDisp
  SaveStress
End Sub

Private Function P3StageMaterialId(ByVal stageIndex As Long) As Long
  P3StageMaterialId = 0
  If stageIndex < 1 Or stageIndex > P3StageN Then Exit Function
  If IsNumeric(P3StageGroup(stageIndex)) Then P3StageMaterialId = CLng(Val(P3StageGroup(stageIndex)))
End Function

Private Sub P3EnsureStateWorkspace()
  If lastDof >= 0 Then
    If P3StateWorkspaceDofN <> lastDof Then
      ReDim P3CommittedDisp(lastDof)
      ReDim P3BoundaryDisp(lastDof)
      ReDim P3FinalInternalForce(lastDof)
      P3StateWorkspaceDofN = lastDof
    End If
  End If
  If NumberOfElement >= 1 Then
    If P3StateWorkspaceElemN <> NumberOfElement Then
      ReDim P3CommittedPlasticStrain(0 To 3, 0 To 3, 0 To NumberOfElement - 1)
      ReDim P3CommittedPlasticMultiplier(0 To 3, 0 To NumberOfElement - 1)
      ReDim P3CommittedYieldFunction(0 To 3, 0 To NumberOfElement - 1)
      ReDim P3CommittedYielded(0 To 3, 0 To NumberOfElement - 1)
      ReDim P3CommittedPrincipalStress(0 To 2, 0 To 3, 0 To NumberOfElement - 1)
      ReDim P3CommittedSpmat(0 To 2, 0 To 15, 0 To 3, 0 To NumberOfElement - 1)
      ReDim P3CommittedDpmat(0 To 2, 0 To 2, 0 To 3, 0 To NumberOfElement - 1)
      ReDim P3CommittedSpmatValid(0 To 3, 0 To NumberOfElement - 1)
      ReDim P3ActivePlasticPoint(0 To 3, 0 To NumberOfElement - 1)
      P3StateWorkspaceElemN = NumberOfElement
    End If
  End If
End Sub

Private Sub P3EnsureSrmSnapWorkspace()
  If lastDof >= 0 Then
    If P3SrmSnapWorkspaceDofN <> lastDof Then
      ReDim P3SrmSnapTDisp(lastDof)
      ReDim P3SrmSnapUDisp(lastDof)
      ReDim P3SrmSnapCommittedDisp(lastDof)
      ReDim P3SrmSnapInternal(lastDof)
      ReDim P3SrmSnapNodeCond(lastDof)
      ReDim P3SrmSnapBoundaryDisp(lastDof)
      ReDim P3SrmSnapApplied(lastDof)
      ReDim P3SrmSnapSelfWeight(lastDof)
      ReDim P3SrmSnapLockedSelf(lastDof)
      ReDim P3SrmSnapLockedApplied(lastDof)
      P3SrmSnapWorkspaceDofN = lastDof
    End If
  End If
  If NumberOfElement >= 1 Then
    If P3SrmSnapWorkspaceElemN <> NumberOfElement Then
      ReDim P3SrmSnapStmat(0 To 16, 0 To 3, 0 To NumberOfElement - 1)
      ReDim P3SrmSnapMStmat(0 To 16, 0 To 3, 0 To NumberOfElement - 1)
      ReDim P3SrmSnapJointContact(0 To 2, 0 To NumberOfElement - 1)
      ReDim P3SrmSnapActive(0 To NumberOfElement - 1)
      P3SrmSnapWorkspaceElemN = NumberOfElement
    End If
  End If
End Sub

Private Sub P3InitializeTrialState()
  Dim i As Long, j As Long, k As Long, im As Long
  P3EnsureStateWorkspace
  P3ActivePlasticPointCount = 0
  P3CurrentElementZeroIncrement = False
  PlasticLogResetState
  P3TrialStateInitialized = True
  P3TrialStateValid = False
  P3PlasticStateChangeCount = 0
  P3CommittedBoundaryReady = False
  For i = 0 To lastDof
    P3BoundaryDisp(i) = Disp(i)
    P3CommittedDisp(i) = 0#
    TDisp(i) = 0#
    UDisp(i) = 0#
    Disp(i) = 0#
  Next i
  For k = 0 To NumberOfElement - 1
    For im = 0 To 3
      For j = 0 To 16
        Elem(k).mStmat(j, im) = 0#
        Elem(k).Stmat(j, im) = 0#
      Next j
      For j = 0 To 3
        Elem(k).P2PlasticStrain(j, im) = 0#
        P3CommittedPlasticStrain(j, im, k) = 0#
        P3CommittedPlasticMultiplier(im, k) = 0#
        P3CommittedYieldFunction(im, k) = 0#
        P3CommittedYielded(im, k) = False
      Next j
      For j = 0 To 2
        Elem(k).P2PrincipalStress(j, im) = 0#
        P3CommittedPrincipalStress(j, im, k) = 0#
      Next j
      Elem(k).P2PlasticMultiplier(im) = 0#
      Elem(k).P2YieldFunction(im) = 0#
      Elem(k).P2Yielded(im) = False
      Elem(k).P2LastSuccessfulSubsteps(im) = 1
      Elem(k).P2EasySubstepStreak(im) = 0
      Elem(k).SpmatValid(im) = False
      P3CommittedSpmatValid(im, k) = False
      P3ActivePlasticPoint(im, k) = False
      For i = 0 To 2
        For j = 0 To 15
          P3CommittedSpmat(i, j, im, k) = 0#
        Next j
        For j = 0 To 2
          P3CommittedDpmat(i, j, im, k) = 0#
        Next j
      Next i
    Next im
    P3ResetElementElasticTangent k
  Next k
  P3InitElementActive
  P6FactorReady = False
  P3UseNonlinearResidual = False
  P3LastConvergedLoadFactor = 0#
  P3LastConvergedIncrement = 0
  P3SuccessfulIncrementCount = 0
  P3GlobalIterationCount = 0
  P3RetryCount = 0
  P3MaxTrialDisp = 0#
  P3MaxCorrection = 0#
  P3LastRelativeResidual = 0#
  P3MaxBoundaryDispError = 0#
  P3RelativeBoundaryDispError = 0#
  If P6MixedUP And P6PressureCount > 0 Then
    For i = 0 To P6PressureCount - 1
      P6Pressure(i) = 0#
      P6CommittedPressure(i) = 0#
      If P6PressureFixed(i) Then
        P6Pressure(i) = 0#
        P6CommittedPressure(i) = 0#
      End If
    Next i
  End If
  P3GravityCommitted = False
  If lastDof >= 0 Then
    ReDim P3LockedSelfWeight(lastDof)
    ReDim P3LockedApplied(lastDof)
  End If
  P3InitElementActive
  If lastDof >= 0 Then
    If Not P3SupportCondReady Then
      P3CaptureSupportBoundary
    Else
      P3RestoreSupportBoundary
    End If
  End If
  If P3StageN <= 0 Then
    P3ApplyPrefixState
  Else
    P3HoldOrphanDofs
    P3AssembleSelfWeightOnly
  End If
End Sub

Private Sub P3CopyCommittedDispToTrial()
  Dim i As Long
  For i = 0 To lastDof
    TDisp(i) = P3CommittedDisp(i)
    UDisp(i) = P3CommittedDisp(i)
  Next i
  If P6MixedUP And P6PressureCount > 0 Then
    For i = 0 To P6PressureCount - 1
      P6Pressure(i) = P6CommittedPressure(i)
    Next i
  End If
End Sub

Private Sub P3RestoreCommittedMaterialState(Optional ByVal invalidateFactor As Boolean = True)
  Dim i As Long, j As Long, k As Long, im As Long
  Dim trialActive As Boolean, committedActive As Boolean, historyYielded As Boolean
  For k = 0 To NumberOfElement - 1
    For im = 0 To 3
      For i = 0 To 16
        Elem(k).Stmat(i, im) = Elem(k).mStmat(i, im)
      Next i
      For i = 0 To 3
        Elem(k).P2PlasticStrain(i, im) = P3CommittedPlasticStrain(i, im, k)
      Next i
      Elem(k).P2PlasticMultiplier(im) = P3CommittedPlasticMultiplier(im, k)
      Elem(k).P2YieldFunction(im) = P3CommittedYieldFunction(im, k)
      historyYielded = P3CommittedYielded(im, k)
      Elem(k).P2Yielded(im) = historyYielded
      For j = 0 To 2
        Elem(k).P2PrincipalStress(j, im) = P3CommittedPrincipalStress(j, im, k)
      Next j
      trialActive = P3ActivePlasticPoint(im, k)
      committedActive = P3IsElementActive(k) And historyYielded
      If trialActive And Not committedActive Then P3ActivePlasticPointCount = P3ActivePlasticPointCount - 1
      If Not trialActive And committedActive Then P3ActivePlasticPointCount = P3ActivePlasticPointCount + 1
      P3ActivePlasticPoint(im, k) = committedActive
      Elem(k).SpmatValid(im) = P3CommittedSpmatValid(im, k)
      For i = 0 To 2
        For j = 0 To 15
          Elem(k).Spmat(i, j, im) = P3CommittedSpmat(i, j, im, k)
        Next j
        For j = 0 To 2
          Elem(k).Dpmat(i, j, im) = P3CommittedDpmat(i, j, im, k)
        Next j
      Next i
    Next im
  Next k
  If P6MixedUP And P6PressureCount > 0 Then
    For k = 0 To P6PressureCount - 1
      P6Pressure(k) = P6CommittedPressure(k)
    Next k
  End If
  P3TrialStateValid = False
  If invalidateFactor Then
    P6FactorReady = False
    P3RequestTangentRebuild "CUTBACK"
    P3MarkAllTangentDirty
    P3AfterCutback = True
  End If
End Sub

Public Sub P3MarkAllTangentDirty()
  Dim k As Long
  If NumberOfElement < 1 Then Exit Sub
  On Error GoTo MarkSkip
  For k = 0 To NumberOfElement - 1
    Elem(k).TangentDirty = True
  Next k
  Exit Sub
MarkSkip:
  Err.Clear
End Sub

Public Sub P3DecaySubstepHint(ByVal elementId As Long, ByVal gaussId As Long)
  Dim lastHint As Long
  If elementId < 0 Or elementId >= NumberOfElement Then Exit Sub
  If gaussId < 0 Or gaussId > 3 Then Exit Sub
  lastHint = Elem(elementId).P2LastSuccessfulSubsteps(gaussId)
  If lastHint <= 1 Then
    Elem(elementId).P2LastSuccessfulSubsteps(gaussId) = 1
    Elem(elementId).P2EasySubstepStreak(gaussId) = 0
    Exit Sub
  End If
  Elem(elementId).P2EasySubstepStreak(gaussId) = Elem(elementId).P2EasySubstepStreak(gaussId) + 1
  If Elem(elementId).P2EasySubstepStreak(gaussId) >= 8 Then
    lastHint = lastHint \ 2
    If lastHint < 1 Then lastHint = 1
    Elem(elementId).P2LastSuccessfulSubsteps(gaussId) = lastHint
    Elem(elementId).P2EasySubstepStreak(gaussId) = 0
  End If
End Sub

Public Sub P3RememberSubstepHint(ByVal elementId As Long, ByVal gaussId As Long, ByVal usedSteps As Long)
  Dim lastHint As Long
  If elementId < 0 Or elementId >= NumberOfElement Then Exit Sub
  If gaussId < 0 Or gaussId > 3 Then Exit Sub
  If usedSteps < 1 Then usedSteps = 1
  lastHint = Elem(elementId).P2LastSuccessfulSubsteps(gaussId)
  If lastHint < 1 Then lastHint = 1
  If usedSteps > lastHint Then
    Elem(elementId).P2LastSuccessfulSubsteps(gaussId) = usedSteps
    Elem(elementId).P2EasySubstepStreak(gaussId) = 0
    Exit Sub
  End If
  Elem(elementId).P2EasySubstepStreak(gaussId) = Elem(elementId).P2EasySubstepStreak(gaussId) + 1
  If Elem(elementId).P2EasySubstepStreak(gaussId) >= 2 Then
    If lastHint > 1 Then
      lastHint = lastHint \ 2
      If lastHint < 1 Then lastHint = 1
      Elem(elementId).P2LastSuccessfulSubsteps(gaussId) = lastHint
    End If
    Elem(elementId).P2EasySubstepStreak(gaussId) = 0
  End If
End Sub

Public Sub P3RecountActivePlasticPoints()
  Dim k As Long, im As Long
  P3ActivePlasticPointCount = 0
  If NumberOfElement < 1 Then Exit Sub
  On Error GoTo RecountSkip
  If UBound(P3ActivePlasticPoint, 2) < NumberOfElement - 1 Then Exit Sub
  For k = 0 To NumberOfElement - 1
    For im = 0 To 3
      If P3IsElementActive(k) And Elem(k).P2Yielded(im) Then
        P3ActivePlasticPoint(im, k) = True
        P3ActivePlasticPointCount = P3ActivePlasticPointCount + 1
      Else
        P3ActivePlasticPoint(im, k) = False
      End If
    Next im
  Next k
  Exit Sub
RecountSkip:
  Err.Clear
End Sub

Private Function P3MaterialFailureIsFatal() As Boolean
  Dim code As Long
  P3MaterialFailureIsFatal = True
  If ResultStatus <> RESULT_MATERIAL_ERROR Then Exit Function
  code = AnalysisErrorNumber - vbObjectError
  If code = P2_FAILURE_NOT_CONVERGED Or code = P2_FAILURE_DENOMINATOR Then
    P3MaterialFailureIsFatal = False
  End If
End Function

Public Sub P3CommitMaterialState()
  Dim i As Long, j As Long, k As Long, im As Long
  Dim wasActive As Boolean, isActive As Boolean, historyYielded As Boolean
  For k = 0 To NumberOfElement - 1
    For im = 0 To 3
      For i = 0 To 16
        Elem(k).mStmat(i, im) = Elem(k).Stmat(i, im)
      Next i
      historyYielded = Elem(k).P2Yielded(im)
      wasActive = P3ActivePlasticPoint(im, k)
      isActive = P3IsElementActive(k) And historyYielded
      If isActive And Not wasActive Then P3ActivePlasticPointCount = P3ActivePlasticPointCount + 1
      If wasActive And Not isActive Then P3ActivePlasticPointCount = P3ActivePlasticPointCount - 1
      P3ActivePlasticPoint(im, k) = isActive
      For i = 0 To 3
        P3CommittedPlasticStrain(i, im, k) = Elem(k).P2PlasticStrain(i, im)
      Next i
      P3CommittedPlasticMultiplier(im, k) = Elem(k).P2PlasticMultiplier(im)
      P3CommittedYieldFunction(im, k) = Elem(k).P2YieldFunction(im)
      P3CommittedYielded(im, k) = historyYielded
      For j = 0 To 2
        P3CommittedPrincipalStress(j, im, k) = Elem(k).P2PrincipalStress(j, im)
      Next j
      P3CommittedSpmatValid(im, k) = Elem(k).SpmatValid(im)
      For i = 0 To 2
        For j = 0 To 15
          P3CommittedSpmat(i, j, im, k) = Elem(k).Spmat(i, j, im)
        Next j
        For j = 0 To 2
          P3CommittedDpmat(i, j, im, k) = Elem(k).Dpmat(i, j, im)
        Next j
      Next i
    Next im
  Next k
  If P6MixedUP And P6PressureCount > 0 Then
    For k = 0 To P6PressureCount - 1
      P6CommittedPressure(k) = P6Pressure(k)
    Next k
  End If
End Sub

Private Function P3EvaluateTrialState(ByRef incrementBase() As Double, ByRef internalForce() As Double) As Boolean
  Dim i As Long, j As Long, k As Long, materialIndex As Long
  Dim t0 As Double
  t0 = Timer
  P3EvaluateTrialState = False
  P3TrialStateValid = False
  P3PlasticStateChangeCount = 0
  For i = 0 To lastDof
    iNForce(i) = 0#
  Next i
  For k = 0 To NumberOfElement - 1
    If Not P3IsElementActive(k) Then GoTo NextEvalElement
    materialIndex = Elem(k).MatNo
    FailureElement = k
    If Elem(k).IsJoint Then
      P3JointEval k, True
      GoTo NextEvalElement
    End If
    With Elem(k)
      P3CurrentElementZeroIncrement = True
      For j = 0 To 15
        If .ElNode(j) < 0 Then GoTo NextEvalDof
        .u(j) = TDisp(.ElNode(j)) - incrementBase(.ElNode(j))
        If Abs(.u(j)) > 1E-30 Then P3CurrentElementZeroIncrement = False
NextEvalDof:
      Next j
      If Not Nrf(.Smat, .Spmat, .u, .Stmat, .mStmat, Material(materialIndex).fai, Material(materialIndex).cohesion, Material(materialIndex).psai, Material(materialIndex).Young, Material(materialIndex).Poisson, .Dmat, .Dpmat, .Bmat, materialIndex) Then
        P6ProfEvalCount = P6ProfEvalCount + 1
        P6ProfEvalMs = P6ProfEvalMs + P6ElapsedMs(t0)
        Exit Function
      End If
      CalcForce .ElNo, .Bmat, .dj, .Stmat
    End With
NextEvalElement:
  Next k
  P3CurrentElementZeroIncrement = False
  If P6MixedUP Then P6AddPressureToInternalForce
  For i = 0 To lastDof
    internalForce(i) = iNForce(i)
  Next i
  P6ProfEvalCount = P6ProfEvalCount + 1
  P6ProfEvalMs = P6ProfEvalMs + P6ElapsedMs(t0)
  P3TrialStateValid = True
  P3EvaluateTrialState = True
End Function

Private Sub P3BuildTargetForce(ByVal selfWeightFactor As Double, ByVal appliedFactor As Double, ByRef targetForce() As Double)
  Dim i As Long
  For i = 0 To lastDof
    targetForce(i) = P3LockedSelfWeight(i) + (P3SelfWeightForce(i) - P3LockedSelfWeight(i)) * selfWeightFactor _
                   + P3LockedApplied(i) + (P3AppliedForce(i) - P3LockedApplied(i)) * appliedFactor
  Next i
End Sub

Private Sub P3UpdateResidualMetrics(ByRef targetForce() As Double, ByRef internalForce() As Double)
  P3AccumulateResidualMetrics targetForce, internalForce, False
End Sub

Private Sub P3UpdateResidualNormOnly(ByRef targetForce() As Double, ByRef internalForce() As Double)
  P3AccumulateResidualMetrics targetForce, internalForce, True
End Sub

Private Sub P3AccumulateResidualMetrics(ByRef targetForce() As Double, ByRef internalForce() As Double, ByVal lite As Boolean)
  Dim i As Long, residualValue As Double, residualSquare As Double, forceSquare As Double, internalSquare As Double
  Dim energyValue As Double, energyScale As Double, dispScale As Double, maxAbsDisp As Double
  Dim pressureIndex As Long, displacementForceSquare As Double, displacementDriven As Boolean
  ResidualNormFull = 0#: ResidualNormFree = 0#: ForceNormFree = 0#: MaxAbsResidualFree = 0#
  If Not lite Then
    P6ReactionSumX = 0#: P6ReactionSumY = 0#
    P6EnsureReactionWorkspace
  End If
  For i = 0 To lastDof
    residualValue = internalForce(i) - targetForce(i)
    If Not lite Then Reaction(i) = residualValue
    If Abs(residualValue) > ResidualNormFull Then ResidualNormFull = Abs(residualValue)
    If NodeCond(i) = 0 Then
      residualSquare = residualSquare + residualValue * residualValue
      forceSquare = forceSquare + targetForce(i) * targetForce(i)
      internalSquare = internalSquare + internalForce(i) * internalForce(i)
      If Abs(residualValue) > MaxAbsResidualFree Then MaxAbsResidualFree = Abs(residualValue)
      If Abs(TDisp(i)) > maxAbsDisp Then maxAbsDisp = Abs(TDisp(i))
    Else
      ' With no force loading, prescribed-displacement reactions provide the
      ' physical force scale. A fixed 1E-12 scale amplifies roundoff residuals.
      If P3ForceRef <= 0.000000000001 Then
        displacementDriven = (Abs(P3BoundaryDisp(i)) > 1E-30)
        If Not displacementDriven And P3StageStartDispReady Then
          displacementDriven = (Abs(P3StageStartDisp(i)) > 1E-30)
        End If
        If displacementDriven Then displacementForceSquare = displacementForceSquare + internalForce(i) * internalForce(i)
      End If
      If Not lite Then
      If i Mod 2 = 0 Then
        P6ReactionSumX = P6ReactionSumX + residualValue
      Else
        P6ReactionSumY = P6ReactionSumY + residualValue
      End If
      End If
    End If
    If Not lite Then energyValue = energyValue + TDisp(i) * residualValue
  Next i
  If P6MixedUP Then
    P6ComputePressureResidual
    For pressureIndex = 0 To P6PressureCount - 1
      If Not P6PressureFixed(pressureIndex) Then
        residualValue = P6PressureResidual(pressureIndex)
        residualSquare = residualSquare + residualValue * residualValue
        If Abs(residualValue) > MaxAbsResidualFree Then MaxAbsResidualFree = Abs(residualValue)
      End If
    Next pressureIndex
  End If
  ResidualNormFree = Sqr(residualSquare)
  ForceNormFree = Sqr(forceSquare)
  P3InternalNormFree = Sqr(internalSquare)
  If maxAbsDisp > P3MaxTrialDisp Then P3MaxTrialDisp = maxAbsDisp
  energyScale = ForceNormFree
  If P3InternalNormFree > energyScale Then energyScale = P3InternalNormFree
  If P3ForceRef > energyScale Then energyScale = P3ForceRef
  If displacementForceSquare > energyScale * energyScale Then energyScale = Sqr(displacementForceSquare)
  If energyScale < 0.000000000001 Then energyScale = 0.000000000001
  RelativeResidualFree = ResidualNormFree / energyScale
  dispScale = maxAbsDisp
  If dispScale < P3DispRef Then dispScale = P3DispRef
  If dispScale < 0.000000000001 Then dispScale = 0.000000000001
  If lite Then
    EnergyError = 0#
  Else
    EnergyError = Abs(energyValue) / (energyScale * dispScale)
  End If
  P3LastRelativeResidual = RelativeResidualFree
End Sub

Private Sub P3ApplyCorrection(ByRef correctionRatio As Double)
  Dim i As Long, correctionNorm As Double, totalNorm As Double
  correctionNorm = 0#: totalNorm = 0#
  For i = 0 To lastDof
    TDisp(i) = TDisp(i) + Disp(i)
    UDisp(i) = TDisp(i)
    If NodeCond(i) = 0 Then
      If Abs(Disp(i)) > correctionNorm Then correctionNorm = Abs(Disp(i))
      If Abs(TDisp(i)) > totalNorm Then totalNorm = Abs(TDisp(i))
    End If
  Next i
  If totalNorm < P3DispRef Then
    If P3DispRef < 0.000000000001 Then
      correctionRatio = correctionNorm
    Else
      correctionRatio = correctionNorm / P3DispRef
    End If
  Else
    correctionRatio = correctionNorm / totalNorm
  End If
End Sub

Private Function P3HasNonzeroPrescribedDisp() As Boolean
  Dim i As Long, prescribedSquare As Double
  prescribedSquare = 0#
  For i = 0 To lastDof
    If NodeCond(i) <> 0 Then prescribedSquare = prescribedSquare + P3BoundaryDisp(i) * P3BoundaryDisp(i)
  Next i
  P3HasNonzeroPrescribedDisp = (prescribedSquare > 1E-30)
End Function

Private Function P3VectorHasChange(ByRef oldVec() As Double, ByRef newVec() As Double) As Boolean
  Dim i As Long, oldOk As Boolean, newOk As Boolean
  P3VectorHasChange = False
  If lastDof < 0 Then Exit Function
  On Error Resume Next
  oldOk = (UBound(oldVec) >= lastDof)
  If Err.Number <> 0 Then oldOk = False
  Err.Clear
  newOk = (UBound(newVec) >= lastDof)
  If Err.Number <> 0 Then newOk = False
  On Error GoTo 0
  If Not newOk Then Exit Function
  If Not oldOk Then
    For i = 0 To lastDof
      If Abs(newVec(i)) > 1E-30 Then
        P3VectorHasChange = True
        Exit Function
      End If
    Next i
    Exit Function
  End If
  For i = 0 To lastDof
    If Abs(newVec(i) - oldVec(i)) > 1E-30 Then
      P3VectorHasChange = True
      Exit Function
    End If
  Next i
End Function

Private Function P3PrescribedBoundaryChanged() As Boolean
  Dim i As Long
  P3PrescribedBoundaryChanged = False
  If lastDof < 0 Then Exit Function
  If Not P3CommittedBoundaryReady Then
    If Not P3SupportCondReady Then
      P3PrescribedBoundaryChanged = P3HasNonzeroPrescribedDisp()
      Exit Function
    End If
    For i = 0 To lastDof
      If NodeCond(i) <> P3SupportNodeCond(i) Then
        P3PrescribedBoundaryChanged = True
        Exit Function
      End If
      If NodeCond(i) <> 0 Then
        If Abs(P3BoundaryDisp(i) - P3SupportBoundaryDisp(i)) > 1E-30 Then
          P3PrescribedBoundaryChanged = True
          Exit Function
        End If
      End If
    Next i
    Exit Function
  End If
  For i = 0 To lastDof
    If NodeCond(i) <> P3CommittedNodeCond(i) Then
      P3PrescribedBoundaryChanged = True
      Exit Function
    End If
    If NodeCond(i) <> 0 Then
      If Abs(P3BoundaryDisp(i) - P3CommittedBoundaryDisp(i)) > 1E-30 Then
        P3PrescribedBoundaryChanged = True
        Exit Function
      End If
    End If
  Next i
End Function

Private Function P3TargetBoundaryDisp(ByVal dofId As Long, ByVal targetFactor As Double) As Double
  Dim startDisp As Double
  If NodeCond(dofId) = 0 Then
    P3TargetBoundaryDisp = 0#
    Exit Function
  End If
  startDisp = 0#
  If P3StageStartDispReady Then
    If dofId >= LBound(P3StageStartDisp) And dofId <= UBound(P3StageStartDisp) Then
      startDisp = P3StageStartDisp(dofId)
    End If
  ElseIf P3CommittedBoundaryReady Then
    If P3CommittedNodeCond(dofId) <> 0 Then startDisp = P3CommittedBoundaryDisp(dofId)
  End If
  P3TargetBoundaryDisp = startDisp + (P3BoundaryDisp(dofId) - startDisp) * targetFactor
End Function

Private Sub P3CaptureStageStartDisp()
  Dim i As Long
  P3StageStartDispReady = False
  If lastDof < 0 Then Exit Sub
  ReDim P3StageStartDisp(lastDof)
  On Error Resume Next
  If UBound(P3CommittedDisp) >= lastDof Then
    For i = 0 To lastDof
      P3StageStartDisp(i) = P3CommittedDisp(i)
    Next i
  Else
    For i = 0 To lastDof
      P3StageStartDisp(i) = TDisp(i)
    Next i
  End If
  On Error GoTo 0
  P3StageStartDispReady = True
End Sub

' Incremental BC factor for this stage. Always the stage target (not delta from previous), so GRAVITY then DISP stays correct.
Private Function P3PrescribedFactor(ByVal stageId As Long, ByVal targetFactor As Double) As Double
  P3PrescribedFactor = targetFactor
End Function

Private Sub P3UpdateBoundaryDispError(ByVal prescribedFactor As Double)
  Dim i As Long, errValue As Double, scaleValue As Double, targetDisp As Double
  P3MaxBoundaryDispError = 0#
  scaleValue = 0#
  For i = 0 To lastDof
    If NodeCond(i) <> 0 Then
      targetDisp = P3TargetBoundaryDisp(i, prescribedFactor)
      errValue = Abs(TDisp(i) - targetDisp)
      If errValue > P3MaxBoundaryDispError Then P3MaxBoundaryDispError = errValue
      If Abs(targetDisp) > scaleValue Then scaleValue = Abs(targetDisp)
    End If
  Next i
  If scaleValue < P3DispRef Then scaleValue = P3DispRef
  If scaleValue < 0.000000000001 Then scaleValue = 0.000000000001
  P3RelativeBoundaryDispError = P3MaxBoundaryDispError / scaleValue
End Sub

Private Function P3BoundaryDispSatisfied() As Boolean
  P3BoundaryDispSatisfied = (P3RelativeBoundaryDispError <= P3_BOUNDARY_DISP_TOLERANCE) Or (P3MaxBoundaryDispError <= 0.000000000001)
End Function

Private Function P3ApplyLineSearch(ByRef incrementBase() As Double, ByRef targetForce() As Double, ByRef internalForce() As Double, ByRef correctionRatio As Double, ByVal residualBefore As Double) As Boolean
  Dim i As Long, tryCount As Long
  Dim alpha As Double, bestAlpha As Double, bestRes As Double
  Dim trialCorr As Double
  P3ApplyLineSearch = False
  P3LineSearchRetryCorrection = False
  mLineSearchFailure = "": mLineSearchTries = 0: mLineSearchMaterialRejects = 0: mLineSearchBestQ = -1#
  If lastDof < 0 Then
    mLineSearchFailure = "LINESEARCH_NO_DOF"
    Exit Function
  End If
  If P3LineSearchSavedN <> lastDof Then
    ReDim P3LineSearchSavedT(lastDof)
    P3LineSearchSavedN = lastDof
  End If
  For i = 0 To lastDof
    P3LineSearchSavedT(i) = TDisp(i)
  Next i
  alpha = 1#
  bestAlpha = -1#
  bestRes = 1E+308
  For tryCount = 1 To P3_LINESEARCH_MAX
    mLineSearchTries = tryCount
    P6LsNoteTry
    For i = 0 To lastDof
      If NodeCond(i) <> 0 Then
        TDisp(i) = P3LineSearchSavedT(i) + Disp(i)
      Else
        TDisp(i) = P3LineSearchSavedT(i) + alpha * Disp(i)
      End If
      UDisp(i) = TDisp(i)
    Next i
    If Not P3EvaluateTrialState(incrementBase, internalForce) Then
      mLineSearchMaterialRejects = mLineSearchMaterialRejects + 1
      If P3MaterialFailureIsFatal() Then
        mLineSearchFailure = "LINESEARCH_MATERIAL_FATAL"
        For i = 0 To lastDof
          TDisp(i) = P3LineSearchSavedT(i)
          UDisp(i) = P3LineSearchSavedT(i)
        Next i
        P3RestoreCommittedMaterialState False
        Exit Function
      End If
      P3RestoreCommittedMaterialState False
      P3ClearTransientFailure
      alpha = 0.5 * alpha
      P3RequestTangentRebuild "LINESEARCH_CUT"
      P3LineSearchCuts = P3LineSearchCuts + 1
      GoTo NextLineSearchAlpha
    End If
    P3UpdateResidualNormOnly targetForce, internalForce
    If ResidualNormFree < bestRes Then
      bestRes = ResidualNormFree
      If residualBefore > 0# Then mLineSearchBestQ = bestRes / residualBefore
      bestAlpha = alpha
    End If
    If residualBefore <= 0# Or ResidualNormFree <= residualBefore * (1# - 0.0001 * alpha) Or RelativeResidualFree <= P3_ENGINEERING_RESIDUAL * 0.1 Then
      P3UpdateResidualMetrics targetForce, internalForce
      trialCorr = 0#
      P3ApplyCorrectionRatioFromDir alpha, P3LineSearchSavedT, trialCorr
      correctionRatio = trialCorr
      If alpha < 0.999 Then
        P3AcceptDampedLineSearch alpha, residualBefore
        P3LineSearchCuts = P3LineSearchCuts + 1
      Else
        P6LsNoteAccept alpha, False
        P6LsNoteQ residualBefore, ResidualNormFree
      End If
      IncrementLogLineSearchAlpha alpha
      P3ApplyLineSearch = True
      Exit Function
    End If
    P3RestoreCommittedMaterialState False
    alpha = 0.5 * alpha
    P3RequestTangentRebuild "LINESEARCH_CUT"
    P3LineSearchCuts = P3LineSearchCuts + 1
NextLineSearchAlpha:
  Next tryCount
  If bestAlpha > 0# And bestRes < residualBefore Then
    For i = 0 To lastDof
      If NodeCond(i) <> 0 Then
        TDisp(i) = P3LineSearchSavedT(i) + Disp(i)
      Else
        TDisp(i) = P3LineSearchSavedT(i) + bestAlpha * Disp(i)
      End If
      UDisp(i) = TDisp(i)
    Next i
    If Not P3EvaluateTrialState(incrementBase, internalForce) Then
      mLineSearchMaterialRejects = mLineSearchMaterialRejects + 1
      If P3MaterialFailureIsFatal() Then
        mLineSearchFailure = "LINESEARCH_MATERIAL_FATAL"
        For i = 0 To lastDof
          TDisp(i) = P3LineSearchSavedT(i)
          UDisp(i) = P3LineSearchSavedT(i)
        Next i
        P3RestoreCommittedMaterialState False
        Exit Function
      End If
      P3RestoreCommittedMaterialState False
      P3ClearTransientFailure
      mLineSearchFailure = "LINESEARCH_BEST_STATE_REJECT"
      GoTo LineSearchReject
    End If
    P3UpdateResidualMetrics targetForce, internalForce
    If ResidualNormFree < residualBefore Then
      trialCorr = 0#
      P3ApplyCorrectionRatioFromDir bestAlpha, P3LineSearchSavedT, trialCorr
      correctionRatio = trialCorr
      P3AcceptDampedLineSearch bestAlpha, residualBefore
      P3LineSearchCuts = P3LineSearchCuts + 1
      IncrementLogLineSearchAlpha bestAlpha
      P3ApplyLineSearch = True
      Exit Function
    End If
  End If
LineSearchReject:
  If Len(mLineSearchFailure) = 0 Then
    If bestAlpha <= 0# Then mLineSearchFailure = "LINESEARCH_NO_VALID_STATE" Else mLineSearchFailure = "LINESEARCH_NO_REDUCTION"
  End If
  For i = 0 To lastDof
    TDisp(i) = P3LineSearchSavedT(i)
    UDisp(i) = P3LineSearchSavedT(i)
  Next i
  P3RestoreCommittedMaterialState False
  P3ClearTransientFailure
  P3RequestTangentRebuild "LINESEARCH_RETRY"
  P3LineSearchRetryCorrection = True
  P3LineSearchCuts = P3LineSearchCuts + 1
End Function

Private Sub P3ApplyCorrectionRatioFromDir(ByVal alpha As Double, ByRef savedT() As Double, ByRef correctionRatio As Double)
  Dim i As Long, correctionNorm As Double, totalNorm As Double
  correctionNorm = 0#: totalNorm = 0#
  For i = 0 To lastDof
    If NodeCond(i) = 0 Then
      If Abs(alpha * Disp(i)) > correctionNorm Then correctionNorm = Abs(alpha * Disp(i))
      If Abs(TDisp(i)) > totalNorm Then totalNorm = Abs(TDisp(i))
    End If
  Next i
  If totalNorm < P3DispRef Then
    If P3DispRef < 0.000000000001 Then
      correctionRatio = correctionNorm
    Else
      correctionRatio = correctionNorm / P3DispRef
    End If
  Else
    correctionRatio = correctionNorm / totalNorm
  End If
End Sub

Private Function P3RunLoadStages() As Boolean
  Dim targetForce() As Double, incrementBase() As Double, internalForce() As Double
  Dim i As Long, stageId As Long, stageSteps As Long, localIteration As Long
  Dim stageFactor As Double, targetFactor As Double, stepSize As Double, baseStepSize As Double, correctionRatio As Double
  Dim selfWeightFactor As Double, appliedFactor As Double, prescribedFactor As Double
  Dim retryCount As Long, failedError As Long, failedElement As Long, failedGauss As Long
  Dim failedIncrement As Long, failedIteration As Long, failedStatus As String, failedMessage As String
  Dim hasSelfWeight As Boolean, hasAppliedLoad As Boolean, hasPrescribedDisp As Boolean
  Dim selfWeightSquare As Double, appliedSquare As Double, previousDisplacementForceSquare As Double
  Dim converged As Boolean, boundaryOk As Boolean
  Dim wantGravity As Boolean, wantApply As Boolean
  Dim lineSearchRebuildUsed As Boolean
  Dim gravityRan As Boolean, applyRan As Boolean
  Dim duMax As Double
  Dim predictorUsed As Boolean, accelerationRetried As Boolean, correctionBuilt As Boolean
  Dim logStepAction As String, recoveryRatio As Double, boundedNextStep As Double
  P3ApplyNewtonPolicy P3RunLogKind
  P3EvalSkipCount = 0
  P3TangentDirtyRebuildCount = 0
  P3ElasticFastCount = 0
  P3TrialStateValid = False
  If Not P3SrmTrialRunning Then P3LastSuccessStepSize = 0#
  ReDim targetForce(lastDof)
  ReDim incrementBase(lastDof)
  ReDim internalForce(lastDof)
  selfWeightSquare = 0#: appliedSquare = 0#
  For i = 0 To lastDof
    selfWeightSquare = selfWeightSquare + P3SelfWeightForce(i) * P3SelfWeightForce(i)
    appliedSquare = appliedSquare + P3AppliedForce(i) * P3AppliedForce(i)
  Next i
  hasSelfWeight = (selfWeightSquare > 1E-30)
  hasAppliedLoad = (appliedSquare > 1E-30)
  hasPrescribedDisp = P3HasNonzeroPrescribedDisp()
  P3ForceRef = Sqr(selfWeightSquare) + Sqr(appliedSquare)
  If P3ForceRef <= 0.000000000001 And P3CommittedBoundaryReady Then
    ' Retain the committed displacement-load reaction scale during unloading.
    ' A kinematic reset has zero committed displacement and must not reuse it.
    For i = 0 To lastDof
      If P3CommittedNodeCond(i) <> 0 And Abs(P3CommittedBoundaryDisp(i)) > 1E-30 And Abs(P3CommittedDisp(i)) > 1E-30 Then
        previousDisplacementForceSquare = previousDisplacementForceSquare + P3FinalInternalForce(i) * P3FinalInternalForce(i)
      End If
    Next i
    If previousDisplacementForceSquare > 0# Then P3ForceRef = Sqr(previousDisplacementForceSquare)
  End If
  If P3ForceRef < 0.000000000001 Then P3ForceRef = 0.000000000001
  P3DispRef = P5ModelScale * 0.000001
  If P3DispRef < 0.000000000001 Then P3DispRef = 0.000000000001
  P3PrevResidualNorm = 0#
  P3RequestTangentRebuild "STAGE_START"
  P3TangentStaleIters = 0
  P3ModifiedNewtonSkips = 0
  P3LineSearchCuts = 0
  P3JointAssembleCount = 0
  wantGravity = P3PlanWantsKind("GRAVITY")
  wantApply = P3PlanWantsKind("APPLY")
  gravityRan = False
  applyRan = False
  P3RunLoadStages = False
  For stageId = 1 To 2
    If stageId = 1 Then
      If Not wantGravity Then GoTo P3NextStage
      If Not hasSelfWeight And Not P3VectorHasChange(P3LockedSelfWeight, P3SelfWeightForce) Then GoTo P3NextStage
      stepSize = P3_GRAVITY_STEP_SIZE
    Else
      If Not wantApply Then GoTo P3NextStage
      If Not hasAppliedLoad And Not hasPrescribedDisp And Not P3VectorHasChange(P3LockedApplied, P3AppliedForce) And Not P3PrescribedBoundaryChanged() Then GoTo P3NextStage
      stepSize = P3_APPLY_STEP_SIZE
      If P3SrmTrialRunning And P3LastSuccessStepSize >= P3_APPLY_STEP_SIZE Then
        stepSize = P3LastSuccessStepSize
      End If
    End If
    If stepSize > P3_MAX_STEP_SIZE Then stepSize = P3_MAX_STEP_SIZE
    If stepSize < P3_MIN_STEP_FACTOR Then stepSize = P3_MIN_STEP_FACTOR
    AccelResetStage
    P3AccelResetAA
    If AccelV1 Or AccelV2a Or AccelV2b Or AccelV4 Or AccelV5 Then P3RequestTangentRebuild "STAGE_START"
    P3CaptureStageStartDisp
    stageFactor = 0#
    baseStepSize = stepSize
    StepRecoveryResetStage baseStepSize
    retryCount = 0
    Do While stageFactor < 1# - 0.000000000001
      targetFactor = stageFactor + stepSize
      If targetFactor > 1# Then targetFactor = 1#
      P6PhysLambda = targetFactor
      P6PhysStep = stepSize
      prescribedFactor = P3PrescribedFactor(stageId, targetFactor)
      accelerationRetried = False
      AccelBaselineRetry = False
      AdaptiveBeginIncrement stageFactor, targetFactor
P3RetrySameIncrement:
      StepRecoveryBeginAttempt
      IncrementLogBegin P3SuccessfulIncrementCount + 1, stageFactor, targetFactor, stepSize, retryCount, AccelBaselineRetry
      P3AccelResetAA
      AccelFirstReusePending = False
      If accelerationRetried Then P3TrialStateValid = False
      P3CopyCommittedDispToTrial
      For i = 0 To lastDof
        incrementBase(i) = P3CommittedDisp(i)
      Next i
      P3ApplyIncrementPredictor targetFactor - stageFactor, prescribedFactor, predictorUsed
      IncrementLogPredictorUsed predictorUsed
      If P6MixedUP Then P6UpdateContinuityG
      If stageId = 1 Then
        selfWeightFactor = targetFactor
        appliedFactor = 0#
      Else
        selfWeightFactor = 0#
        appliedFactor = targetFactor
      End If
      P3BuildTargetForce selfWeightFactor, appliedFactor, targetForce
      P3ClearTransientFailure
      mAccelAttemptFailure = "": mLineSearchFailure = "": mLineSearchTries = 0: mLineSearchMaterialRejects = 0: mLineSearchBestQ = -1#
      localIteration = 0: correctionRatio = 0#: converged = False
      P6LsWatch = False
      P3PrevResidualNorm = 0#
      P3TangentStaleIters = 0
      lineSearchRebuildUsed = False
      Do
        AccelBeginIteration
        CurrentIncrement = P3SuccessfulIncrementCount + 1
        CurrentIteration = localIteration
        If P3TrialStateValid Then
          P3EvalSkipCount = P3EvalSkipCount + 1
        Else
          If Not P3EvaluateTrialState(incrementBase, internalForce) Then
            mAccelAttemptFailure = "TRIAL_STATE_FAILURE"
            Exit Do
          End If
        End If
        P3UpdateResidualMetrics targetForce, internalForce
        P3UpdateBoundaryDispError prescribedFactor
        boundaryOk = P3BoundaryDispSatisfied()
        If RelativeResidualFree <= P3_ENGINEERING_RESIDUAL And boundaryOk Then
          If localIteration = 0 Then
            If Not hasPrescribedDisp Or P3MaxBoundaryDispError <= 0.000000000001 Then
              converged = True
              If P3TangentAge = 1 Then P6Age1Success = P6Age1Success + 1
              Exit Do
            End If
          ElseIf correctionRatio <= P3_ENGINEERING_CORRECTION Or RelativeResidualFree <= P3_RESIDUAL_TOLERANCE Then
            converged = True
            If P3TangentAge = 1 Then P6Age1Success = P6Age1Success + 1
            Exit Do
          End If
        End If
        If localIteration >= P3_MAX_GLOBAL_ITERATIONS Then
          mAccelAttemptFailure = "GLOBAL_ITERATION_LIMIT"
          SetAnalysisFailure RESULT_NONCONVERGED, "P3全体反復が上限回数に達しました。Ver=" & FEM_BUILD_STAMP & " 相対残差=" & Format$(RelativeResidualFree, "0.000E+00") & " 補正比=" & Format$(correctionRatio, "0.000E+00") & " 拘束変位誤差=" & Format$(P3MaxBoundaryDispError, "0.000E+00") & "。状態は前増分へ戻します。", vbObjectError + 3202, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
          Exit Do
        End If
        For i = 0 To lastDof
          Disp(i) = 0#
          If NodeCond(i) <> 0 Then Disp(i) = P3TargetBoundaryDisp(i, prescribedFactor) - TDisp(i)
          Force(i) = targetForce(i) - internalForce(i)
        Next i
        correctionBuilt = False
        If P3ShouldRebuildTangent(localIteration, ResidualNormFree) Then
          correctionBuilt = True
          P6NoteAgeOutcome True
          P6PhysOnRebuild
          If Not P3RebuildTangentFromSpmat() Then
            mAccelAttemptFailure = "TANGENT_BUILD_FAILURE"
            Exit Do
          End If
        ElseIf P3JointContactChangedSinceFactor() Then
          P3TangentJustRebuilt = True
          P6FactorReady = False
          SetTotalMat
          P3JointAssembleCount = P3JointAssembleCount + 1
        Else
          P6NoteAgeOutcome False
          P3ModifiedNewtonSkips = P3ModifiedNewtonSkips + 1
          P6PerfModNewtonSkips = P6PerfModNewtonSkips + 1
        End If
        SetBoundaryCondition
        If Not BandSolver() Then
          mAccelAttemptFailure = "BAND_SOLVE_FAILURE"
          Exit Do
        End If
        P3NoteSolveAge
        P3GlobalIterationCount = P3GlobalIterationCount + 1
        P6PerfNewtonCount = P6PerfNewtonCount + 1
        P3PrevResidualNorm = ResidualNormFree
        If Not P3AccelApplyCorrection(incrementBase, targetForce, internalForce, correctionRatio, P3PrevResidualNorm) Then
          If Len(mLineSearchFailure) > 0 Then mAccelAttemptFailure = mLineSearchFailure
          If Len(mAccelAttemptFailure) = 0 Then mAccelAttemptFailure = P3AccelFailureReason()
          StepRecoveryRejectedCorrection mAccelAttemptFailure
          IncrementLogNoteCorrectionFailure mAccelAttemptFailure
          AdaptiveV2aRejectedCorrection mAccelAttemptFailure
          If P3LineSearchRetryCorrection And Not lineSearchRebuildUsed Then
            lineSearchRebuildUsed = True
            P3RequestTangentRebuild "LINESEARCH_RETRY"
            localIteration = localIteration + 1
            If localIteration >= P3_MAX_GLOBAL_ITERATIONS Then
              mAccelAttemptFailure = "GLOBAL_ITERATION_LIMIT_AFTER_LINESEARCH_RETRY"
              SetAnalysisFailure RESULT_NONCONVERGED, "P3全体反復が上限回数に達しました。Ver=" & FEM_BUILD_STAMP & " 相対残差=" & Format$(RelativeResidualFree, "0.000E+00") & " 補正比=" & Format$(correctionRatio, "0.000E+00") & " 拘束変位誤差=" & Format$(P3MaxBoundaryDispError, "0.000E+00") & "。状態は前増分へ戻します。", vbObjectError + 3202, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
              Exit Do
            End If
          Else
            Exit Do
          End If
        Else
          AccelNoteCorrection P3PrevResidualNorm, ResidualNormFree, correctionBuilt
          lineSearchRebuildUsed = False
          If correctionRatio > P3MaxCorrection Then P3MaxCorrection = correctionRatio
          localIteration = localIteration + 1
          If femIoMode = "ITER" Or P3GlobalIterationCount <= 12 Then P3TraceIteration "correction"
        End If
      Loop
      IncrementLogCorrectionRatio correctionRatio, (localIteration > 0)
      If converged Then
        recoveryRatio = StepRecoveryAccepted(localIteration, targetFactor - stageFactor, stepSize, targetFactor, _
          correctionRatio, retryCount, accelerationRetried, boundaryOk, AdaptiveMethodLossThisIncrement())
        AdaptiveEndIncrement True, localIteration, accelerationRetried
        AccelRememberIncrement incrementBase, targetFactor - stageFactor, localIteration, retryCount
        If AccelTrace Then P6SolverEvent "INCREMENT_ACCEPT", "fs=" & Format$(P3CurrentStrengthFactor, "0.000000000000") & ";lambda_start=" & Format$(stageFactor, "0.000000000000") & ";lambda_target=" & Format$(targetFactor, "0.000000000000") & ";delta_lambda=" & Format$(targetFactor - stageFactor, "0.000000000000") & ";iterations=" & CStr(localIteration) & ";retries=" & CStr(retryCount) & ";predictor=" & CStr(predictorUsed) & ";baseline_retry=" & CStr(AccelBaselineRetry) & ";relres=" & Format$(RelativeResidualFree, "0.000E+00")
        duMax = 0#
        For i = 0 To lastDof
          If NodeCond(i) = 0 Then
            If Abs(TDisp(i) - P3CommittedDisp(i)) > duMax Then duMax = Abs(TDisp(i) - P3CommittedDisp(i))
          End If
          P3FinalInternalForce(i) = internalForce(i)
          P3CommittedDisp(i) = TDisp(i)
        Next i
        P3CommitMaterialState
        PlasticLogAccepted
        P3SuccessfulIncrementCount = P3SuccessfulIncrementCount + 1
        P6PerfIncrementCount = P6PerfIncrementCount + 1
        P3LastConvergedIncrement = P3SuccessfulIncrementCount
        P3Ls50Pending = False
        P3Ls50Watch = False
        P3AfterCutback = False
        P6RollPush targetFactor, duMax
        If (P3SuccessfulIncrementCount Mod 20) = 0 Then
          P3PerfTrace "INC", stepSize, retryCount, duMax, targetFactor, targetFactor
        Else
          P6PhysWatchIncrement stepSize, duMax, targetFactor
        End If
        If stageId = 2 Then P3LastConvergedLoadFactor = targetFactor
        stageFactor = targetFactor
        retryCount = 0
        If localIteration <= 2 Then
          stepSize = stepSize * 1.5
          logStepAction = "GROW_1_5"
        ElseIf localIteration <= 5 Then
          stepSize = stepSize
          logStepAction = "KEEP"
          If recoveryRatio > 1# Then
            stepSize = stepSize * recoveryRatio
            logStepAction = "RECOVER_1_1"
          End If
        ElseIf localIteration <= 8 Then
          stepSize = stepSize * 0.7
          logStepAction = "SHRINK_0_7"
        Else
          stepSize = stepSize * 0.5
          logStepAction = "SHRINK_0_5"
        End If
        boundedNextStep = StepRecoveryBoundNextStep(stepSize)
        If boundedNextStep < stepSize Then
          stepSize = boundedNextStep
          logStepAction = logStepAction & "|RECOVERY_RESTORE"
        End If
        If stepSize > P3_MAX_STEP_SIZE Then
          stepSize = P3_MAX_STEP_SIZE
          logStepAction = logStepAction & "|MAX_CAP"
        End If
        If stepSize > 1# - stageFactor And (1# - stageFactor) > 0# Then
          stepSize = 1# - stageFactor
          logStepAction = logStepAction & "|REMAINING_CAP"
        End If
        If stepSize < P3_MIN_STEP_FACTOR Then
          stepSize = P3_MIN_STEP_FACTOR
          logStepAction = logStepAction & "|MIN_FLOOR"
        End If
        StepRecoveryNoteNextStep logStepAction, stepSize
        P3LastSuccessStepSize = stepSize
        If stageFactor >= 1# - 0.000000000001 Then
          IncrementLogFinish True, localIteration, "ACCEPT", -1#, logStepAction & "|STAGE_COMPLETE", ""
        Else
          IncrementLogFinish True, localIteration, "ACCEPT", stepSize, logStepAction, ""
        End If
        If SuppressUserMessages Then SaveP0Progress "p3_" & CStr(stageId) & "_increment_" & CStr(P3SuccessfulIncrementCount)
      Else
        ' Acceleration failures do not establish SRM FAIL. Retry the same target
        ' from committed history with all optional solver features disabled.
        StepRecoveryFailedAttempt P3AccelFailureReason()
        If Not accelerationRetried And (predictorUsed Or StepRecoveryAttemptProbe() Or StepRecoveryAttemptObserved() Or _
          ((AccelV2a Or AccelV2b Or AccelV4 Or AccelV5) And Not AdaptiveEnabled) Or _
          (AdaptiveEnabled And AdaptiveIncrementUsed())) And Not P3IsFatalFailure() Then
          P6SolverEvent "ACCEL_FALLBACK", "same_increment=True;lambda_start=" & Format$(stageFactor, "0.000000000000") & ";lambda_target=" & Format$(targetFactor, "0.000000000000") & ";predictor=" & CStr(predictorUsed) & ";failure_status=" & ResultStatus & ";failure_error=" & CStr(AnalysisErrorNumber) & ";failure_reason=" & P3AccelFailureReason() & ";line_search_retry=" & CStr(P3LineSearchRetryCorrection) & ";line_search_tries=" & CStr(mLineSearchTries) & ";line_search_material_rejects=" & CStr(mLineSearchMaterialRejects) & ";line_search_best_q=" & Format$(mLineSearchBestQ, "0.000000")
          AdaptiveFailedAttempt
          IncrementLogFinish False, localIteration, "BASELINE_RETRY", stepSize, "RETRY_SAME_TARGET", P3AccelFailureReason()
          accelerationRetried = True
          AccelFallbackCount = AccelFallbackCount + 1
          P3RestoreCommittedMaterialState
          AccelHistoryReady = False: AccelIncrementEligible = False: AccelV2aPreviousIterations = 0: AccelV2aPreviousStep = 0#
          AccelBaselineRetry = True
          P3ClearTransientFailure
          GoTo P3RetrySameIncrement
        End If
        AdaptiveEndIncrement False, localIteration, accelerationRetried
        AccelHistoryReady = False: AccelIncrementEligible = False: AccelV2aPreviousIterations = 0: AccelV2aPreviousStep = 0#
        P3AccelResetAA
        failedStatus = ResultStatus
        failedMessage = AnalysisMessage
        failedError = AnalysisErrorNumber
        failedElement = FailureElement
        failedGauss = FailureGaussPoint
        failedIncrement = CurrentIncrement
        failedIteration = CurrentIteration
        P3RestoreCommittedMaterialState
        P3CopyCommittedDispToTrial
        retryCount = retryCount + 1
        P3RetryCount = P3RetryCount + 1
        P6PerfCutbackCount = P6PerfCutbackCount + 1
        P3PerfTrace "CUTBACK", stepSize, retryCount, 0#, stageFactor, targetFactor
        If SuppressUserMessages Then SaveP0Progress "p3_retry_" & CStr(P3RetryCount)
        If P3UserCancel Or P3SearchFatal Or failedError = 18 Then
          StepRecoveryTerminateWindow
          IncrementLogFinish False, localIteration, "ABORT", -1#, "USER_CANCEL_OR_FATAL", P3AccelFailureReason()
          SetAnalysisFailure failedStatus, failedMessage, failedError, failedElement, failedGauss, failedIncrement, failedIteration
          Exit Function
        End If
        If retryCount > P3_MAX_STEP_RETRIES Or stepSize * 0.5 < P3_MIN_STEP_FACTOR Then
          StepRecoveryTerminateWindow
          IncrementLogFinish False, localIteration, "FINAL_FAILURE", -1#, "RETRY_OR_MIN_STEP_LIMIT", P3AccelFailureReason()
          If failedStatus = RESULT_MATERIAL_ERROR Or failedStatus = RESULT_GLOBAL_SINGULAR Then
            SetAnalysisFailure failedStatus, failedMessage, failedError, failedElement, failedGauss, failedIncrement, failedIteration
          Else
            If Len(failedMessage) = 0 Then failedMessage = "P3荷重増分が収束せず、最小増分に達しました。"
            SetAnalysisFailure RESULT_NONCONVERGED, failedMessage, vbObjectError + 3203, failedElement, failedGauss, failedIncrement, failedIteration
          End If
          Exit Function
        End If
        stepSize = stepSize * 0.5
        StepRecoveryCutback stepSize
        IncrementLogFinish False, localIteration, "CUTBACK", stepSize, "CUTBACK", P3AccelFailureReason()
      End If
    Loop
      If stageId = 1 Then
        gravityRan = True
        P3GravityCommitted = True
      Else
        applyRan = True
      End If
P3NextStage:
  Next stageId
  If P6ConsolActive And Not P3SrmEnabled Then
    If Not P3RunConsolidationSteps(targetForce, incrementBase, internalForce) Then Exit Function
  End If
  P3CommitLoadLock gravityRan, applyRan
  P3UseNonlinearResidual = True
  P3RunLoadStages = True
End Function

Private Function P3RunConsolidationSteps(ByRef targetForce() As Double, ByRef incrementBase() As Double, ByRef internalForce() As Double) As Boolean
  Dim i As Long, stepId As Long, stepCount As Long, localIteration As Long
  Dim correctionRatio As Double, converged As Boolean
  P3RunConsolidationSteps = False
  stepCount = CLng(P6ReadSetting("CONSOL_NSTEP", 10#))
  If stepCount < 1 Then
    P3RunConsolidationSteps = True
    Exit Function
  End If
  P6ConsolDt = P6ReadSetting("CONSOL_DT", 86400#)
  If P6ConsolDt < 0# Then P6ConsolDt = 0#
  P3BuildTargetForce 0#, 0#, targetForce
  P6BuildBiotCapacity
  P6RebuildBandWithSchur
  For stepId = 1 To stepCount
    P6ConsolStep = stepId
    P6ConsolTime = P6ConsolTime + P6ConsolDt
    P3CopyCommittedDispToTrial
    For i = 0 To lastDof
      incrementBase(i) = P3CommittedDisp(i)
    Next i
    P6UpdateContinuityG
    P3ClearTransientFailure
    localIteration = 0: correctionRatio = 0#: converged = False
    Do
      CurrentIncrement = P3SuccessfulIncrementCount + 1
      CurrentIteration = localIteration
      If Not P3EvaluateTrialState(incrementBase, internalForce) Then Exit Do
      P3UpdateResidualMetrics targetForce, internalForce
      If RelativeResidualFree <= P3_ENGINEERING_RESIDUAL Then
        If localIteration = 0 Or correctionRatio <= P3_ENGINEERING_CORRECTION Or RelativeResidualFree <= P3_RESIDUAL_TOLERANCE Then
          converged = True
          Exit Do
        End If
        If localIteration >= 2 Then
          converged = True
          Exit Do
        End If
      End If
      If localIteration >= P3_MAX_GLOBAL_ITERATIONS Then
        SetAnalysisFailure RESULT_NONCONVERGED, "圧密ステップが上限回数に達しました。Ver=" & FEM_BUILD_STAMP & " t=" & Format$(P6ConsolTime, "0.000E+00") & " 相対残差=" & Format$(RelativeResidualFree, "0.000E+00"), vbObjectError + 3202, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
        Exit Do
      End If
      For i = 0 To lastDof
        Disp(i) = 0#
        If NodeCond(i) <> 0 Then Disp(i) = P3BoundaryDisp(i) - TDisp(i)
        Force(i) = targetForce(i) - internalForce(i)
      Next i
      SetBoundaryCondition
      If Not BandSolver() Then Exit Do
      P3GlobalIterationCount = P3GlobalIterationCount + 1
      P6PerfNewtonCount = P6PerfNewtonCount + 1
      P3ApplyCorrection correctionRatio
      If correctionRatio > P3MaxCorrection Then P3MaxCorrection = correctionRatio
      localIteration = localIteration + 1
    Loop
    If Not converged Then
      P3RestoreCommittedMaterialState
      P3CopyCommittedDispToTrial
      Exit Function
    End If
    For i = 0 To lastDof
      P3FinalInternalForce(i) = internalForce(i)
      P3CommittedDisp(i) = TDisp(i)
    Next i
    P3CommitMaterialState
    PlasticLogAccepted
    P3SuccessfulIncrementCount = P3SuccessfulIncrementCount + 1
    P6PerfIncrementCount = P6PerfIncrementCount + 1
    P3LastConvergedIncrement = P3SuccessfulIncrementCount
    If SuppressUserMessages Then SaveP0Progress "consol_step_" & CStr(stepId)
  Next stepId
  P6MaxPressure = 0#
  If P6PressureCount > 0 Then
    For i = 0 To P6PressureCount - 1
      If Abs(P6Pressure(i)) > P6MaxPressure Then P6MaxPressure = Abs(P6Pressure(i))
    Next i
  End If
  P3RunConsolidationSteps = True
End Function

Private Function P3NormalizeStageKind(ByVal rawKind As String) As String
  Dim kindText As String
  kindText = UCase$(Trim$(rawKind))
  If kindText = "自重" Or kindText = "GRAVITY" Or kindText = "WEIGHT" Then
    P3NormalizeStageKind = "GRAVITY"
  ElseIf kindText = "載荷" Or kindText = "APPLY" Or kindText = "LOAD" Or kindText = "DISP" Then
    P3NormalizeStageKind = "APPLY"
  ElseIf kindText = "除荷" Or kindText = "UNLOAD" Then
    P3NormalizeStageKind = "UNLOAD"
  ElseIf kindText = "強度低減" Or kindText = "SRM" Or kindText = "FS" Then
    P3NormalizeStageKind = "SRM"
  ElseIf kindText = "BIRTH" Or kindText = "誕生" Or kindText = "盛土" Then
    P3NormalizeStageKind = "BIRTH"
  ElseIf kindText = "DEATH" Or kindText = "死滅" Or kindText = "掘削" Or kindText = "KILL" Then
    P3NormalizeStageKind = "DEATH"
  ElseIf kindText = "RESET_U" Or kindText = "変位リセット" Then
    P3NormalizeStageKind = "RESET_U"
  ElseIf kindText = "RESET_STRESS" Or kindText = "応力リセット" Then
    P3NormalizeStageKind = "RESET_STRESS"
  ElseIf kindText = "MATSET" Or kindText = "材料変更" Then
    P3NormalizeStageKind = "MATSET"
  Else
    P3NormalizeStageKind = kindText
  End If
End Function

Private Sub P3LoadStagePlan()
  Dim ws As Worksheet, lastRow As Long, rowNo As Long, kindText As String
  Dim enabledValue As Double
  Dim planText As String
  P3SrmEnabled = False
  P3PlanHasGravity = True
  P3PlanHasApply = True
  P3StagePlanText = "GRAVITY+APPLY"
  P3StageN = 0
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("ステージ")
  On Error GoTo 0
  If ws Is Nothing Then Exit Sub
  lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
  If lastRow < 2 Then Exit Sub
  ReDim P3StageKind(1 To lastRow)
  ReDim P3StageId(1 To lastRow)
  ReDim P3StageGroup(1 To lastRow)
  ReDim P3StageParam(1 To lastRow)
  ReDim P3StageOn(1 To lastRow)
  For rowNo = 2 To lastRow
    kindText = P3NormalizeStageKind(CStr(ws.Cells(rowNo, 2).value2))
    If Len(kindText) = 0 Then GoTo NextStageRow
    enabledValue = 0#
    If IsNumeric(ws.Cells(rowNo, 5).value2) Then enabledValue = CDbl(ws.Cells(rowNo, 5).value2)
    P3StageN = P3StageN + 1
    P3StageKind(P3StageN) = kindText
    If IsNumeric(ws.Cells(rowNo, 1).value2) Then
      P3StageId(P3StageN) = CLng(Val(CStr(ws.Cells(rowNo, 1).value2)))
    Else
      P3StageId(P3StageN) = P3StageN
    End If
    If P3StageId(P3StageN) <= 0 Then P3StageId(P3StageN) = P3StageN
    P3StageGroup(P3StageN) = Trim$(CStr(ws.Cells(rowNo, 3).value2))
    P3StageParam(P3StageN) = Trim$(CStr(ws.Cells(rowNo, 4).value2))
    P3StageOn(P3StageN) = (enabledValue <> 0#)
    If P3StageOn(P3StageN) Then
      If kindText <> "GRAVITY" And kindText <> "APPLY" And kindText <> "LOAD" And kindText <> "UNLOAD" And kindText <> "SRM" And kindText <> "BIRTH" And kindText <> "DEATH" And kindText <> "RESET_U" And kindText <> "RESET_STRESS" And kindText <> "MATSET" Then
        SetAnalysisFailure RESULT_INPUT_ERROR, "未対応のステージ種別です。行=" & CStr(rowNo) & " 種別=" & kindText, vbObjectError + 3233, -1, -1, CurrentIncrement, CurrentIteration
        Exit Sub
      End If
      If Len(planText) > 0 Then planText = planText & "+"
      planText = planText & kindText
      If kindText = "SRM" Then P3SrmEnabled = True
    End If
NextStageRow:
  Next rowNo
  If P3StageN = 0 Then Exit Sub
  P3FlagsFromPrefix P3StageN
  If Len(planText) > 0 Then P3StagePlanText = planText
End Sub

Private Sub P3FlagsFromPrefix(ByVal lastIndex As Long)
  Dim i As Long
  P3PlanHasGravity = False
  P3PlanHasApply = False
  If lastIndex > P3StageN Then lastIndex = P3StageN
  If lastIndex < 0 Then lastIndex = 0
  P3ActivePrefix = lastIndex
  For i = 1 To lastIndex
    If P3StageOn(i) Then
      If P3StageKind(i) = "GRAVITY" Then P3PlanHasGravity = True
      If P3StageKind(i) = "APPLY" Or P3StageKind(i) = "LOAD" Then P3PlanHasApply = True
    End If
  Next i
  If P3StageN = 0 Then
    P3PlanHasGravity = True
    P3PlanHasApply = True
  End If
End Sub

Private Function P3StageGroupHitsMaterial(ByVal groupText As String, ByVal materialId As Long) As Boolean
  Dim token As String
  token = Trim$(groupText)
  If Len(token) = 0 Then
    P3StageGroupHitsMaterial = True
    Exit Function
  End If
  If Not IsNumeric(token) Then
    P3StageGroupHitsMaterial = True
    Exit Function
  End If
  P3StageGroupHitsMaterial = (CLng(Val(token)) = materialId)
End Function

Private Function P3MaterialNetLoaded(ByVal materialId As Long) As Boolean
  Dim i As Long, lastIndex As Long, loaded As Boolean, sawLoad As Boolean
  If P3StageN <= 0 Then
    P3MaterialNetLoaded = True
    Exit Function
  End If
  lastIndex = P3ActivePrefix
  If lastIndex <= 0 Then lastIndex = P3StageN
  loaded = False
  For i = 1 To lastIndex
    If Not P3StageOn(i) Then GoTo NextNetRow
    If P3StageKind(i) = "APPLY" Or P3StageKind(i) = "LOAD" Then
      If P3StageGroupHitsMaterial(P3StageGroup(i), materialId) Then
        loaded = True
        sawLoad = True
      End If
    ElseIf P3StageKind(i) = "UNLOAD" Then
      If P3StageGroupHitsMaterial(P3StageGroup(i), materialId) Then
        loaded = False
        sawLoad = True
      End If
    End If
NextNetRow:
  Next i
  If sawLoad Then
    P3MaterialNetLoaded = loaded
  Else
    P3MaterialNetLoaded = False
  End If
End Function

Public Function P3ResultStageNumber() As Long
  Dim i As Long, n As Long
  If P3LastCompletedStage > 0 Then
    P3ResultStageNumber = P3LastCompletedStage
    Exit Function
  End If
  n = 0
  For i = 1 To P3StageN
    If P3StageOn(i) Then n = i
  Next i
  If n <= 0 Then n = 1
  P3ResultStageNumber = n
End Function

Private Function P3FirstSrmIndex() As Long
  Dim i As Long
  P3FirstSrmIndex = 0
  For i = 1 To P3StageN
    If P3StageOn(i) And P3StageKind(i) = "SRM" Then
      P3FirstSrmIndex = i
      Exit Function
    End If
  Next i
End Function

Private Function P3HasStageAfter(ByVal stageIndex As Long) As Boolean
  Dim i As Long
  P3HasStageAfter = False
  For i = stageIndex + 1 To P3StageN
    If P3StageOn(i) Then
      P3HasStageAfter = True
      Exit Function
    End If
  Next i
End Function

Private Sub P3ResetStageResultSheet()
  Dim ws As Worksheet, lastRow As Long
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("ステージ結果")
  On Error GoTo 0
  If ws Is Nothing Then Exit Sub
  lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
  If lastRow < 2 Then lastRow = 2
  ws.range(ws.Cells(2, 1), ws.Cells(lastRow, 10)).ClearContents
  P3StageResultRow = 1
End Sub

Private Sub P3LogStageResult(ByVal stageNo As Long, ByVal kindText As String, ByVal groupText As String, ByVal okFlag As Boolean, ByVal factorValue As Double, ByVal messageText As String, Optional ByVal exportCheckpoint As Boolean = True)
  Dim ws As Worksheet
  If P3ReplayQuiet Then Exit Sub
  p3StageResultLogged = True
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("ステージ結果")
  On Error GoTo 0
  If Not ws Is Nothing Then
    P3StageResultRow = P3StageResultRow + 1
    ws.Cells(P3StageResultRow, 1).value2 = stageNo
    ws.Cells(P3StageResultRow, 2).value2 = kindText
    ws.Cells(P3StageResultRow, 3).value2 = groupText
    If okFlag Then
      ws.Cells(P3StageResultRow, 4).value2 = "PASS"
    Else
      ws.Cells(P3StageResultRow, 4).value2 = "FAIL"
    End If
    ws.Cells(P3StageResultRow, 5).value2 = factorValue
    ws.Cells(P3StageResultRow, 6).value2 = P3MaxTrialDisp
    ws.Cells(P3StageResultRow, 7).value2 = P6ReactionSumX
    ws.Cells(P3StageResultRow, 8).value2 = P6ReactionSumY
    ws.Cells(P3StageResultRow, 9).value2 = P3SrmTrialCount
    ws.Cells(P3StageResultRow, 10).value2 = messageText
  End If
  If stageNo > 0 And okFlag Then P3LastCompletedStage = stageNo
  If exportCheckpoint Then FEMIoExportStage stageNo, kindText, okFlag
End Sub

Private Function P3LoadingSheetHasAction() As Boolean
  Dim ws As Worksheet, lastRow As Long, rowNo As Long
  Dim kindText As String, valueText As Double
  P3LoadingSheetHasAction = False
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("載荷")
  On Error GoTo 0
  If ws Is Nothing Then Exit Function
  If UCase$(Trim$(CStr(ws.Cells(1, 1).value2))) <> "材料番号" And UCase$(Trim$(CStr(ws.Cells(1, 1).value2))) <> "グループ" Then Exit Function
  lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
  For rowNo = 2 To lastRow
    kindText = UCase$(Trim$(CStr(ws.Cells(rowNo, 4).value2)))
    If kindText = "FIX" Or kindText = "DISP" Then
      P3LoadingSheetHasAction = True
      Exit Function
    End If
    If kindText = "FORCE" Then
      P3LoadingSheetHasAction = True
      Exit Function
    End If
  Next rowNo
End Function

Private Function P3NodeMatchesSelector(ByVal originalNode As Long, ByVal selectorText As String, ByVal xmin As Double, ByVal xmax As Double, ByVal yMin As Double, ByVal yMax As Double) As Boolean
  Dim token As String, boxParts() As String, bx1 As Double, by1 As Double, bx2 As Double, by2 As Double
  Dim nodeX As Double, nodeY As Double, tol As Double
  P3NodeMatchesSelector = False
  token = UCase$(Trim$(selectorText))
  If originalNode < LBound(x) Or originalNode > UBound(x) Then Exit Function
  nodeX = x(originalNode): nodeY = y(originalNode)
  tol = 0.000000001 * (Abs(xmax - xmin) + Abs(yMax - yMin) + 1#)
  If IsNumeric(token) Then
    If CLng(Val(token)) = originalNode + 1 Then P3NodeMatchesSelector = True
    Exit Function
  End If
  If token = "TOP" Then
    P3NodeMatchesSelector = (Abs(nodeY - yMax) <= tol)
  ElseIf token = "BOTTOM" Then
    P3NodeMatchesSelector = (Abs(nodeY - yMin) <= tol)
  ElseIf token = "LEFT" Then
    P3NodeMatchesSelector = (Abs(nodeX - xmin) <= tol)
  ElseIf token = "RIGHT" Then
    P3NodeMatchesSelector = (Abs(nodeX - xmax) <= tol)
  ElseIf left$(token, 4) = "BOX:" Then
    boxParts = Split(mid$(token, 5), ",")
    If UBound(boxParts) >= 3 Then
      bx1 = CDbl(Val(boxParts(0))): by1 = CDbl(Val(boxParts(1)))
      bx2 = CDbl(Val(boxParts(2))): by2 = CDbl(Val(boxParts(3)))
      If nodeX >= bx1 - tol And nodeX <= bx2 + tol And nodeY >= by1 - tol And nodeY <= by2 + tol Then P3NodeMatchesSelector = True
    End If
  End If
End Function

Private Function P3NodeInBox(ByVal originalNode As Long, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Boolean
  Dim nodeX As Double, nodeY As Double, leftX As Double, rightX As Double, bottomY As Double, topY As Double, tol As Double
  P3NodeInBox = False
  If originalNode < LBound(x) Or originalNode > UBound(x) Then Exit Function
  nodeX = x(originalNode): nodeY = y(originalNode)
  If x1 < x2 Then leftX = x1: rightX = x2 Else leftX = x2: rightX = x1
  If y1 < y2 Then bottomY = y1: topY = y2 Else bottomY = y2: topY = y1
  tol = 0.000000001 * (Abs(rightX - leftX) + Abs(topY - bottomY) + 1#)
  P3NodeInBox = (nodeX >= leftX - tol And nodeX <= rightX + tol And nodeY >= bottomY - tol And nodeY <= topY + tol)
End Function

Private Function P3NodeOnMaterial(ByVal originalNode As Long, ByVal materialId As Long) As Boolean
  Dim elementId As Long, j As Long, freeNode As Long
  P3NodeOnMaterial = False
  If P3LoadIndexReady And P3LoadIndexActiveGen = P3ActiveSetGen Then
    If originalNode >= 0 And originalNode <= P3LoadIndexNodeN And materialId >= 0 And materialId <= P3LoadIndexMatN Then
      P3NodeOnMaterial = P3NodeOnMatCache(originalNode, materialId)
    End If
    Exit Function
  End If
  If originalNode < LBound(FEMFreeNodeMap) Or originalNode > UBound(FEMFreeNodeMap) Then Exit Function
  freeNode = FEMFreeNodeMap(originalNode)
  If freeNode < 0 Then Exit Function
  freeNode = P6GetInternalFreeNode(freeNode)
  If freeNode < 0 Then Exit Function
  If materialId <= 0 Then
    P3NodeOnMaterial = True
    Exit Function
  End If
  If materialId > NumberOfMaterial Then Exit Function
  For elementId = 0 To NumberOfElement - 1
    If Not P3IsElementActive(elementId) Then GoTo NextNodeMaterialElement
    If Elem(elementId).MatNo = materialId - 1 Then
      For j = 0 To 7
        If Elem(elementId).node(j) = freeNode Then
          P3NodeOnMaterial = True
          Exit Function
        End If
      Next j
    End If
NextNodeMaterialElement:
  Next elementId
End Function

Private Sub P3BuildLoadIndex()
  Dim orig As Long, elementId As Long, j As Long, materialId As Long
  Dim freeNode As Long, origFree As Long, origNode As Long
  Dim invFree() As Long
  If NumberOfNode < 1 Then
    P3LoadIndexReady = False
    Exit Sub
  End If
  ReDim P3NodeOnMatCache(0 To NumberOfNode - 1, 0 To NumberOfMaterial)
  For orig = 0 To NumberOfNode - 1
    P3NodeOnMatCache(orig, 0) = True
  Next orig
  If NumberOfFreeNode >= 1 Then
    ReDim invFree(0 To NumberOfFreeNode - 1)
    For orig = 0 To NumberOfFreeNode - 1
      invFree(orig) = -1
    Next orig
    For orig = 0 To NumberOfNode - 1
      On Error Resume Next
      freeNode = FEMFreeNodeMap(orig)
      If Err.Number = 0 Then
        If freeNode >= 0 And freeNode <= UBound(invFree) Then invFree(freeNode) = orig
      End If
      Err.Clear
      On Error GoTo 0
    Next orig
    For elementId = 0 To NumberOfElement - 1
      If P3IsElementActive(elementId) Then
        materialId = Elem(elementId).MatNo + 1
        If materialId >= 1 And materialId <= NumberOfMaterial Then
          For j = 0 To 7
            origFree = P6GetOriginalFreeNode(Elem(elementId).node(j))
            origNode = -1
            If origFree >= 0 And origFree <= UBound(invFree) Then origNode = invFree(origFree)
            If origNode >= 0 And origNode <= NumberOfNode - 1 Then
              P3NodeOnMatCache(origNode, materialId) = True
            End If
          Next j
        End If
      End If
    Next elementId
  End If
  P3LoadIndexNodeN = NumberOfNode - 1
  P3LoadIndexMatN = NumberOfMaterial
  P3LoadIndexActiveGen = P3ActiveSetGen
  P3LoadIndexReady = True
End Sub

Private Function P3NodeOnMaterialUncached(ByVal originalNode As Long, ByVal materialId As Long) As Boolean
  Dim elementId As Long, j As Long, freeNode As Long
  P3NodeOnMaterialUncached = False
  If originalNode < LBound(FEMFreeNodeMap) Or originalNode > UBound(FEMFreeNodeMap) Then Exit Function
  freeNode = FEMFreeNodeMap(originalNode)
  If freeNode < 0 Then Exit Function
  freeNode = P6GetInternalFreeNode(freeNode)
  If freeNode < 0 Then Exit Function
  If materialId <= 0 Then
    P3NodeOnMaterialUncached = True
    Exit Function
  End If
  If materialId > NumberOfMaterial Then Exit Function
  For elementId = 0 To NumberOfElement - 1
    If P3IsElementActive(elementId) Then
      If Elem(elementId).MatNo = materialId - 1 Then
        For j = 0 To 7
          If Elem(elementId).node(j) = freeNode Then
            P3NodeOnMaterialUncached = True
            Exit Function
          End If
        Next j
      End If
    End If
  Next elementId
End Function

Private Sub P3OverlayLoadingSheet()
  Dim ws As Worksheet, lastRow As Long, rowNo As Long, orig As Long, freeNode As Long, dofId As Long
  Dim kindText As String, dirText As String, selectorText As String, groupText As String
  Dim xmin As Double, xmax As Double, yMin As Double, yMax As Double, loadValue As Double
  Dim materialId As Long, hasBound As Boolean, i As Long
  Dim t0 As Double
  Dim errNum As Long, errSrc As String, errDesc As String
  t0 = Timer
  On Error Resume Next
  Set ws = ThisWorkbook.Worksheets("載荷")
  Err.Clear
  On Error GoTo LoadFail
  If lastDof >= 0 Then
    P3RestoreSupportBoundary
    ReDim P3AppliedForce(lastDof)
  End If
  If ws Is Nothing Then GoTo LoadDone
  If Not P3LoadingSheetHasAction() Then GoTo LoadDone
  If Not P3LoadIndexReady Or P3LoadIndexActiveGen <> P3ActiveSetGen Then P3BuildLoadIndex
  lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
  For rowNo = 2 To lastRow
    groupText = Trim$(CStr(ws.Cells(rowNo, 1).value2))
    If left$(groupText, 1) = "(" Then GoTo NextLoadRow
    selectorText = CStr(ws.Cells(rowNo, 2).value2)
    dirText = UCase$(Trim$(CStr(ws.Cells(rowNo, 3).value2)))
    kindText = UCase$(Trim$(CStr(ws.Cells(rowNo, 4).value2)))
    loadValue = 0#
    If IsNumeric(ws.Cells(rowNo, 5).value2) Then loadValue = CDbl(ws.Cells(rowNo, 5).value2)
    If kindText <> "DISP" And kindText <> "FORCE" And kindText <> "FIX" Then GoTo NextLoadRow
    If dirText <> "X" And dirText <> "Y" Then
      Err.Raise vbObjectError + 3232, "FEM.P3OverlayLoadingSheet", "載荷の方向はXまたはYです。行=" & CStr(rowNo) & " 値=" & dirText
    End If
    materialId = 0
    If IsNumeric(groupText) Then materialId = CLng(Val(groupText))
    If Not P3MaterialNetLoaded(materialId) Then GoTo NextLoadRow
    hasBound = False
    For orig = 0 To NumberOfNode - 1
      If P3NodeOnMaterial(orig, materialId) Then
        If Not hasBound Then
          xmin = x(orig): xmax = x(orig): yMin = y(orig): yMax = y(orig)
          hasBound = True
        Else
          If x(orig) < xmin Then xmin = x(orig)
          If x(orig) > xmax Then xmax = x(orig)
          If y(orig) < yMin Then yMin = y(orig)
          If y(orig) > yMax Then yMax = y(orig)
        End If
      End If
    Next orig
    If Not hasBound Then GoTo NextLoadRow
    For orig = 0 To NumberOfNode - 1
      freeNode = FEMFreeNodeMap(orig)
      If freeNode < 0 Then GoTo NextOrig
    If Not P3NodeOnMaterial(orig, materialId) Then GoTo NextOrig
    If UCase$(Trim$(selectorText)) = "BOX" Then
      If Not P3NodeInBox(orig, CDbl(Val(CStr(ws.Cells(rowNo, 6).value2))), CDbl(Val(CStr(ws.Cells(rowNo, 7).value2))), CDbl(Val(CStr(ws.Cells(rowNo, 8).value2))), CDbl(Val(CStr(ws.Cells(rowNo, 9).value2)))) Then GoTo NextOrig
    Else
      If Not P3NodeMatchesSelector(orig, selectorText, xmin, xmax, yMin, yMax) Then GoTo NextOrig
    End If
      dofId = P6GetInternalFreeNode(freeNode) * 2
      If dirText = "Y" Then dofId = dofId + 1
      If dofId < 0 Or dofId > lastDof Then GoTo NextOrig
      If kindText = "FORCE" Then
        P3AppliedForce(dofId) = P3AppliedForce(dofId) + loadValue
      Else
        NodeCond(dofId) = 1
        If kindText = "FIX" Then
          P3BoundaryDisp(dofId) = 0#
        Else
          P3BoundaryDisp(dofId) = loadValue
        End If
      End If
NextOrig:
    Next orig
NextLoadRow:
  Next rowNo
  GoTo LoadDone
LoadFail:
  errNum = Err.Number
  errSrc = Err.source
  errDesc = Err.Description
  Resume LoadDone
LoadDone:
  On Error GoTo 0
  P3BumpConstraintGeneration
  P6ProfLoadMs = P6ProfLoadMs + P6ElapsedMs(t0)
  If errNum <> 0 Then Err.Raise errNum, errSrc, errDesc
End Sub

Private Function P3PlanWantsKind(ByVal kindName As String) As Boolean
  If kindName = "GRAVITY" Then
    If P3StagePlanText = vbNullString Then
      P3PlanWantsKind = True
    Else
      P3PlanWantsKind = P3PlanHasGravity
    End If
  ElseIf kindName = "APPLY" Then
    If P3StagePlanText = vbNullString Then
      P3PlanWantsKind = True
    Else
      P3PlanWantsKind = P3PlanHasApply
    End If
  Else
    P3PlanWantsKind = False
  End If
End Function

Private Sub P3CaptureSrmSuccessSnapshot(ByVal strengthFactor As Double)
  Dim i As Long, k As Long, im As Long, gp As Long
  If NumberOfElement < 1 Or lastDof < 0 Then
    P3SrmSnapReady = False
    Exit Sub
  End If
  P3EnsureSrmSnapWorkspace
  For k = 0 To NumberOfElement - 1
    For im = 0 To 3
      For i = 0 To 16
        P3SrmSnapStmat(i, im, k) = Elem(k).Stmat(i, im)
        P3SrmSnapMStmat(i, im, k) = Elem(k).mStmat(i, im)
      Next i
    Next im
    For gp = 0 To 2
      P3SrmSnapJointContact(gp, k) = Elem(k).JointContact(gp)
    Next gp
  Next k
  For i = 0 To lastDof
    P3SrmSnapTDisp(i) = TDisp(i)
    P3SrmSnapUDisp(i) = UDisp(i)
    P3SrmSnapCommittedDisp(i) = P3CommittedDisp(i)
    P3SrmSnapInternal(i) = P3FinalInternalForce(i)
  Next i
  PlasticLogCaptureSuccess
  P3SrmSnapCommittedPlasticStrain = P3CommittedPlasticStrain
  P3SrmSnapCommittedPlasticMultiplier = P3CommittedPlasticMultiplier
  P3SrmSnapCommittedYieldFunction = P3CommittedYieldFunction
  P3SrmSnapCommittedYielded = P3CommittedYielded
  P3SrmSnapCommittedPrincipalStress = P3CommittedPrincipalStress
  P3SrmSnapCommittedSpmat = P3CommittedSpmat
  P3SrmSnapCommittedDpmat = P3CommittedDpmat
  P3SrmSnapCommittedSpmatValid = P3CommittedSpmatValid
  P3SrmSnapActivePlastic = P3ActivePlasticPoint
  P3SrmSnapActiveCount = P3ActivePlasticPointCount
  P3SrmSnapSuccessInc = P3SuccessfulIncrementCount
  P3SrmSnapLastInc = P3LastConvergedIncrement
  P3SrmSnapLastLoad = P3LastConvergedLoadFactor
  P3SrmSnapMaxDisp = 0#
  For i = 0 To lastDof
    If Abs(TDisp(i)) > P3SrmSnapMaxDisp Then P3SrmSnapMaxDisp = Abs(TDisp(i))
  Next i
  P3SrmSnapMaxCorr = P3MaxCorrection
  P3SrmSnapRelRes = RelativeResidualFree
  P3SrmSnapResFull = ResidualNormFull
  P3SrmSnapResFree = ResidualNormFree
  P3SrmSnapMaxAbsRes = MaxAbsResidualFree
  P3SrmSnapEnergy = EnergyError
  P3SrmSnapRx = P6ReactionSumX
  P3SrmSnapRy = P6ReactionSumY
  P3SrmSnapInc = CurrentIncrement
  P3SrmSnapIter = CurrentIteration
  P3SrmSnapFs = strengthFactor
  P3SrmSnapPrefix = P3ActivePrefix
  If NumberOfElement >= 1 Then
    For k = 0 To NumberOfElement - 1
      P3SrmSnapActive(k) = P3IsElementActive(k)
    Next k
  End If
  For i = 0 To lastDof
    P3SrmSnapNodeCond(i) = NodeCond(i)
    P3SrmSnapBoundaryDisp(i) = P3BoundaryDisp(i)
    If UBound(P3AppliedForce) >= lastDof Then P3SrmSnapApplied(i) = P3AppliedForce(i)
    If UBound(P3SelfWeightForce) >= lastDof Then P3SrmSnapSelfWeight(i) = P3SelfWeightForce(i)
    On Error Resume Next
    P3SrmSnapLockedSelf(i) = P3LockedSelfWeight(i)
    P3SrmSnapLockedApplied(i) = P3LockedApplied(i)
    On Error GoTo 0
  Next i
  P3SrmSnapReady = True
End Sub

Private Function P3RestoreSrmSuccessSnapshot() As Boolean
  Dim i As Long, k As Long, im As Long, gp As Long
  P3RestoreSrmSuccessSnapshot = False
  If Not P3SrmSnapReady Then Exit Function
  If NumberOfElement < 1 Or lastDof < 0 Then Exit Function
  If Not P3SetMaterialForStrengthFactor(P3SrmSnapFs) Then Exit Function
  For i = 0 To lastDof
    TDisp(i) = P3SrmSnapTDisp(i)
    UDisp(i) = P3SrmSnapUDisp(i)
    P3CommittedDisp(i) = P3SrmSnapCommittedDisp(i)
    P3FinalInternalForce(i) = P3SrmSnapInternal(i)
    Disp(i) = 0#
  Next i
  PlasticLogRestoreSuccess
  P3CommittedPlasticStrain = P3SrmSnapCommittedPlasticStrain
  P3CommittedPlasticMultiplier = P3SrmSnapCommittedPlasticMultiplier
  P3CommittedYieldFunction = P3SrmSnapCommittedYieldFunction
  P3CommittedYielded = P3SrmSnapCommittedYielded
  P3CommittedPrincipalStress = P3SrmSnapCommittedPrincipalStress
  P3CommittedSpmat = P3SrmSnapCommittedSpmat
  P3CommittedDpmat = P3SrmSnapCommittedDpmat
  P3CommittedSpmatValid = P3SrmSnapCommittedSpmatValid
  P3ActivePlasticPoint = P3SrmSnapActivePlastic
  P3ActivePlasticPointCount = P3SrmSnapActiveCount
  P3SuccessfulIncrementCount = P3SrmSnapSuccessInc
  P3LastConvergedIncrement = P3SrmSnapLastInc
  P3LastConvergedLoadFactor = P3SrmSnapLastLoad
  P3CurrentStrengthFactor = P3SrmSnapFs
  FSS = P3SrmSnapFs
  P3SrmLower = P3SrmSnapFs
  P3MaxTrialDisp = P3SrmSnapMaxDisp
  P3MaxCorrection = P3SrmSnapMaxCorr
  RelativeResidualFree = P3SrmSnapRelRes
  ResidualNormFull = P3SrmSnapResFull
  ResidualNormFree = P3SrmSnapResFree
  MaxAbsResidualFree = P3SrmSnapMaxAbsRes
  EnergyError = P3SrmSnapEnergy
  P3LastRelativeResidual = P3SrmSnapRelRes
  P6ReactionSumX = P3SrmSnapRx
  P6ReactionSumY = P3SrmSnapRy
  CurrentIncrement = P3SrmSnapInc
  CurrentIteration = P3SrmSnapIter
  P3ActivePrefix = P3SrmSnapPrefix
  If NumberOfElement >= 1 Then
    If Not P3ElementActiveReady Then P3InitElementActive
    For k = 0 To NumberOfElement - 1
      If k <= UBound(P3SrmSnapActive) Then P3ElementActive(k) = P3SrmSnapActive(k)
    Next k
  End If
  ReDim P3AppliedForce(lastDof)
  ReDim P3SelfWeightForce(lastDof)
  ReDim P3LockedSelfWeight(lastDof)
  ReDim P3LockedApplied(lastDof)
  ReDim P3BoundaryDisp(lastDof)
  For i = 0 To lastDof
    NodeCond(i) = P3SrmSnapNodeCond(i)
    P3BoundaryDisp(i) = P3SrmSnapBoundaryDisp(i)
    P3AppliedForce(i) = P3SrmSnapApplied(i)
    P3SelfWeightForce(i) = P3SrmSnapSelfWeight(i)
    P3LockedSelfWeight(i) = P3SrmSnapLockedSelf(i)
    P3LockedApplied(i) = P3SrmSnapLockedApplied(i)
  Next i
  P3BumpConstraintGeneration
  P6InvalidateActiveDependentCaches
  For k = 0 To NumberOfElement - 1
    For im = 0 To 3
      For i = 0 To 16
        Elem(k).Stmat(i, im) = P3SrmSnapStmat(i, im, k)
        Elem(k).mStmat(i, im) = P3SrmSnapMStmat(i, im, k)
      Next i
    Next im
    For gp = 0 To 2
      Elem(k).JointContact(gp) = P3SrmSnapJointContact(gp, k)
    Next gp
  Next k
  P3RestoreCommittedMaterialState False
  P3MarkAllTangentDirty
  P3TrialStateValid = False
  If Not P3RebuildTangentFromSpmat() Then Exit Function
  FEMAppendRunLog "SRM復元", "PASS", "保存したFs=" & Format$(P3SrmSnapFs, "0.000") & " 変位max=" & Format$(P3SrmSnapMaxDisp, "0.000E+00") & " 残差=" & Format$(P3SrmSnapRelRes, "0.000E+00")
  P3RestoreSrmSuccessSnapshot = True
End Function

Public Sub P3PrepareOutputResidual()
  Dim i As Long
  Dim appliedOk As Boolean
  If lastDof < 0 Then Exit Sub
  P3UseNonlinearResidual = True
  On Error Resume Next
  appliedOk = (UBound(P3AppliedForce) >= lastDof)
  If Err.Number <> 0 Then appliedOk = False
  Err.Clear
  If UBound(P3SelfWeightForce) < lastDof Then
    On Error GoTo 0
    Exit Sub
  End If
  If Err.Number <> 0 Then
    On Error GoTo 0
    Exit Sub
  End If
  On Error GoTo 0
  ReDim OriginalForce(lastDof)
  For i = 0 To lastDof
    OriginalForce(i) = P3SelfWeightForce(i)
    If appliedOk Then OriginalForce(i) = OriginalForce(i) + P3AppliedForce(i)
  Next i
  ComputeResidual
End Sub

Private Function P3CanReuseGravityForFs1() As Boolean
  Dim i As Long, kindText As String
  P3CanReuseGravityForFs1 = False
  If P3GravityReuseStatus <> "PASS" And P3GravityReuseStatus <> "DEFER" Then Exit Function
  If Abs(P3CurrentStrengthFactor - 1#) > 0.000000000001 Then Exit Function
  If P3SrmReplayLimit < 1 Then Exit Function
  For i = 1 To P3SrmReplayLimit
    If P3StageOn(i) Then
      kindText = P3StageKind(i)
      If kindText <> "GRAVITY" Then Exit Function
    End If
  Next i
  P3CanReuseGravityForFs1 = True
End Function

Private Function P3TryStrengthFactor(ByVal strengthFactor As Double) As Boolean
  P3TryStrengthFactor = False
  If V3CostSearchEnabled Then V3BeginTrialCost
  P3ClearTransientFailure
  P3SrmTrialFailNote = vbNullString
  If Not P3SetMaterialForStrengthFactor(strengthFactor) Then
    FEMAppendRunLog "SRM試行", "INVALID_INPUT", "Fs=" & Format$(strengthFactor, "0.000000000000") & ";" & AnalysisMessage
    Exit Function
  End If
  P3SrmTrialCount = P3SrmTrialCount + 1
  P6PerfMark
  FEMProgressFs P3SrmTrialCount, strengthFactor
  If SuppressUserMessages Then SaveP0Progress "srm_try_" & Format$(strengthFactor, "0.000")
  P3SrmTrialRunning = True
  If P3SrmLogStageNo > 0 Then
    P3RunLogStageNo = P3SrmLogStageNo
    P3RunLogKind = P3SrmLogKind
  End If
  FEMAppendRunLog "SRM試行", "TRY", "Fs=" & Format$(strengthFactor, "0.000")
  If P3ReplayStages(P3SrmReplayLimit) Then
    P3TryStrengthFactor = True
    P3SrmLower = strengthFactor
    P3CaptureSrmSuccessSnapshot strengthFactor
    If P3SrmLogStageNo > 0 Then
      P3RunLogStageNo = P3SrmLogStageNo
      P3RunLogKind = P3SrmLogKind
    End If
    FEMAppendRunLog "SRM試行", "PASS", "Fs=" & Format$(strengthFactor, "0.000") & " 残差=" & Format$(RelativeResidualFree, "0.000E+00") & " 変位max=" & Format$(P3SrmSnapMaxDisp, "0.000E+00")
    P6PerfEmit "SRM試行", "PASS"
    If V3CostSearchEnabled Then V3RecordTrialCost True
  Else
    If Len(P3SrmTrialFailNote) = 0 Then P3SrmTrialFailNote = AnalysisMessage
    If P3SrmLogStageNo > 0 Then
      P3RunLogStageNo = P3SrmLogStageNo
      P3RunLogKind = P3SrmLogKind
    End If
    FEMAppendRunLog "SRM試行", "FAIL", "Fs=" & Format$(strengthFactor, "0.000") & " " & P3SrmTrialFailNote
    P6PerfEmit "SRM試行", "FAIL"
    If V3CostSearchEnabled And Not P3IsFatalFailure() Then V3RecordTrialCost False
  End If
  P3SrmTrialRunning = False
End Function

Private Function P3SrmAbortIfFatal() As Boolean
  P3SrmAbortIfFatal = False
  If Not P3IsFatalFailure() Then Exit Function
  P3SrmNote = "SRM停止: " & AnalysisMessage
  P3SrmAbortIfFatal = True
End Function

Private Function P3RunStrengthReduction() As Boolean
  Dim fsValue As Double, fMax As Double, fTol As Double
  Dim fLow As Double, fHigh As Double, stepValue As Double
  Dim bisectCount As Long, baselineWidth As Double
  P3RunStrengthReduction = False
  P3SrmTrialCount = 0
  P3SrmLower = 0#
  P3SrmUpper = 0#
  P3SrmSnapReady = False
  V3CostSearchEnabled = (P6ReadSetting("ACCEL_V3_COST_SEARCH", 0#) = 1#)
  V3BracketStep = 0
  V3LastPassCostSec = 0#: V3LastFailCostSec = 0#
  V3PassCostSamples = 0: V3FailCostSamples = 0
  fMax = P6ReadSetting("SRM_FMAX", 3#)
  fTol = P6ReadSetting("SRM_TOL", 0.025)
  If fMax <= 1# Then fMax = 3#
  If fTol < 0.001 Then fTol = 0.001
  fsValue = P6ReadSetting("SRM_FIXED_FS", 0#)
  If fsValue > 0# Then
    FEMAppendRunLog "SRM", "FIXED", "Fs=" & Format$(fsValue, "0.000") & " 探索は行わない"
    If P3TryStrengthFactor(fsValue) Then
      FSS = fsValue
      P3PublishFos fsValue, 0#, False
      P3RunStrengthReduction = True
    End If
    Exit Function
  End If
  fsValue = 1#
  If P3CanReuseGravityForFs1() Then
    P6SrmFs1ReuseCount = P6SrmFs1ReuseCount + 1
    P3SrmTrialCount = P3SrmTrialCount + 1
    P6PerfMark
    FEMProgressFs P3SrmTrialCount, 1#
    If P3SrmLogStageNo > 0 Then
      P3RunLogStageNo = P3SrmLogStageNo
      P3RunLogKind = P3SrmLogKind
    End If
    FEMAppendRunLog "SRM試行", "TRY", "Fs=1.000 REUSE_FROM_GRAVITY"
    If P3GravityReuseStatus = "PASS" Then
      P3CaptureSrmSuccessSnapshot 1#
      FEMAppendRunLog "SRM試行", "PASS", "Fs=1.000 PASS_REUSED_FROM_GRAVITY 残差=" & Format$(RelativeResidualFree, "0.000E+00")
      P6PerfEmitReuse "PASS_REUSED"
      fLow = 1#
      stepValue = 0.1
      Do
        fsValue = fLow + stepValue
        If fsValue > fMax Then fsValue = fMax
        If P3TryStrengthFactor(fsValue) Then
          fLow = fsValue
          If Abs(fLow - fMax) <= 0.000000001 Then
            P3SrmLower = fLow
            P3SrmUpper = fMax
            FSS = fLow
            P3PublishFos fLow, fMax, False
            P3RunStrengthReduction = True
            Exit Function
          End If
        Else
          If P3SrmAbortIfFatal() Then Exit Function
          fHigh = fsValue
          Exit Do
        End If
      Loop
    Else
      FEMAppendRunLog "SRM試行", "FAIL", "Fs=1.000 FAIL_REUSED_FROM_GRAVITY"
      P6PerfEmitReuse "FAIL_REUSED"
      fHigh = 1#
      fsValue = 0.9
      Do While fsValue >= 0.2 - 0.000000001
        If P3TryStrengthFactor(fsValue) Then
          fLow = fsValue
          Exit Do
        End If
        If P3SrmAbortIfFatal() Then Exit Function
        fHigh = fsValue
        fsValue = fsValue - 0.1
      Loop
      If fLow <= 0# Then
        SetAnalysisFailure RESULT_NONCONVERGED, "SRM: Fs=0.2でも自重／載荷が収束しません。メッシュ・拘束・強度を確認してください。", vbObjectError + 3220, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
        P3SrmNote = "SRM失敗: 下限でも非収束"
        Exit Function
      End If
    End If
  ElseIf P3TryStrengthFactor(fsValue) Then
    fLow = fsValue
    stepValue = 0.1
    Do
      fsValue = fLow + stepValue
      If fsValue > fMax Then fsValue = fMax
      If P3TryStrengthFactor(fsValue) Then
        fLow = fsValue
        If Abs(fLow - fMax) <= 0.000000001 Then
          P3SrmLower = fLow
          P3SrmUpper = fMax
          FSS = fLow
          P3PublishFos fLow, fMax, False
          P3SrmNote = P3SrmNote & " FMAXまで収束"
          P3RunStrengthReduction = True
          Exit Function
        End If
      Else
        If P3SrmAbortIfFatal() Then Exit Function
        fHigh = fsValue
        Exit Do
      End If
    Loop
  Else
    If P3SrmAbortIfFatal() Then Exit Function
    fHigh = 1#
    fsValue = 0.9
    Do While fsValue >= 0.2 - 0.000000001
      If P3TryStrengthFactor(fsValue) Then
        fLow = fsValue
        Exit Do
      End If
      If P3SrmAbortIfFatal() Then Exit Function
      fHigh = fsValue
      fsValue = fsValue - 0.1
    Loop
    If fLow <= 0# Then
      SetAnalysisFailure RESULT_NONCONVERGED, "SRM: Fs=0.2でも自重／載荷が収束しません。メッシュ・拘束・強度を確認してください。", vbObjectError + 3220, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
      P3SrmNote = "SRM失敗: 下限でも非収束"
      Exit Function
    End If
  End If
  P3SrmUpper = fHigh
  If V3CostSearchEnabled Then
    baselineWidth = fHigh - fLow
    Do While baselineWidth > fTol: baselineWidth = baselineWidth * 0.5: Loop
    fTol = baselineWidth
  End If
  AccelSrmWidthTarget = fTol
  bisectCount = 0
  Do While (fHigh - fLow) > fTol And bisectCount < 40
    fsValue = NextFsByBracket(fLow, fHigh, fTol)
    If (fHigh - fLow) <= fTol Then Exit Do
    If P3TryStrengthFactor(fsValue) Then
      fLow = fsValue
    Else
      If P3SrmAbortIfFatal() Then Exit Function
      fHigh = fsValue
    End If
    bisectCount = bisectCount + 1
  Loop
  If P3SrmSnapReady And Abs(P3SrmSnapFs - fLow) <= 0.000000001 Then
    If Not P3RestoreSrmSuccessSnapshot() Then
      SetAnalysisFailure RESULT_NONCONVERGED, "SRM: 収束Fs=" & Format$(fLow, "0.000") & " の保存状態を復元できませんでした。", vbObjectError + 3221, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
      P3SrmNote = "SRMスナップショット復元失敗"
      Exit Function
    End If
  ElseIf Not P3TryStrengthFactor(fLow) Then
    SetAnalysisFailure RESULT_NONCONVERGED, "SRM: 最後に収束したFs=" & Format$(fLow, "0.000") & " の再解析に失敗しました。", vbObjectError + 3221, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
    P3SrmNote = "SRM再解析失敗"
    Exit Function
  End If
  FSS = fLow
  P3SrmLower = fLow
  P3SrmUpper = fHigh
  P3PublishFos fLow, fHigh, True
  P3ClearTransientFailure
  P3RunStrengthReduction = True
End Function

Private Function P3ReplayStages(ByVal lastIndex As Long) As Boolean
  Dim i As Long, lastI As Long, startI As Long, ranAny As Boolean
  Dim previousQuiet As Boolean
  P3ReplayStages = False
  lastI = lastIndex
  If lastI > P3StageN Then lastI = P3StageN
  If lastI < 0 Then lastI = 0
  startI = P3SrmReplayStart
  If startI < 0 Then startI = 0
  previousQuiet = P3ReplayQuiet
  P3ReplayQuiet = True
  If startI <= 0 Then
    P3InitializeTrialState
    If lastI <= 0 Then
      P3FlagsFromPrefix 0
      P3ReplayStages = P3RunLoadStages()
      P3ReplayQuiet = previousQuiet
      Exit Function
    End If
    For i = 1 To lastI
      If P3StageOn(i) Then
        If P3StageKind(i) <> "SRM" Then
          If Not P3ExecuteStage(i) Then
            P3ReplayQuiet = previousQuiet
            Exit Function
          End If
        End If
      End If
    Next i
  Else
    If Not P3OriginSnapReady Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "SRM replay originがありません。RESET/MATSET後のチェックポイントに origin.csv が必要です。", vbObjectError + 3101, -1, -1, 0, 0
      P3ReplayQuiet = previousQuiet
      Exit Function
    End If
    P3ZeroTrialKinematics
    P3RestoreReplayOrigin
    ranAny = False
    For i = startI + 1 To lastI
      If P3StageOn(i) Then
        If P3StageKind(i) <> "SRM" And P3StageKind(i) <> "RESET_U" And P3StageKind(i) <> "RESET_STRESS" And P3StageKind(i) <> "MATSET" Then
          ranAny = True
          If Not P3ExecuteStage(i) Then
            P3ReplayQuiet = previousQuiet
            Exit Function
          End If
        End If
      End If
    Next i
    If Not ranAny Then
      P3FlagsFromPrefix lastI
      If Not P3RunLoadStages() Then
        P3ReplayQuiet = previousQuiet
        Exit Function
      End If
    End If
  End If
  P3ReplayQuiet = previousQuiet
  P3ReplayStages = True
End Function

Private Function P3ExecuteStage(ByVal stageIndex As Long) As Boolean
  Dim kindText As String, materialId As Long, logNo As Long, requiredFs As Double
  P3ExecuteStage = False
  p3StageResultLogged = False
  If stageIndex < 1 Or stageIndex > P3StageN Then Exit Function
  If Not P3StageOn(stageIndex) Then
    P3ExecuteStage = True
    Exit Function
  End If
  P3ActivePrefix = stageIndex
  kindText = P3StageKind(stageIndex)
  FEMLastProc = "stage " & CStr(stageIndex) & " " & kindText
  materialId = P3StageMaterialId(stageIndex)
  logNo = P3StageId(stageIndex)
  P3RunLogStageNo = logNo
  P3RunLogKind = kindText
  If kindText = "SRM" Then
    P3SrmLogStageNo = logNo
    P3SrmLogKind = "SRM"
  End If
  If Not P3ReplayQuiet Then FEMProgressStage logNo, kindText
  If Not P3ReplayQuiet Then FEMAppendRunLog "ステージ開始", "RUN", kindText
  If kindText = "GRAVITY" Then
    P3PlanHasGravity = True
    P3PlanHasApply = False
    If Not P3ReplayQuiet Then P6PerfMark
    If Not P3RunLoadStages() Then
      If P3CanDeferGravityToSrm(stageIndex) Then
        If Not P3ReplayQuiet Then
          P3LogStageResult logNo, kindText, P3StageGroup(stageIndex), False, 1#, "Fs=1で自重が非収束。後続SRMでFs<1を探索。 " & AnalysisMessage
          FEMAppendRunLog "ステージ", "DEFER", "GRAVITY失敗→SRMへ " & AnalysisMessage
          P6PerfEmit "GRAVITY", "DEFER"
          P3GravityReuseStatus = "DEFER"
        End If
        P3ClearTransientFailure
        P3ExecuteStage = True
        Exit Function
      End If
      If Not P3ReplayQuiet Then P6PerfEmit "GRAVITY", "FAIL"
      Exit Function
    End If
    If Not P3ReplayQuiet Then
      P3LogStageResult logNo, kindText, P3StageGroup(stageIndex), True, 1#, "自重"
      P3WriteStageSnapshot
      P6PerfEmit "GRAVITY", "PASS"
      P3GravityReuseStatus = "PASS"
    End If
  ElseIf kindText = "APPLY" Or kindText = "LOAD" Then
    P3OverlayLoadingSheet
    P3HoldOrphanDofs
    P3PlanHasGravity = False
    P3PlanHasApply = True
    If Not P3RunLoadStages() Then Exit Function
    P3LogStageResult logNo, kindText, P3StageGroup(stageIndex), True, 1#, "載荷"
    P3WriteStageSnapshot
  ElseIf kindText = "UNLOAD" Then
    P3OverlayLoadingSheet
    P3HoldOrphanDofs
    P3PlanHasGravity = False
    P3PlanHasApply = True
    If Not P3RunLoadStages() Then Exit Function
    P3LogStageResult logNo, kindText, P3StageGroup(stageIndex), True, 1#, "除荷"
    P3WriteStageSnapshot
  ElseIf kindText = "BIRTH" Then
    If materialId <= 0 Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "BIRTHの材料番号がありません。ステージ=" & CStr(logNo), vbObjectError + 3230, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    P3SetMaterialElementsActive materialId, True
    P3HoldOrphanDofs
    P3AssembleSelfWeightOnly
    P3RefreshActiveStiffness
    P3PlanHasGravity = True
    P3PlanHasApply = False
    If Not P3RunLoadStages() Then Exit Function
    P3LogStageResult logNo, kindText, P3StageGroup(stageIndex), True, 1#, "BIRTH 材料=" & CStr(materialId)
    P3WriteStageSnapshot
  ElseIf kindText = "DEATH" Then
    If materialId <= 0 Then
      SetAnalysisFailure RESULT_INPUT_ERROR, "DEATHの材料番号がありません。ステージ=" & CStr(logNo), vbObjectError + 3231, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    P3SetMaterialElementsActive materialId, False
    P3HoldOrphanDofs
    P3AssembleSelfWeightOnly
    P3RefreshActiveStiffness
    P3PlanHasGravity = True
    P3PlanHasApply = False
    If Not P3RunLoadStages() Then Exit Function
    P3LogStageResult logNo, kindText, P3StageGroup(stageIndex), True, 1#, "DEATH 材料=" & CStr(materialId)
    P3WriteStageSnapshot
  ElseIf kindText = "RESET_U" Then
    P3ResetKinematicState False
    P3SrmReplayStart = stageIndex
    P3CaptureReplayOrigin
    If Not P3ReplayQuiet Then
      P3LogStageResult logNo, kindText, P3StageGroup(stageIndex), True, 1#, "変位・応力をゼロ。座標はそのまま"
      P3WriteStageSnapshot
    End If
  ElseIf kindText = "RESET_STRESS" Then
    P3ResetKinematicState True
    P3SrmReplayStart = stageIndex
    P3CaptureReplayOrigin
    If Not P3ReplayQuiet Then
      P3LogStageResult logNo, kindText, P3StageGroup(stageIndex), True, 1#, "座標を変形後にし応力をゼロ"
      P3WriteStageSnapshot
    End If
  ElseIf kindText = "MATSET" Then
    If Not P3ApplyMatSet(stageIndex) Then Exit Function
    P3SrmReplayStart = stageIndex
    P3CaptureReplayOrigin
    If Not P3ReplayQuiet Then
      If Len(P3StateNote) > 0 Then
        P3LogStageResult logNo, kindText, P3StageGroup(stageIndex), True, 1#, "材料を書き換え Fs基準を更新 / " & P3StateNote
      Else
        P3LogStageResult logNo, kindText, P3StageGroup(stageIndex), True, 1#, "材料を書き換え Fs基準を更新"
      End If
      P3WriteStageSnapshot
    End If
  ElseIf kindText = "SRM" Then
    If P6BandBenchOnly Then
      P6FlushBandBench
      FEMAppendRunLog "BAND_LU_BENCH", "ONLY", "重力後にベンチ済み。SRMは実行しない"
      P3LogStageResult logNo, "SRM", P3StageGroup(stageIndex), True, 1#, "BAND_LU_BENCH=1 のためSRM未実行", False
      P3ExecuteStage = True
      Exit Function
    End If
    P3SrmReplayStart = P3LastOriginIndex(stageIndex - 1)
    P3FlagsFromPrefix stageIndex - 1
    P3SrmReplayLimit = stageIndex - 1
    If P3SrmReplayLimit < 0 Then P3SrmReplayLimit = 0
    P3ActivePrefix = P3SrmReplayLimit
    If Not P3RunStrengthReduction() Then Exit Function
    If P3HasStageAfter(stageIndex) Then
      requiredFs = 1#
      If IsNumeric(P3StageParam(stageIndex)) Then requiredFs = CDbl(P3StageParam(stageIndex))
      If requiredFs <= 0# Then requiredFs = 1#
      If FSS + 0.000000001 < requiredFs Then
        SetAnalysisFailure RESULT_NONCONVERGED, "途中SRM: Fs=" & Format$(FSS, "0.000") & " < 必要 " & Format$(requiredFs, "0.000") & " のため次ステージへ進めません。", vbObjectError + 3222, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
        P3SrmNote = "ゲート未達で中断 Fs=" & Format$(FSS, "0.000")
        P3LogStageResult logNo, "SRM", P3StageGroup(stageIndex), False, FSS, P3SrmNote, True
        Exit Function
      End If
      P3LogStageResult logNo, "SRM", P3StageGroup(stageIndex), True, P3FosDisplayFs(), P3SrmNote, False
      P3WriteStageSnapshot
      FEMProgressPhase "ステージ " & CStr(logNo) & "  SRM通過", "FOS = " & Format$(P3FosDisplayFs(), "0.000") & "  強度を戻して履歴を再生中"
      If Not P3SetMaterialForStrengthFactor(1#) Then
        P3LogStageResult logNo, "SRM", P3StageGroup(stageIndex), False, FSS, "ゲート通過後のFs=1復帰に失敗", True
        Exit Function
      End If
      If Not P3ReplayStages(stageIndex - 1) Then
        P3LogStageResult logNo, "SRM", P3StageGroup(stageIndex), False, FSS, "ゲート通過後の履歴再生に失敗", True
        Exit Function
      End If
      P3LogStageResult logNo, "SRM", "", True, FSS, "ゲート通過。強度を戻し履歴を再生して続行", True
    Else
      P3LogStageResult logNo, "SRM", P3StageGroup(stageIndex), True, P3FosDisplayFs(), P3SrmNote, True
      P3WriteStageSnapshot
    End If
  Else
    SetAnalysisFailure RESULT_INPUT_ERROR, "未対応のステージ種別です。ステージ=" & CStr(logNo) & " 種別=" & kindText, vbObjectError + 3233, -1, -1, CurrentIncrement, CurrentIteration
    P3LogStageResult logNo, kindText, P3StageGroup(stageIndex), False, 1#, "未対応種別"
    Exit Function
  End If
  P3ExecuteStage = True
End Function

' Replay bookkeeping for stages already covered by the checkpoint. RESET/SRM are not re-executed; origin comes from origin.csv.
Private Sub P3ApplySkippedStage(ByVal stageIndex As Long)
  Dim kindText As String, materialId As Long
  If stageIndex < 1 Or stageIndex > P3StageN Then Exit Sub
  If Not P3StageOn(stageIndex) Then Exit Sub
  kindText = P3StageKind(stageIndex)
  materialId = P3StageMaterialId(stageIndex)
  P3ActivePrefix = stageIndex
  Select Case kindText
    Case "GRAVITY"
      P3PlanHasGravity = True
    Case "APPLY", "LOAD"
      P3OverlayLoadingSheet
      P3HoldOrphanDofs
      P3PlanHasApply = True
    Case "UNLOAD"
      P3OverlayLoadingSheet
      P3HoldOrphanDofs
      P3PlanHasApply = True
    Case "BIRTH"
      ' Checkpoint already has the born material's Gauss history; only restore the active flag.
      If materialId > 0 Then P3SetMaterialElementsActive materialId, True, False
      P3HoldOrphanDofs
    Case "DEATH"
      If materialId > 0 Then P3SetMaterialElementsActive materialId, False
      P3HoldOrphanDofs
    Case "MATSET"
      If Not P3ApplyMatSet(stageIndex) Then Exit Sub
    Case "RESET_U", "RESET_STRESS", "SRM"
      ' チェックポイントがリセット／SRM後の変位・Gaussを持っているので再実行しない
  End Select
End Sub

Private Function P3RunAnalysis() As Boolean
  Dim strengthFactor As Double, srmIndex As Long, requiredFs As Double
  Dim stageIndex As Long, kindText As String, materialId As Long
  Dim okFlag As Boolean, resumeIndex As Long
  On Error GoTo P3RunFailed
  FEMLastProc = "P3RunAnalysis"
  AccelResetRun
  P3RunAnalysis = False
  P1Formulation = P1_FORMULATION_PLANE_STRAIN
  P3SrmTrialCount = 0
  P3SrmLower = 0#
  P3SrmUpper = 0#
  P3SrmNote = "SRMなし"
  P3GravityReuseStatus = "NONE"
  P6SrmFs1ReuseCount = 0
  P3SrmSnapReady = False
  P3ReplayQuiet = False
  P3SrmTrialRunning = False
  P3LastCompletedStage = 0
  P3SupportCondReady = False
  P3ElementActiveReady = False
  P3ActivePrefix = 0
  P3SrmReplayStart = 0
  P3OriginSnapReady = False
  FEMEnsureStageSheets
  P3ResetStageResultSheet
  P3LoadStagePlan
  If ResultStatus = RESULT_INPUT_ERROR Then Exit Function
  P3CaptureOriginalMaterial
  strengthFactor = 1#
  FSS = 1#
  If Not P3SetMaterialForStrengthFactor(strengthFactor) Then Exit Function
  P3InitializeTrialState
  If P3StageN <= 0 Then
    P3FlagsFromPrefix 0
    If P3RunLoadStages() Then
      FSS = strengthFactor
      P3LogStageResult 0, P3StagePlanText, "", True, strengthFactor, "通常解析"
      P3RunAnalysis = True
    End If
    Exit Function
  End If
  femIoResumeStage = FEMIoTryResume()
  If ResultStatus = RESULT_INPUT_ERROR Then Exit Function
  resumeIndex = 0
  If femIoResumeStage > 0 Then
    resumeIndex = femIoResumeIndex
    If resumeIndex > 0 Then P3FlagsFromPrefix resumeIndex
  End If
  For stageIndex = 1 To P3StageN
    If resumeIndex > 0 And stageIndex <= resumeIndex Then
      P3ApplySkippedStage stageIndex
      If ResultStatus = RESULT_INPUT_ERROR Or ResultStatus = RESULT_MATERIAL_ERROR Then Exit Function
      P3LogStageResult P3StageId(stageIndex), P3StageKind(stageIndex), P3StageGroup(stageIndex), True, 1#, "再開スキップ（外部出力）", False
      GoTo P3NextPlanStage
    End If
    If resumeIndex > 0 And stageIndex = resumeIndex + 1 Then
      If Not FEMIoFinishResume() Then Exit Function
    End If
    If Not P3ExecuteStage(stageIndex) Then
      If Not p3StageResultLogged Then P3LogStageResult P3StageId(stageIndex), P3StageKind(stageIndex), P3StageGroup(stageIndex), False, P3CurrentStrengthFactor, AnalysisMessage
      Exit Function
    End If
P3NextPlanStage:
  Next stageIndex
  If resumeIndex > 0 And resumeIndex >= P3StageN Then
    If Not FEMIoFinishResume() Then Exit Function
  End If
  P3RunAnalysis = True
  Exit Function
P3RunFailed:
  FEMLastProc = FEMLastProc & " err=" & CStr(Err.Number)
  If Err.Number = 9 Then
    SetAnalysisFailure RESULT_RUNTIME_ERROR, "インデックスが有効範囲にありません。場所=" & FEMLastProc & " 要素=" & CStr(FailureElement) & " Ver=" & FEM_BUILD_STAMP, Err.Number, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
  Else
    SetAnalysisFailure RESULT_RUNTIME_ERROR, Err.Description & " 場所=" & FEMLastProc & " Ver=" & FEM_BUILD_STAMP, Err.Number, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
  End If
  P3RunAnalysis = False
End Function

Function PlaCalc() As Boolean
  PlaCalc = P3RunAnalysis()
  Exit Function
End Function
Public Sub P2ResetOutput(ByRef outputState As P2_MaterialPointOutput)
  Dim i As Long, j As Long
  For i = 0 To 3
    outputState.TrialStress(i) = 0#
    outputState.Stress(i) = 0#
    outputState.PlasticStrain(i) = 0#
  Next i
  For i = 0 To 2
    For j = 0 To 2
      outputState.Tangent(i, j) = 0#
    Next j
  Next i
  outputState.PlasticMultiplier = 0#
  outputState.TrialYieldFunction = 0#
  outputState.YieldFunction = 0#
  outputState.PrincipalStress(0) = 0#
  outputState.PrincipalStress(1) = 0#
  outputState.PrincipalStress(2) = 0#
  outputState.principalAngle = 0#
  outputState.yielded = False
  outputState.PlasticOccurred = False
  outputState.converged = False
  outputState.Elastic = False
  outputState.failureCode = 0
  outputState.failureMessage = vbNullString
  outputState.iterations = 0
  outputState.UsedSubsteps = 0
End Sub

Public Function P2IsFinite(ByVal value As Double) As Boolean
  P2IsFinite = False
  If value <> value Then Exit Function
  If Abs(value) > 1E+300 Then Exit Function
  P2IsFinite = True
End Function

Private Function P2Atan2(ByVal yValue As Double, ByVal xValue As Double) As Double
  If xValue > 0# Then
    P2Atan2 = Atn(yValue / xValue)
  ElseIf xValue < 0# Then
    If yValue >= 0# Then
      P2Atan2 = Atn(yValue / xValue) + 3.14159265358979
    Else
      P2Atan2 = Atn(yValue / xValue) - 3.14159265358979
    End If
  ElseIf yValue > 0# Then
    P2Atan2 = 1.5707963267949
  ElseIf yValue < 0# Then
    P2Atan2 = -1.5707963267949
  Else
    P2Atan2 = 0#
  End If
End Function

Private Sub P2SortPrincipal3(ByVal value1 As Double, ByVal value2 As Double, ByVal value3 As Double, ByRef maximumValue As Double, ByRef middleValue As Double, ByRef minimumValue As Double, ByRef maximumMode As Long, ByRef minimumMode As Long)
  Dim values(1 To 3) As Double, modes(1 To 3) As Long
  Dim i As Long, j As Long, tempValue As Double, tempMode As Long
  values(1) = value1: modes(1) = 1
  values(2) = value2: modes(2) = 2
  values(3) = value3: modes(3) = 3
  For i = 1 To 2
    For j = i + 1 To 3
      If values(i) < values(j) Then
        tempValue = values(i): values(i) = values(j): values(j) = tempValue
        tempMode = modes(i): modes(i) = modes(j): modes(j) = tempMode
      End If
    Next j
  Next i
  maximumValue = values(1)
  middleValue = values(2)
  minimumValue = values(3)
  maximumMode = modes(1)
  minimumMode = modes(3)
End Sub

Public Function P2VonMises3D(ByVal stressX As Double, ByVal stressY As Double, ByVal stressZ As Double, ByVal shearXY As Double) As Double
  P2VonMises3D = Sqr(0.5 * ((stressX - stressY) ^ 2 + (stressY - stressZ) ^ 2 + (stressZ - stressX) ^ 2) + 3# * shearXY ^ 2)
End Function

Private Function P2EvaluateYield(ByVal sx As Double, ByVal sy As Double, ByVal sz As Double, ByVal txy As Double, ByVal sinPhi As Double, ByVal cohesion As Double, ByVal cosPhi As Double, ByRef yieldValue As Double, ByRef sigmaMax As Double, ByRef sigmaMid As Double, ByRef sigmaMin As Double, ByRef principalAngle As Double, ByRef radius As Double, ByRef maximumMode As Long, ByRef minimumMode As Long) As Boolean
  Dim meanStress As Double, deviatoricNormal As Double
  Dim planeSigma1 As Double, planeSigma2 As Double
  On Error GoTo Failed
  meanStress = 0.5 * (sx + sy)
  deviatoricNormal = 0.5 * (sx - sy)
  radius = Sqr(deviatoricNormal * deviatoricNormal + txy * txy)
  planeSigma1 = meanStress + radius
  planeSigma2 = meanStress - radius
  principalAngle = 0.5 * P2Atan2(2# * txy, sx - sy)
  P2SortPrincipal3 planeSigma1, planeSigma2, sz, sigmaMax, sigmaMid, sigmaMin, maximumMode, minimumMode
  yieldValue = sigmaMax - sigmaMin - (sigmaMax + sigmaMin) * sinPhi - 2# * cohesion * cosPhi
  If Not P2IsFinite(yieldValue) Then Exit Function
  If Not P2IsFinite(sigmaMax) Then Exit Function
  If Not P2IsFinite(sigmaMid) Then Exit Function
  If Not P2IsFinite(sigmaMin) Then Exit Function
  If Not P2IsFinite(principalAngle) Then Exit Function
  P2EvaluateYield = True
  Exit Function
Failed:
  P2EvaluateYield = False
End Function

Private Sub P2FlowGradients3D(ByVal principalAngle As Double, ByVal stressX As Double, ByVal stressY As Double, ByVal stressZ As Double, ByVal shearXY As Double, ByVal maximumMode As Long, ByVal minimumMode As Long, ByVal sigmaMaxValue As Double, ByVal sigmaMinValue As Double, ByVal sinPhi As Double, ByVal sinPsi As Double, ByRef nx As Double, ByRef ny As Double, ByRef nz As Double, ByRef nt As Double, ByRef mx As Double, ByRef my As Double, ByRef mz As Double, ByRef mt As Double)
  Dim cAngle As Double, sAngle As Double, c2 As Double, s2 As Double, cs As Double
  Dim n1 As Double, n3 As Double, m1 As Double, m3 As Double
  Dim nPlane1 As Double, nPlane2 As Double, nZValue As Double
  Dim mPlane1 As Double, mPlane2 As Double, mZValue As Double
  Dim meanStress As Double, deviatoricNormal As Double, radiusValue As Double
  Dim principalValue(1 To 3) As Double, i As Long, maxCount As Long, minCount As Long
  Dim tieTolerance As Double, maxWeight As Double, minWeight As Double
  cAngle = Cos(principalAngle)
  sAngle = Sin(principalAngle)
  c2 = cAngle * cAngle
  s2 = sAngle * sAngle
  cs = cAngle * sAngle
  n1 = 1# - sinPhi
  n3 = -1# - sinPhi
  m1 = 1# - sinPsi
  m3 = -1# - sinPsi
  meanStress = 0.5 * (stressX + stressY)
  deviatoricNormal = 0.5 * (stressX - stressY)
  radiusValue = Sqr(deviatoricNormal * deviatoricNormal + shearXY * shearXY)
  principalValue(1) = meanStress + radiusValue
  principalValue(2) = meanStress - radiusValue
  principalValue(3) = stressZ
  tieTolerance = 0.0000000001 * (Abs(sigmaMaxValue) + Abs(sigmaMinValue) + Abs(stressZ) + 1#)
  For i = 1 To 3
    If Abs(principalValue(i) - sigmaMaxValue) <= tieTolerance Then maxCount = maxCount + 1
    If Abs(principalValue(i) - sigmaMinValue) <= tieTolerance Then minCount = minCount + 1
  Next i
  If maxCount < 1 Then maxCount = 1
  If minCount < 1 Then minCount = 1
  nPlane1 = 0#: nPlane2 = 0#: nZValue = 0#
  mPlane1 = 0#: mPlane2 = 0#: mZValue = 0#
  For i = 1 To 3
    If Abs(principalValue(i) - sigmaMaxValue) <= tieTolerance Then
      maxWeight = n1 / maxCount
      If i = 1 Then
        nPlane1 = nPlane1 + maxWeight
        mPlane1 = mPlane1 + m1 / maxCount
      ElseIf i = 2 Then
        nPlane2 = nPlane2 + maxWeight
        mPlane2 = mPlane2 + m1 / maxCount
      Else
        nZValue = nZValue + maxWeight
        mZValue = mZValue + m1 / maxCount
      End If
    End If
    If Abs(principalValue(i) - sigmaMinValue) <= tieTolerance Then
      minWeight = n3 / minCount
      If i = 1 Then
        nPlane1 = nPlane1 + minWeight
        mPlane1 = mPlane1 + m3 / minCount
      ElseIf i = 2 Then
        nPlane2 = nPlane2 + minWeight
        mPlane2 = mPlane2 + m3 / minCount
      Else
        nZValue = nZValue + minWeight
        mZValue = mZValue + m3 / minCount
      End If
    End If
  Next i
  nx = nPlane1 * c2 + nPlane2 * s2
  ny = nPlane1 * s2 + nPlane2 * c2
  nz = nZValue
  nt = 2# * cs * (nPlane1 - nPlane2)
  mx = mPlane1 * c2 + mPlane2 * s2
  my = mPlane1 * s2 + mPlane2 * c2
  mz = mZValue
  mt = 2# * cs * (mPlane1 - mPlane2)
End Sub

Private Sub P2SetFailure(ByRef outputState As P2_MaterialPointOutput, ByVal failureCode As Long, ByVal failureMessage As String)
  outputState.converged = False
  outputState.Elastic = False
  outputState.failureCode = failureCode
  outputState.failureMessage = failureMessage
End Sub

Private Sub P2CopyOutput(ByRef sourceState As P2_MaterialPointOutput, ByRef targetState As P2_MaterialPointOutput)
  Dim i As Long, j As Long
  For i = 0 To 3
    targetState.TrialStress(i) = sourceState.TrialStress(i)
    targetState.Stress(i) = sourceState.Stress(i)
    targetState.PlasticStrain(i) = sourceState.PlasticStrain(i)
  Next i
  For i = 0 To 2
    For j = 0 To 2
      targetState.Tangent(i, j) = sourceState.Tangent(i, j)
    Next j
  Next i
  targetState.PlasticMultiplier = sourceState.PlasticMultiplier
  targetState.TrialYieldFunction = sourceState.TrialYieldFunction
  targetState.YieldFunction = sourceState.YieldFunction
  targetState.PrincipalStress(0) = sourceState.PrincipalStress(0)
  targetState.PrincipalStress(1) = sourceState.PrincipalStress(1)
  targetState.PrincipalStress(2) = sourceState.PrincipalStress(2)
  targetState.principalAngle = sourceState.principalAngle
  targetState.yielded = sourceState.yielded
  targetState.PlasticOccurred = sourceState.PlasticOccurred
  targetState.converged = sourceState.converged
  targetState.Elastic = sourceState.Elastic
  targetState.failureCode = sourceState.failureCode
  targetState.failureMessage = sourceState.failureMessage
  targetState.iterations = sourceState.iterations
  targetState.UsedSubsteps = sourceState.UsedSubsteps
End Sub

Private Function P2MaterialConstantsAreValid(ByVal Young As Double, ByVal Poisson As Double, ByVal frictionAngle As Double, ByVal cohesion As Double, ByVal dilationAngle As Double, ByVal tolerance As Double, ByRef failMessage As String) As Boolean
  P2MaterialConstantsAreValid = False
  If Young <= 0# Then
    failMessage = "ヤング率は正で指定してください。"
    Exit Function
  End If
  If Poisson <= -1# Or Poisson >= 0.5 Then
    failMessage = "ポアソン比が平面ひずみの範囲外です。値=" & Format$(Poisson, "0.##########")
    Exit Function
  End If
  If frictionAngle < 0# Then
    failMessage = "摩擦角は0度以上で指定してください。"
    Exit Function
  End If
  If frictionAngle >= 89.999 Then
    failMessage = "摩擦角が90度以上または特異点に近すぎます。"
    Exit Function
  End If
  If dilationAngle < 0# Then
    failMessage = "ダイレタンシー角は0度以上で指定してください。"
    Exit Function
  End If
  If dilationAngle >= 89.999 Then
    failMessage = "ダイレタンシー角が90度以上または特異点に近すぎます。"
    Exit Function
  End If
  If cohesion < 0# Then
    failMessage = "粘着力は0以上で指定してください。"
    Exit Function
  End If
  If tolerance <= 0# Then
    failMessage = "材料点許容誤差は正で指定してください。"
    Exit Function
  End If
  P2MaterialConstantsAreValid = True
End Function

Private Function P2ValidateInput(ByRef inputState As P2_MaterialPointInput, ByRef outputState As P2_MaterialPointOutput) As Boolean
  Dim i As Long, failMessage As String
  P2ValidateInput = False
  If Not inputState.UseCachedConstants Then
    If Not P2MaterialConstantsAreValid(inputState.Young, inputState.Poisson, inputState.frictionAngle, inputState.cohesion, inputState.dilationAngle, inputState.tolerance, failMessage) Then
      P2SetFailure outputState, P2_FAILURE_INVALID_INPUT, failMessage
      Exit Function
    End If
  End If
  For i = 0 To 3
    If Not P2IsFinite(inputState.PreviousStress(i)) Then
      P2SetFailure outputState, P2_FAILURE_INVALID_INPUT, "平面ひずみの確定応力に有限でない値があります。"
      Exit Function
    End If
    If i <= 2 Then
      If Not P2IsFinite(inputState.StrainIncrement(i)) Then
        P2SetFailure outputState, P2_FAILURE_INVALID_INPUT, "面内ひずみ増分に有限でない値があります。"
        Exit Function
      End If
    End If
  Next i
  P2ValidateInput = True
End Function

Private Sub P2PrincipalGradient3D(ByVal principalAngle As Double, ByVal principalMode As Long, ByRef gx As Double, ByRef gy As Double, ByRef gz As Double, ByRef gt As Double)
  Dim cAngle As Double, sAngle As Double
  cAngle = Cos(principalAngle)
  sAngle = Sin(principalAngle)
  gx = 0#: gy = 0#: gz = 0#: gt = 0#
  Select Case principalMode
    Case 1
      gx = cAngle * cAngle
      gy = sAngle * sAngle
      gt = 2# * cAngle * sAngle
    Case 2
      gx = sAngle * sAngle
      gy = cAngle * cAngle
      gt = -2# * cAngle * sAngle
    Case 3
      gz = 1#
  End Select
End Sub

Private Sub P2BuildFaceGradient3D(ByVal principalAngle As Double, ByVal maximumMode As Long, ByVal minimumMode As Long, ByVal sinValue As Double, ByRef gx As Double, ByRef gy As Double, ByRef gz As Double, ByRef gt As Double)
  Dim maxX As Double, maxY As Double, maxZ As Double, maxT As Double
  Dim minX As Double, minY As Double, minZ As Double, minT As Double
  P2PrincipalGradient3D principalAngle, maximumMode, maxX, maxY, maxZ, maxT
  P2PrincipalGradient3D principalAngle, minimumMode, minX, minY, minZ, minT
  gx = (1# - sinValue) * maxX + (-1# - sinValue) * minX
  gy = (1# - sinValue) * maxY + (-1# - sinValue) * minY
  gz = (1# - sinValue) * maxZ + (-1# - sinValue) * minZ
  gt = (1# - sinValue) * maxT + (-1# - sinValue) * minT
End Sub

Private Sub P2ElasticApplyGradient(ByVal gx As Double, ByVal gy As Double, ByVal gz As Double, ByVal gt As Double, ByVal d00 As Double, ByVal d01 As Double, ByVal d22 As Double, ByRef qx As Double, ByRef qy As Double, ByRef qz As Double, ByRef qt As Double)
  qx = d00 * gx + d01 * gy + d01 * gz
  qy = d01 * gx + d00 * gy + d01 * gz
  qz = d01 * gx + d01 * gy + d00 * gz
  qt = d22 * gt
End Sub

Private Function P2ModeValue(ByVal principalMode As Long, ByVal sigmaMaxValue As Double, ByVal sigmaMidValue As Double, ByVal sigmaMinValue As Double) As Double
  Select Case principalMode
    Case 1
      P2ModeValue = sigmaMaxValue
    Case 2
      P2ModeValue = sigmaMidValue
    Case 3
      P2ModeValue = sigmaMinValue
    Case Else
      P2ModeValue = 0#
  End Select
End Function

' Yield of the NEW Mohr-Coulomb face after a corner switch, evaluated at the trial stress (not the old face's yieldTrial).
Private Function P2FaceYieldAtStress(ByVal maximumMode As Long, ByVal minimumMode As Long, ByVal stressX As Double, ByVal stressY As Double, ByVal stressZ As Double, ByVal shearXY As Double, ByVal principalAngle As Double, ByVal sinPhi As Double, ByVal cohesion As Double, ByVal cosPhi As Double) As Double
  Dim sMax As Double, sMin As Double
  sMax = P2PhysicalModeValue(maximumMode, stressX, stressY, stressZ, shearXY, principalAngle)
  sMin = P2PhysicalModeValue(minimumMode, stressX, stressY, stressZ, shearXY, principalAngle)
  P2FaceYieldAtStress = sMax - sMin - (sMax + sMin) * sinPhi - 2# * cohesion * cosPhi
End Function

Private Function P2PhysicalModeValue(ByVal principalMode As Long, ByVal stressX As Double, ByVal stressY As Double, ByVal stressZ As Double, ByVal shearXY As Double, ByVal principalAngle As Double) As Double
  Dim meanStress As Double, deviatoricNormal As Double, cAngle As Double, sAngle As Double
  meanStress = 0.5 * (stressX + stressY)
  deviatoricNormal = 0.5 * (stressX - stressY)
  cAngle = Cos(principalAngle)
  sAngle = Sin(principalAngle)
  Select Case principalMode
    Case 1
      P2PhysicalModeValue = meanStress + deviatoricNormal * (cAngle * cAngle - sAngle * sAngle) + 2# * shearXY * cAngle * sAngle
    Case 2
      P2PhysicalModeValue = meanStress - deviatoricNormal * (cAngle * cAngle - sAngle * sAngle) - 2# * shearXY * cAngle * sAngle
    Case 3
      P2PhysicalModeValue = stressZ
    Case Else
      P2PhysicalModeValue = 0#
  End Select
End Function

Private Function P2TryEdgeReturn(ByRef inputState As P2_MaterialPointInput, ByVal trialSx As Double, ByVal trialSy As Double, ByVal trialSz As Double, ByVal trialTxy As Double, ByVal d00 As Double, ByVal d01 As Double, ByVal d22 As Double, ByVal sinPhi As Double, ByVal cosPhi As Double, ByVal sinPsi As Double, ByVal principalAngle As Double, ByVal trialSigmaMax As Double, ByVal trialSigmaMid As Double, ByVal trialSigmaMin As Double, ByVal initialMaximumMode As Long, ByVal initialMinimumMode As Long, ByVal currentMaximumMode As Long, ByVal currentMinimumMode As Long, ByRef outputState As P2_MaterialPointOutput) As Boolean
  Dim middleMode As Long, edgeModeA As Long, edgeModeB As Long, edgeModeC As Long
  Dim upperEdge As Boolean, lowerEdge As Boolean
  Dim valueA As Double, valueB As Double, valueC As Double
  Dim yieldAtTrial As Double, equalityAtTrial As Double
  Dim n1x As Double, n1y As Double, n1z As Double, n1t As Double
  Dim m1x As Double, m1y As Double, m1z As Double, m1t As Double
  Dim n2x As Double, n2y As Double, n2z As Double, n2t As Double
  Dim m2x As Double, m2y As Double, m2z As Double, m2t As Double
  Dim eqx As Double, eqy As Double, eqz As Double, eqt As Double
  Dim q1x As Double, q1y As Double, q1z As Double, q1t As Double
  Dim q2x As Double, q2y As Double, q2z As Double, q2t As Double
  Dim a11 As Double, a12 As Double, a21 As Double, a22 As Double, determinant As Double
  Dim lambda1 As Double, lambda2 As Double, lambdaTolerance As Double
  Dim stressX As Double, stressY As Double, stressZ As Double, shearXY As Double
  Dim finalYield As Double, finalMax As Double, finalMid As Double, finalMin As Double
  Dim finalAngle As Double, finalRadius As Double, finalMaxMode As Long, finalMinMode As Long
  Dim correctedA As Double, correctedB As Double, correctedC As Double, edgeTolerance As Double
  Dim elasticMap(0 To 3, 0 To 2) As Double
  Dim b1 As Double, b2 As Double, dl1 As Double, dl2 As Double
  Dim col As Long
  On Error GoTo Failed
  P2TryEdgeReturn = False
  middleMode = 6 - initialMaximumMode - initialMinimumMode
  If currentMaximumMode <> initialMaximumMode And currentMinimumMode = initialMinimumMode Then
    upperEdge = True
  ElseIf currentMinimumMode <> initialMinimumMode And currentMaximumMode = initialMaximumMode Then
    lowerEdge = True
  Else
    Exit Function
  End If
  edgeModeA = initialMaximumMode
  edgeModeB = middleMode
  edgeModeC = initialMinimumMode
  valueA = P2PhysicalModeValue(edgeModeA, trialSx, trialSy, trialSz, trialTxy, principalAngle)
  valueB = P2PhysicalModeValue(edgeModeB, trialSx, trialSy, trialSz, trialTxy, principalAngle)
  valueC = P2PhysicalModeValue(edgeModeC, trialSx, trialSy, trialSz, trialTxy, principalAngle)
  yieldAtTrial = valueA - valueC - (valueA + valueC) * sinPhi - 2# * inputState.cohesion * cosPhi
  If upperEdge Then
    equalityAtTrial = valueA - valueB
  Else
    equalityAtTrial = valueB - valueC
  End If
  P2BuildFaceGradient3D principalAngle, edgeModeA, edgeModeC, sinPhi, n1x, n1y, n1z, n1t
  P2BuildFaceGradient3D principalAngle, edgeModeA, edgeModeC, sinPsi, m1x, m1y, m1z, m1t
  If upperEdge Then
    P2BuildFaceGradient3D principalAngle, edgeModeB, edgeModeC, sinPhi, n2x, n2y, n2z, n2t
    P2BuildFaceGradient3D principalAngle, edgeModeB, edgeModeC, sinPsi, m2x, m2y, m2z, m2t
    P2PrincipalGradient3D principalAngle, edgeModeA, eqx, eqy, eqz, eqt
    P2PrincipalGradient3D principalAngle, edgeModeB, b1, b2, dl1, dl2
    eqx = eqx - b1: eqy = eqy - b2: eqz = eqz - dl1: eqt = eqt - dl2
  Else
    P2BuildFaceGradient3D principalAngle, edgeModeA, edgeModeB, sinPhi, n2x, n2y, n2z, n2t
    P2BuildFaceGradient3D principalAngle, edgeModeA, edgeModeB, sinPsi, m2x, m2y, m2z, m2t
    P2PrincipalGradient3D principalAngle, edgeModeB, eqx, eqy, eqz, eqt
    P2PrincipalGradient3D principalAngle, edgeModeC, b1, b2, dl1, dl2
    eqx = eqx - b1: eqy = eqy - b2: eqz = eqz - dl1: eqt = eqt - dl2
  End If
  P2ElasticApplyGradient m1x, m1y, m1z, m1t, d00, d01, d22, q1x, q1y, q1z, q1t
  P2ElasticApplyGradient m2x, m2y, m2z, m2t, d00, d01, d22, q2x, q2y, q2z, q2t
  a11 = n1x * q1x + n1y * q1y + n1z * q1z + n1t * q1t
  a12 = n1x * q2x + n1y * q2y + n1z * q2z + n1t * q2t
  a21 = eqx * q1x + eqy * q1y + eqz * q1z + eqt * q1t
  a22 = eqx * q2x + eqy * q2y + eqz * q2z + eqt * q2t
  determinant = a11 * a22 - a12 * a21
  If Abs(determinant) <= 0.00000000000001 * (Abs(a11) + Abs(a12) + Abs(a21) + Abs(a22) + 1#) Then Exit Function
  lambda1 = (yieldAtTrial * a22 - a12 * equalityAtTrial) / determinant
  lambda2 = (a11 * equalityAtTrial - a21 * yieldAtTrial) / determinant
  lambdaTolerance = 0.0000000001 * (Abs(yieldAtTrial) + Abs(equalityAtTrial) + 1#)
  If lambda1 < -lambdaTolerance Or lambda2 < -lambdaTolerance Then Exit Function
  If lambda1 < 0# Then lambda1 = 0#
  If lambda2 < 0# Then lambda2 = 0#
  stressX = trialSx - lambda1 * q1x - lambda2 * q2x
  stressY = trialSy - lambda1 * q1y - lambda2 * q2y
  stressZ = trialSz - lambda1 * q1z - lambda2 * q2z
  shearXY = trialTxy - lambda1 * q1t - lambda2 * q2t
  If Not P2EvaluateYield(stressX, stressY, stressZ, shearXY, sinPhi, inputState.cohesion, cosPhi, finalYield, finalMax, finalMid, finalMin, finalAngle, finalRadius, finalMaxMode, finalMinMode) Then Exit Function
  correctedA = P2PhysicalModeValue(edgeModeA, stressX, stressY, stressZ, shearXY, finalAngle)
  correctedB = P2PhysicalModeValue(edgeModeB, stressX, stressY, stressZ, shearXY, finalAngle)
  correctedC = P2PhysicalModeValue(edgeModeC, stressX, stressY, stressZ, shearXY, finalAngle)
  edgeTolerance = inputState.tolerance * (Abs(finalMax) + Abs(finalMin) + 2# * inputState.cohesion * cosPhi + 1#) * 10#
  If Abs(finalYield) > edgeTolerance Then Exit Function
  If upperEdge Then
    If Abs(correctedA - correctedB) > edgeTolerance Or correctedB < correctedC - edgeTolerance Then Exit Function
  Else
    If correctedA < correctedB - edgeTolerance Or Abs(correctedB - correctedC) > edgeTolerance Then Exit Function
  End If
  elasticMap(0, 0) = d00: elasticMap(0, 1) = d01: elasticMap(0, 2) = 0#
  elasticMap(1, 0) = d01: elasticMap(1, 1) = d00: elasticMap(1, 2) = 0#
  elasticMap(2, 0) = d01: elasticMap(2, 1) = d01: elasticMap(2, 2) = 0#
  elasticMap(3, 0) = 0#: elasticMap(3, 1) = 0#: elasticMap(3, 2) = d22
  For col = 0 To 2
    b1 = n1x * elasticMap(0, col) + n1y * elasticMap(1, col) + n1z * elasticMap(2, col) + n1t * elasticMap(3, col)
    b2 = eqx * elasticMap(0, col) + eqy * elasticMap(1, col) + eqz * elasticMap(2, col) + eqt * elasticMap(3, col)
    dl1 = (a22 * b1 - a12 * b2) / determinant
    dl2 = (-a21 * b1 + a11 * b2) / determinant
    outputState.Tangent(0, col) = elasticMap(0, col) - q1x * dl1 - q2x * dl2
    outputState.Tangent(1, col) = elasticMap(1, col) - q1y * dl1 - q2y * dl2
    outputState.Tangent(2, col) = elasticMap(3, col) - q1t * dl1 - q2t * dl2
  Next col
  outputState.Stress(0) = stressX
  outputState.Stress(1) = stressY
  outputState.Stress(2) = shearXY
  outputState.Stress(3) = stressZ
  outputState.PlasticStrain(0) = lambda1 * m1x + lambda2 * m2x
  outputState.PlasticStrain(1) = lambda1 * m1y + lambda2 * m2y
  outputState.PlasticStrain(2) = lambda1 * m1t + lambda2 * m2t
  outputState.PlasticStrain(3) = lambda1 * m1z + lambda2 * m2z
  outputState.PlasticMultiplier = lambda1 + lambda2
  outputState.PrincipalStress(0) = finalMax
  outputState.PrincipalStress(1) = finalMid
  outputState.PrincipalStress(2) = finalMin
  outputState.principalAngle = finalAngle * 180# / 3.14159265358979
  outputState.YieldFunction = finalYield
  outputState.yielded = True
  outputState.Elastic = False
  outputState.converged = True
  outputState.failureCode = 0
  outputState.failureMessage = "Mohr-Coulomb 3Dエッジリターン。"
  outputState.iterations = 1
  P2TryEdgeReturn = True
  Exit Function
Failed:
  P2TryEdgeReturn = False
End Function

Private Function P2IsStressAtVertex(ByVal sx As Double, ByVal sy As Double, ByVal sz As Double, ByVal txy As Double, ByVal vertexStress As Double, ByVal toleranceValue As Double, ByVal scaleValue As Double) As Boolean
  Dim residualValue As Double
  residualValue = Abs(sx - vertexStress) + Abs(sy - vertexStress) + Abs(sz - vertexStress) + Abs(txy)
  P2IsStressAtVertex = (residualValue <= toleranceValue * scaleValue * 100#)
End Function

Private Sub P2ComplianceFromStressDelta(ByVal youngModulus As Double, ByVal poissonRatio As Double, ByVal shearModulus As Double, ByVal deltaSx As Double, ByVal deltaSy As Double, ByVal deltaSz As Double, ByVal deltaTxy As Double, ByRef elasticEx As Double, ByRef elasticEy As Double, ByRef elasticEz As Double, ByRef elasticGamma As Double)
  elasticEx = (deltaSx - poissonRatio * (deltaSy + deltaSz)) / youngModulus
  elasticEy = (deltaSy - poissonRatio * (deltaSx + deltaSz)) / youngModulus
  elasticEz = (deltaSz - poissonRatio * (deltaSx + deltaSy)) / youngModulus
  If Abs(shearModulus) <= 1E-30 Then
    elasticGamma = 0#
  Else
    elasticGamma = deltaTxy / shearModulus
  End If
End Sub

Private Function P2PlasticStrainInVertexCone(ByVal trialSx As Double, ByVal trialSy As Double, ByVal trialSz As Double, ByVal trialTxy As Double, ByVal sinPsi As Double, ByVal plasticEx As Double, ByVal plasticEy As Double, ByVal plasticGamma As Double, ByVal plasticEz As Double) As Boolean
  Dim ang As Double, cAngle As Double, sAngle As Double, c2 As Double, s2 As Double, cs As Double
  Dim m1 As Double, m3 As Double, mPlane1 As Double, mPlane2 As Double, mZValue As Double
  Dim flowM(0 To 3, 0 To 5) As Double
  Dim pairMax(0 To 5) As Long, pairMin(0 To 5) As Long
  Dim a(0 To 3, 0 To 5) As Double, ata() As Double, rhs() As Double, lam() As Double
  Dim target(0 To 3) As Double, pred(0 To 3) As Double
  Dim mask As Long, bitValue As Long, columnCount As Long, pairId As Long, rowId As Long, colId As Long, otherId As Long
  Dim residualValue As Double, bestResidual As Double, plasticNorm As Double, projValue As Double
  Dim lambdaOK As Boolean
  P2PlasticStrainInVertexCone = False
  target(0) = plasticEx: target(1) = plasticEy: target(2) = plasticGamma: target(3) = plasticEz
  plasticNorm = Abs(plasticEx) + Abs(plasticEy) + Abs(plasticGamma) + Abs(plasticEz)
  If plasticNorm <= 1E-30 Then
    P2PlasticStrainInVertexCone = True
    Exit Function
  End If
  ang = 0.5 * P2Atan2(2# * trialTxy, trialSx - trialSy)
  cAngle = Cos(ang): sAngle = Sin(ang)
  c2 = cAngle * cAngle: s2 = sAngle * sAngle: cs = cAngle * sAngle
  m1 = 1# - sinPsi
  m3 = -1# - sinPsi
  pairMax(0) = 1: pairMin(0) = 2
  pairMax(1) = 1: pairMin(1) = 3
  pairMax(2) = 2: pairMin(2) = 1
  pairMax(3) = 2: pairMin(3) = 3
  pairMax(4) = 3: pairMin(4) = 1
  pairMax(5) = 3: pairMin(5) = 2
  For pairId = 0 To 5
    mPlane1 = 0#: mPlane2 = 0#: mZValue = 0#
    If pairMax(pairId) = 1 Then mPlane1 = m1
    If pairMax(pairId) = 2 Then mPlane2 = m1
    If pairMax(pairId) = 3 Then mZValue = m1
    If pairMin(pairId) = 1 Then mPlane1 = mPlane1 + m3
    If pairMin(pairId) = 2 Then mPlane2 = mPlane2 + m3
    If pairMin(pairId) = 3 Then mZValue = mZValue + m3
    flowM(0, pairId) = mPlane1 * c2 + mPlane2 * s2
    flowM(1, pairId) = mPlane1 * s2 + mPlane2 * c2
    flowM(2, pairId) = 2# * cs * (mPlane1 - mPlane2)
    flowM(3, pairId) = mZValue
  Next pairId
  bestResidual = 1E+300
  For mask = 1 To 63
    columnCount = 0
    bitValue = 1
    For pairId = 0 To 5
      If (mask And bitValue) <> 0 Then
        For rowId = 0 To 3
          a(rowId, columnCount) = flowM(rowId, pairId)
        Next rowId
        columnCount = columnCount + 1
      End If
      bitValue = bitValue * 2
    Next pairId
    ReDim ata(0 To columnCount - 1, 0 To columnCount - 1)
    ReDim rhs(0 To columnCount - 1)
    For colId = 0 To columnCount - 1
      projValue = 0#
      For rowId = 0 To 3
        projValue = projValue + a(rowId, colId) * target(rowId)
      Next rowId
      rhs(colId) = projValue
      For otherId = 0 To columnCount - 1
        projValue = 0#
        For rowId = 0 To 3
          projValue = projValue + a(rowId, colId) * a(rowId, otherId)
        Next rowId
        ata(colId, otherId) = projValue
      Next otherId
    Next colId
    If Not P6InvertDense(ata, columnCount) Then GoTo NextMask
    ReDim lam(0 To columnCount - 1)
    lambdaOK = True
    For colId = 0 To columnCount - 1
      projValue = 0#
      For otherId = 0 To columnCount - 1
        projValue = projValue + ata(colId, otherId) * rhs(otherId)
      Next otherId
      If projValue < -0.000000001 Then
        lambdaOK = False
        Exit For
      End If
      If projValue < 0# Then projValue = 0#
      lam(colId) = projValue
    Next colId
    If Not lambdaOK Then GoTo NextMask
    pred(0) = 0#: pred(1) = 0#: pred(2) = 0#: pred(3) = 0#
    For colId = 0 To columnCount - 1
      For rowId = 0 To 3
        pred(rowId) = pred(rowId) + a(rowId, colId) * lam(colId)
      Next rowId
    Next colId
    residualValue = Abs(pred(0) - target(0)) + Abs(pred(1) - target(1)) + Abs(pred(2) - target(2)) + Abs(pred(3) - target(3))
    If residualValue < bestResidual Then bestResidual = residualValue
NextMask:
  Next mask
  If bestResidual <= 0.00000001 * (plasticNorm + 0.000000000001) Then
    P2PlasticStrainInVertexCone = True
  End If
End Function

Private Function P2TryVertexReturn(ByRef inputState As P2_MaterialPointInput, ByVal trialSx As Double, ByVal trialSy As Double, ByVal trialSz As Double, ByVal trialTxy As Double, ByVal d00 As Double, ByVal d01 As Double, ByVal d22 As Double, ByVal sinPhi As Double, ByVal cosPhi As Double, ByVal sinPsi As Double, ByRef outputState As P2_MaterialPointOutput) As Boolean
  Dim vertexStress As Double, trialMean As Double, scaleValue As Double
  Dim deltaSx As Double, deltaSy As Double, deltaSz As Double, deltaTxy As Double
  Dim plasticEx As Double, plasticEy As Double, plasticEz As Double, plasticGamma As Double
  Dim plasticVol As Double, plasticNorm As Double
  Dim checkSx As Double, checkSy As Double, checkSz As Double, checkTxy As Double
  Dim residualValue As Double
  P2TryVertexReturn = False
  If sinPhi <= 0.000000000001 Then Exit Function
  If inputState.Young <= 0# Then Exit Function
  vertexStress = -inputState.cohesion * cosPhi / sinPhi
  trialMean = (trialSx + trialSy + trialSz) / 3#
  scaleValue = Abs(vertexStress) + Abs(trialMean) + Abs(inputState.cohesion) + 1#
  ' 圧縮正では頂点より引張側（平均がより小さい）だけが頂点錐の補空間。
  If trialMean > vertexStress + inputState.tolerance * scaleValue Then Exit Function
  deltaSx = trialSx - vertexStress
  deltaSy = trialSy - vertexStress
  deltaSz = trialSz - vertexStress
  deltaTxy = trialTxy
  P2ComplianceFromStressDelta inputState.Young, inputState.Poisson, d22, deltaSx, deltaSy, deltaSz, deltaTxy, plasticEx, plasticEy, plasticEz, plasticGamma
  plasticVol = plasticEx + plasticEy + plasticEz
  plasticNorm = Abs(plasticEx) + Abs(plasticEy) + Abs(plasticEz) + Abs(plasticGamma)
  If plasticNorm < 1# Then plasticNorm = 1#
  If Abs(sinPsi) <= 0.000000000001 Then
    If Abs(plasticVol) > 0.000000001 * plasticNorm And Abs(plasticVol) > inputState.tolerance Then Exit Function
  End If
  If Not P2PlasticStrainInVertexCone(trialSx, trialSy, trialSz, trialTxy, sinPsi, plasticEx, plasticEy, plasticGamma, plasticEz) Then Exit Function
  checkSx = inputState.PreviousStress(0) + d00 * (inputState.StrainIncrement(0) - plasticEx) + d01 * (inputState.StrainIncrement(1) - plasticEy) + d01 * (0# - plasticEz)
  checkSy = inputState.PreviousStress(1) + d01 * (inputState.StrainIncrement(0) - plasticEx) + d00 * (inputState.StrainIncrement(1) - plasticEy) + d01 * (0# - plasticEz)
  checkSz = inputState.PreviousStress(3) + d01 * (inputState.StrainIncrement(0) - plasticEx) + d01 * (inputState.StrainIncrement(1) - plasticEy) + d00 * (0# - plasticEz)
  checkTxy = inputState.PreviousStress(2) + d22 * (inputState.StrainIncrement(2) - plasticGamma)
  residualValue = Abs(checkSx - vertexStress) + Abs(checkSy - vertexStress) + Abs(checkSz - vertexStress) + Abs(checkTxy)
  If residualValue > 0.00000001 * scaleValue Then Exit Function
  outputState.Stress(0) = vertexStress
  outputState.Stress(1) = vertexStress
  outputState.Stress(2) = 0#
  outputState.Stress(3) = vertexStress
  outputState.PlasticStrain(0) = plasticEx
  outputState.PlasticStrain(1) = plasticEy
  outputState.PlasticStrain(2) = plasticGamma
  outputState.PlasticStrain(3) = plasticEz
  outputState.PlasticMultiplier = Sqr(plasticEx * plasticEx + plasticEy * plasticEy + 0.5 * plasticGamma * plasticGamma + plasticEz * plasticEz)
  outputState.Tangent(0, 0) = 0#: outputState.Tangent(0, 1) = 0#: outputState.Tangent(0, 2) = 0#
  outputState.Tangent(1, 0) = 0#: outputState.Tangent(1, 1) = 0#: outputState.Tangent(1, 2) = 0#
  outputState.Tangent(2, 0) = 0#: outputState.Tangent(2, 1) = 0#: outputState.Tangent(2, 2) = 0#
  outputState.PrincipalStress(0) = vertexStress
  outputState.PrincipalStress(1) = vertexStress
  outputState.PrincipalStress(2) = vertexStress
  outputState.principalAngle = 0#
  outputState.YieldFunction = 0#
  outputState.yielded = True
  outputState.Elastic = False
  outputState.converged = True
  outputState.failureCode = P2_FAILURE_VERTEX
  outputState.failureMessage = "Mohr-Coulomb 3D頂点への整合戻し。"
  outputState.iterations = 1
  P2TryVertexReturn = True
End Function

Private Function P2MaterialPointUpdateCore(ByRef inputState As P2_MaterialPointInput, ByRef outputState As P2_MaterialPointOutput)
  Dim d00 As Double, d01 As Double, d22 As Double, lameLambda As Double, shearModulus As Double
  Dim sinPhi As Double, cosPhi As Double, sinPsi As Double
  Dim sx As Double, sy As Double, sz As Double, txy As Double
  Dim trialSx As Double, trialSy As Double, trialSz As Double, trialTxy As Double
  Dim yieldTrial As Double, yieldCurrent As Double, yieldScale As Double
  Dim sigmaMax As Double, sigmaMid As Double, sigmaMin As Double, principalAngle As Double, radius As Double
  Dim trialSigmaMax As Double, trialSigmaMid As Double, trialSigmaMin As Double, trialPrincipalAngle As Double
  Dim maximumMode As Long, minimumMode As Long, initialMaximumMode As Long, initialMinimumMode As Long
  Dim nx As Double, ny As Double, nz As Double, nt As Double, mx As Double, my As Double, mz As Double, mt As Double
  Dim qx As Double, qy As Double, qz As Double, qt As Double
  Dim rx As Double, ry As Double, rt As Double, denominator As Double
  Dim lambdaValue As Double, deltaLambda As Double, lambdaTolerance As Double
  Dim iter As Long, maxIterations As Long, faceRestart As Long
  Dim piValue As Double
  On Error GoTo Failed
  P2MaterialPointUpdateCore = False
  piValue = 3.14159265358979
  If inputState.UseCachedConstants Then
    d00 = inputState.CachedElasticD00
    d01 = inputState.CachedElasticD01
    d22 = inputState.CachedElasticD22
    sinPhi = inputState.CachedSinFriction
    cosPhi = inputState.CachedCosFriction
    sinPsi = inputState.CachedSinDilation
  Else
    shearModulus = inputState.Young / (2# * (1# + inputState.Poisson))
    lameLambda = inputState.Young * inputState.Poisson / ((1# + inputState.Poisson) * (1# - 2# * inputState.Poisson))
    d00 = lameLambda + 2# * shearModulus
    d01 = lameLambda
    d22 = shearModulus
    sinPhi = Sin(inputState.frictionAngle * piValue / 180#)
    cosPhi = Cos(inputState.frictionAngle * piValue / 180#)
    sinPsi = Sin(inputState.dilationAngle * piValue / 180#)
  End If
  trialSx = inputState.PreviousStress(0) + d00 * inputState.StrainIncrement(0) + d01 * inputState.StrainIncrement(1)
  trialSy = inputState.PreviousStress(1) + d01 * inputState.StrainIncrement(0) + d00 * inputState.StrainIncrement(1)
  trialTxy = inputState.PreviousStress(2) + d22 * inputState.StrainIncrement(2)
  trialSz = inputState.PreviousStress(3) + d01 * (inputState.StrainIncrement(0) + inputState.StrainIncrement(1))
  outputState.TrialStress(0) = trialSx
  outputState.TrialStress(1) = trialSy
  outputState.TrialStress(2) = trialTxy
  outputState.TrialStress(3) = trialSz
  If Not P2EvaluateYield(trialSx, trialSy, trialSz, trialTxy, sinPhi, inputState.cohesion, cosPhi, yieldTrial, sigmaMax, sigmaMid, sigmaMin, principalAngle, radius, maximumMode, minimumMode) Then
    P2SetFailure outputState, P2_FAILURE_NONFINITE, "平面ひずみの弾性試行応力または3D降伏関数が有限ではありません。"
    Exit Function
  End If
  outputState.TrialYieldFunction = yieldTrial
  trialSigmaMax = sigmaMax
  trialSigmaMid = sigmaMid
  trialSigmaMin = sigmaMin
  trialPrincipalAngle = principalAngle
  yieldScale = Abs(sigmaMax) + Abs(sigmaMin) + 2# * inputState.cohesion * cosPhi
  If yieldScale < 1# Then yieldScale = 1#
  If yieldTrial <= inputState.tolerance * yieldScale Then
    outputState.Stress(0) = trialSx
    outputState.Stress(1) = trialSy
    outputState.Stress(2) = trialTxy
    outputState.Stress(3) = trialSz
    outputState.Tangent(0, 0) = d00: outputState.Tangent(0, 1) = d01: outputState.Tangent(0, 2) = 0#
    outputState.Tangent(1, 0) = d01: outputState.Tangent(1, 1) = d00: outputState.Tangent(1, 2) = 0#
    outputState.Tangent(2, 0) = 0#: outputState.Tangent(2, 1) = 0#: outputState.Tangent(2, 2) = d22
    outputState.PrincipalStress(0) = sigmaMax
    outputState.PrincipalStress(1) = sigmaMid
    outputState.PrincipalStress(2) = sigmaMin
    outputState.principalAngle = principalAngle * 180# / piValue
    outputState.YieldFunction = yieldTrial
    outputState.yielded = False
    outputState.Elastic = True
    outputState.converged = True
    outputState.iterations = 1
    P2MaterialPointUpdateCore = True
    Exit Function
  End If

  If P2IsStressAtVertex(inputState.PreviousStress(0), inputState.PreviousStress(1), inputState.PreviousStress(3), inputState.PreviousStress(2), -inputState.cohesion * cosPhi / (sinPhi + 1E-30), inputState.tolerance, yieldScale) Then
    If P2TryVertexReturn(inputState, trialSx, trialSy, trialSz, trialTxy, d00, d01, d22, sinPhi, cosPhi, sinPsi, outputState) Then
      P2MaterialPointUpdateCore = True
      Exit Function
    End If
  End If

  If radius <= 0.000000000001 * yieldScale And Abs(sigmaMax - sigmaMin) <= 0.000000000001 * yieldScale Then
    If P2TryVertexReturn(inputState, trialSx, trialSy, trialSz, trialTxy, d00, d01, d22, sinPhi, cosPhi, sinPsi, outputState) Then
      P2MaterialPointUpdateCore = True
      Exit Function
    End If
  End If

  P2FlowGradients3D principalAngle, trialSx, trialSy, trialSz, trialTxy, maximumMode, minimumMode, sigmaMax, sigmaMin, sinPhi, sinPsi, nx, ny, nz, nt, mx, my, mz, mt
  initialMaximumMode = maximumMode
  initialMinimumMode = minimumMode
  qx = d00 * mx + d01 * my + d01 * mz
  qy = d01 * mx + d00 * my + d01 * mz
  qz = d01 * mx + d01 * my + d00 * mz
  qt = d22 * mt
  rx = nx * d00 + ny * d01 + nz * d01
  ry = nx * d01 + ny * d00 + nz * d01
  rt = nt * d22
  denominator = nx * qx + ny * qy + nz * qz + nt * qt
  If Abs(denominator) <= 0.00000000000001 * yieldScale Then
    P2SetFailure outputState, P2_FAILURE_DENOMINATOR, "3D塑性補正の分母が0または小さすぎます。微小値を加えて続行しません。"
    Exit Function
  End If
  lambdaValue = yieldTrial / denominator
  If lambdaValue < 0# Then
    If Abs(lambdaValue) <= inputState.tolerance Then
      lambdaValue = 0#
    Else
      P2SetFailure outputState, P2_FAILURE_DENOMINATOR, "3D塑性乗数が負になりました。材料点の符号規約を確認してください。"
      Exit Function
    End If
  End If
  maxIterations = inputState.maxIterations
  If maxIterations <= 0 Then maxIterations = P2_DEFAULT_MAX_ITERATIONS
  If maxIterations > 100 Then maxIterations = 100
  lambdaTolerance = inputState.tolerance / (Abs(denominator) + 1#)
  For iter = 1 To maxIterations
NextFaceIteration:
    sx = trialSx - lambdaValue * qx
    sy = trialSy - lambdaValue * qy
    txy = trialTxy - lambdaValue * qt
    sz = trialSz - lambdaValue * qz
    If Not P2EvaluateYield(sx, sy, sz, txy, sinPhi, inputState.cohesion, cosPhi, yieldCurrent, sigmaMax, sigmaMid, sigmaMin, principalAngle, radius, maximumMode, minimumMode) Then
      P2SetFailure outputState, P2_FAILURE_NONFINITE, "3D塑性補正後の応力または降伏関数が有限ではありません。"
      Exit Function
    End If
    If maximumMode <> initialMaximumMode Or minimumMode <> initialMinimumMode Then
      If P2TryEdgeReturn(inputState, trialSx, trialSy, trialSz, trialTxy, d00, d01, d22, sinPhi, cosPhi, sinPsi, trialPrincipalAngle, trialSigmaMax, trialSigmaMid, trialSigmaMin, initialMaximumMode, initialMinimumMode, maximumMode, minimumMode, outputState) Then
        outputState.iterations = iter
        P2MaterialPointUpdateCore = True
        Exit Function
      End If
      If P2TryVertexReturn(inputState, trialSx, trialSy, trialSz, trialTxy, d00, d01, d22, sinPhi, cosPhi, sinPsi, outputState) Then
        outputState.iterations = iter
        P2MaterialPointUpdateCore = True
        Exit Function
      End If
      faceRestart = faceRestart + 1
      If faceRestart > 3 Then
        P2SetFailure outputState, P2_FAILURE_NOT_CONVERGED, "面・エッジ・頂点のいずれでも流れ則と整合する戻しが得られません。"
        Exit Function
      End If
      P2BuildFaceGradient3D principalAngle, maximumMode, minimumMode, sinPhi, nx, ny, nz, nt
      P2BuildFaceGradient3D principalAngle, maximumMode, minimumMode, sinPsi, mx, my, mz, mt
      initialMaximumMode = maximumMode
      initialMinimumMode = minimumMode
      qx = d00 * mx + d01 * my + d01 * mz
      qy = d01 * mx + d00 * my + d01 * mz
      qz = d01 * mx + d01 * my + d00 * mz
      qt = d22 * mt
      rx = nx * d00 + ny * d01 + nz * d01
      ry = nx * d01 + ny * d00 + nz * d01
      rt = nt * d22
      denominator = nx * qx + ny * qy + nz * qz + nt * qt
      If Abs(denominator) <= 0.00000000000001 * yieldScale Then
        P2SetFailure outputState, P2_FAILURE_DENOMINATOR, "面切替後の3D塑性補正の分母が0です。"
        Exit Function
      End If
      yieldTrial = P2FaceYieldAtStress(maximumMode, minimumMode, trialSx, trialSy, trialSz, trialTxy, trialPrincipalAngle, sinPhi, inputState.cohesion, cosPhi)
      If Not P2IsFinite(yieldTrial) Then
        P2SetFailure outputState, P2_FAILURE_NONFINITE, "面切替後のtrial降伏関数が有限ではありません。"
        Exit Function
      End If
      lambdaValue = yieldTrial / denominator
      If lambdaValue < 0# Then lambdaValue = 0#
      GoTo NextFaceIteration
    End If
    yieldScale = Abs(sigmaMax) + Abs(sigmaMin) + 2# * inputState.cohesion * cosPhi
    If yieldScale < 1# Then yieldScale = 1#
    If Abs(yieldCurrent) <= inputState.tolerance * yieldScale Then Exit For
    deltaLambda = yieldCurrent / denominator
    If Abs(deltaLambda) <= lambdaTolerance Then Exit For
    lambdaValue = lambdaValue + deltaLambda
    If lambdaValue < -lambdaTolerance Then
      P2SetFailure outputState, P2_FAILURE_DENOMINATOR, "3D塑性乗数の更新で負の値になりました。"
      Exit Function
    End If
    If lambdaValue < 0# Then lambdaValue = 0#
  Next iter
  If iter > maxIterations Then
    If P2TryEdgeReturn(inputState, trialSx, trialSy, trialSz, trialTxy, d00, d01, d22, sinPhi, cosPhi, sinPsi, trialPrincipalAngle, trialSigmaMax, trialSigmaMid, trialSigmaMin, initialMaximumMode, initialMinimumMode, maximumMode, minimumMode, outputState) Then
      outputState.iterations = maxIterations
      P2MaterialPointUpdateCore = True
      Exit Function
    End If
    If P2TryVertexReturn(inputState, trialSx, trialSy, trialSz, trialTxy, d00, d01, d22, sinPhi, cosPhi, sinPsi, outputState) Then
      outputState.iterations = maxIterations
      P2MaterialPointUpdateCore = True
      Exit Function
    End If
    P2SetFailure outputState, P2_FAILURE_NOT_CONVERGED, "3D材料点の塑性補正が反復上限に達しました。"
    outputState.iterations = maxIterations
    Exit Function
  End If
  outputState.Stress(0) = sx
  outputState.Stress(1) = sy
  outputState.Stress(2) = txy
  outputState.Stress(3) = sz
  outputState.PlasticMultiplier = lambdaValue
  outputState.PlasticStrain(0) = lambdaValue * mx
  outputState.PlasticStrain(1) = lambdaValue * my
  outputState.PlasticStrain(2) = lambdaValue * mt
  outputState.PlasticStrain(3) = lambdaValue * mz
  outputState.Tangent(0, 0) = d00 - qx * rx / denominator
  outputState.Tangent(0, 1) = d01 - qx * ry / denominator
  outputState.Tangent(0, 2) = -qx * rt / denominator
  outputState.Tangent(1, 0) = d01 - qy * rx / denominator
  outputState.Tangent(1, 1) = d00 - qy * ry / denominator
  outputState.Tangent(1, 2) = -qy * rt / denominator
  outputState.Tangent(2, 0) = -qt * rx / denominator
  outputState.Tangent(2, 1) = -qt * ry / denominator
  outputState.Tangent(2, 2) = d22 - qt * rt / denominator
  outputState.PrincipalStress(0) = sigmaMax
  outputState.PrincipalStress(1) = sigmaMid
  outputState.PrincipalStress(2) = sigmaMin
  outputState.principalAngle = principalAngle * 180# / piValue
  outputState.YieldFunction = yieldCurrent
  outputState.yielded = True
  outputState.Elastic = False
  outputState.converged = True
  outputState.iterations = iter
  For iter = 0 To 3
    If Not P2IsFinite(outputState.Stress(iter)) Then
      P2SetFailure outputState, P2_FAILURE_NONFINITE, "平面ひずみ更新応力に有限でない値があります。"
      Exit Function
    End If
    If Not P2IsFinite(outputState.PlasticStrain(iter)) Then
      P2SetFailure outputState, P2_FAILURE_NONFINITE, "平面ひずみ塑性ひずみに有限でない値があります。"
      Exit Function
    End If
  Next iter
  P2MaterialPointUpdateCore = True
  Exit Function
Failed:
  P2SetFailure outputState, P2_FAILURE_NONFINITE, "平面ひずみ材料点更新で算術エラーが発生しました。"
End Function

Private Function P2RunEqualSubsteps(ByRef inputState As P2_MaterialPointInput, ByRef outputState As P2_MaterialPointOutput, ByVal stepCount As Long) As Boolean
  Dim stepInput As P2_MaterialPointInput, stepOutput As P2_MaterialPointOutput
  Dim currentSx As Double, currentSy As Double, currentSz As Double, currentTxy As Double
  Dim stepId As Long, steps As Long
  P2RunEqualSubsteps = False
  steps = stepCount
  If steps < 1 Then steps = 1
  currentSx = inputState.PreviousStress(0)
  currentSy = inputState.PreviousStress(1)
  currentTxy = inputState.PreviousStress(2)
  currentSz = inputState.PreviousStress(3)
  For stepId = 1 To steps
    stepInput = inputState
    stepInput.maxSubsteps = 1
    stepInput.EnableSubstepping = False
    stepInput.PreviousStress(0) = currentSx
    stepInput.PreviousStress(1) = currentSy
    stepInput.PreviousStress(2) = currentTxy
    stepInput.PreviousStress(3) = currentSz
    stepInput.StrainIncrement(0) = inputState.StrainIncrement(0) / steps
    stepInput.StrainIncrement(1) = inputState.StrainIncrement(1) / steps
    stepInput.StrainIncrement(2) = inputState.StrainIncrement(2) / steps
    P2ResetOutput stepOutput
    If Not P2MaterialPointUpdateCore(stepInput, stepOutput) Then
      P2CopyOutput stepOutput, outputState
      Exit Function
    End If
    currentSx = stepOutput.Stress(0)
    currentSy = stepOutput.Stress(1)
    currentTxy = stepOutput.Stress(2)
    currentSz = stepOutput.Stress(3)
  Next stepId
  P2CopyOutput stepOutput, outputState
  outputState.Stress(0) = currentSx
  outputState.Stress(1) = currentSy
  outputState.Stress(2) = currentTxy
  outputState.Stress(3) = currentSz
  outputState.converged = True
  P2RunEqualSubsteps = True
End Function

Private Function P2FillFiniteDifferenceTangent(ByRef inputState As P2_MaterialPointInput, ByRef outputState As P2_MaterialPointOutput, ByVal stepCount As Long) As Boolean
  Dim pertInput As P2_MaterialPointInput, pertOutput As P2_MaterialPointOutput
  Dim col As Long, i As Long, j As Long
  Dim h As Double, strainScale As Double
  Dim baseSx As Double, baseSy As Double, baseTxy As Double
  Dim savedTangent(0 To 2, 0 To 2) As Double
  P2FillFiniteDifferenceTangent = False
  For i = 0 To 2
    For j = 0 To 2
      savedTangent(i, j) = outputState.Tangent(i, j)
    Next j
  Next i
  baseSx = outputState.Stress(0)
  baseSy = outputState.Stress(1)
  baseTxy = outputState.Stress(2)
  strainScale = Abs(inputState.StrainIncrement(0)) + Abs(inputState.StrainIncrement(1)) + Abs(inputState.StrainIncrement(2))
  h = 0.000001 * (1# + strainScale)
  If h < 0.0000000001 Then h = 0.0000000001
  For col = 0 To 2
    pertInput = inputState
    pertInput.StrainIncrement(col) = pertInput.StrainIncrement(col) + h
    If Not P2RunEqualSubsteps(pertInput, pertOutput, stepCount) Then
      For i = 0 To 2
        For j = 0 To 2
          outputState.Tangent(i, j) = savedTangent(i, j)
        Next j
      Next i
      Exit Function
    End If
    outputState.Tangent(0, col) = (pertOutput.Stress(0) - baseSx) / h
    outputState.Tangent(1, col) = (pertOutput.Stress(1) - baseSy) / h
    outputState.Tangent(2, col) = (pertOutput.Stress(2) - baseTxy) / h
  Next col
  P2FillFiniteDifferenceTangent = True
End Function

Public Function P2MaterialPointUpdate(ByRef inputState As P2_MaterialPointInput, ByRef outputState As P2_MaterialPointOutput)
  Dim stepInput As P2_MaterialPointInput, stepOutput As P2_MaterialPointOutput, failureOutput As P2_MaterialPointOutput
  Dim currentSx As Double, currentSy As Double, currentSz As Double, currentTxy As Double
  Dim totalPlastic(0 To 3) As Double, totalLambda As Double
  Dim stepId As Long, stepCount As Long, maxSubsteps As Long
  Dim totalIterations As Long, yieldedAny As Boolean, allOK As Boolean
  Dim specialFailureCode As Long, specialFailureMessage As String
  Dim d00 As Double, d01 As Double, d22 As Double, lameLambda As Double, shearModulus As Double
  Dim yieldValue As Double, sigmaMax As Double, sigmaMid As Double, sigmaMin As Double, angle As Double, radius As Double
  Dim sinPhi As Double, cosPhi As Double, yieldScale As Double
  Dim maximumMode As Long, minimumMode As Long, piValue As Double
  P2ResetOutput outputState
  AccelMaterialCalls = AccelMaterialCalls + 1
  P2MaterialPointUpdate = False
  If Not P2ValidateInput(inputState, outputState) Then Exit Function
  maxSubsteps = inputState.maxSubsteps
  If maxSubsteps <= 0 Then maxSubsteps = 1
  If maxSubsteps > P2_DEFAULT_MAX_SUBSTEPS Then maxSubsteps = P2_DEFAULT_MAX_SUBSTEPS
  If Not inputState.EnableSubstepping Then maxSubsteps = 1
  piValue = 3.14159265358979
  If inputState.UseCachedConstants Then
    d00 = inputState.CachedElasticD00
    d01 = inputState.CachedElasticD01
    d22 = inputState.CachedElasticD22
    sinPhi = inputState.CachedSinFriction
    cosPhi = inputState.CachedCosFriction
  Else
    shearModulus = inputState.Young / (2# * (1# + inputState.Poisson))
    lameLambda = inputState.Young * inputState.Poisson / ((1# + inputState.Poisson) * (1# - 2# * inputState.Poisson))
    d00 = lameLambda + 2# * shearModulus
    d01 = lameLambda
    d22 = shearModulus
    sinPhi = Sin(inputState.frictionAngle * piValue / 180#)
    cosPhi = Cos(inputState.frictionAngle * piValue / 180#)
  End If
  outputState.TrialStress(0) = inputState.PreviousStress(0) + d00 * inputState.StrainIncrement(0) + d01 * inputState.StrainIncrement(1)
  outputState.TrialStress(1) = inputState.PreviousStress(1) + d01 * inputState.StrainIncrement(0) + d00 * inputState.StrainIncrement(1)
  outputState.TrialStress(2) = inputState.PreviousStress(2) + d22 * inputState.StrainIncrement(2)
  outputState.TrialStress(3) = inputState.PreviousStress(3) + d01 * (inputState.StrainIncrement(0) + inputState.StrainIncrement(1))
  If Not P2EvaluateYield(outputState.TrialStress(0), outputState.TrialStress(1), outputState.TrialStress(3), outputState.TrialStress(2), sinPhi, inputState.cohesion, cosPhi, yieldValue, sigmaMax, sigmaMid, sigmaMin, angle, radius, maximumMode, minimumMode) Then
    P2SetFailure outputState, P2_FAILURE_NONFINITE, "全増分の平面ひずみ弾性試行応力が有限ではありません。"
    Exit Function
  End If
  outputState.TrialYieldFunction = yieldValue
  yieldScale = Abs(sigmaMax) + Abs(sigmaMin) + 2# * inputState.cohesion * cosPhi
  If yieldScale < 1# Then yieldScale = 1#
  If yieldValue <= inputState.tolerance * yieldScale Then
    outputState.Stress(0) = outputState.TrialStress(0)
    outputState.Stress(1) = outputState.TrialStress(1)
    outputState.Stress(2) = outputState.TrialStress(2)
    outputState.Stress(3) = outputState.TrialStress(3)
    outputState.Tangent(0, 0) = d00: outputState.Tangent(0, 1) = d01: outputState.Tangent(0, 2) = 0#
    outputState.Tangent(1, 0) = d01: outputState.Tangent(1, 1) = d00: outputState.Tangent(1, 2) = 0#
    outputState.Tangent(2, 0) = 0#: outputState.Tangent(2, 1) = 0#: outputState.Tangent(2, 2) = d22
    outputState.PrincipalStress(0) = sigmaMax
    outputState.PrincipalStress(1) = sigmaMid
    outputState.PrincipalStress(2) = sigmaMin
    outputState.principalAngle = angle * 180# / piValue
    outputState.YieldFunction = yieldValue
    outputState.yielded = False
    outputState.PlasticOccurred = False
    outputState.Elastic = True
    outputState.converged = True
    outputState.iterations = 1
    outputState.UsedSubsteps = 1
    AccelNoteSubsteps 1
    P3ElasticFastCount = P3ElasticFastCount + 1
    P2MaterialPointUpdate = True
    Exit Function
  End If
  stepCount = 1
  If inputState.PreferredSubsteps > 1 Then
    stepCount = inputState.PreferredSubsteps
    If stepCount > maxSubsteps Then stepCount = maxSubsteps
  End If
  Do
    currentSx = inputState.PreviousStress(0)
    currentSy = inputState.PreviousStress(1)
    currentTxy = inputState.PreviousStress(2)
    currentSz = inputState.PreviousStress(3)
    totalPlastic(0) = 0#: totalPlastic(1) = 0#: totalPlastic(2) = 0#: totalPlastic(3) = 0#
    totalLambda = 0#
    totalIterations = 0
    yieldedAny = False
    specialFailureCode = 0
    specialFailureMessage = vbNullString
    allOK = True
    For stepId = 1 To stepCount
      stepInput.Young = inputState.Young
      stepInput.Poisson = inputState.Poisson
      stepInput.frictionAngle = inputState.frictionAngle
      stepInput.cohesion = inputState.cohesion
      stepInput.dilationAngle = inputState.dilationAngle
      stepInput.tolerance = inputState.tolerance
      stepInput.maxIterations = inputState.maxIterations
      stepInput.maxSubsteps = 1
      stepInput.EnableSubstepping = False
      stepInput.UseCachedConstants = inputState.UseCachedConstants
      stepInput.CachedElasticD00 = inputState.CachedElasticD00
      stepInput.CachedElasticD01 = inputState.CachedElasticD01
      stepInput.CachedElasticD22 = inputState.CachedElasticD22
      stepInput.CachedSinFriction = inputState.CachedSinFriction
      stepInput.CachedCosFriction = inputState.CachedCosFriction
      stepInput.CachedSinDilation = inputState.CachedSinDilation
      stepInput.CachedCosDilation = inputState.CachedCosDilation
      stepInput.PreviousStress(0) = currentSx
      stepInput.PreviousStress(1) = currentSy
      stepInput.PreviousStress(2) = currentTxy
      stepInput.PreviousStress(3) = currentSz
      stepInput.StrainIncrement(0) = inputState.StrainIncrement(0) / stepCount
      stepInput.StrainIncrement(1) = inputState.StrainIncrement(1) / stepCount
      stepInput.StrainIncrement(2) = inputState.StrainIncrement(2) / stepCount
      P2ResetOutput stepOutput
      If Not P2MaterialPointUpdateCore(stepInput, stepOutput) Then
        P2CopyOutput stepOutput, failureOutput
        allOK = False
        Exit For
      End If
      currentSx = stepOutput.Stress(0)
      currentSy = stepOutput.Stress(1)
      currentTxy = stepOutput.Stress(2)
      currentSz = stepOutput.Stress(3)
      totalPlastic(0) = totalPlastic(0) + stepOutput.PlasticStrain(0)
      totalPlastic(1) = totalPlastic(1) + stepOutput.PlasticStrain(1)
      totalPlastic(2) = totalPlastic(2) + stepOutput.PlasticStrain(2)
      totalPlastic(3) = totalPlastic(3) + stepOutput.PlasticStrain(3)
      totalLambda = totalLambda + stepOutput.PlasticMultiplier
      totalIterations = totalIterations + stepOutput.iterations
      If stepOutput.yielded Then yieldedAny = True
      If stepOutput.failureCode <> 0 Or Len(stepOutput.failureMessage) > 0 Then
        specialFailureCode = stepOutput.failureCode
        specialFailureMessage = stepOutput.failureMessage
      End If
    Next stepId
    If allOK Then Exit Do
    If stepCount >= maxSubsteps Then
      P2CopyOutput failureOutput, outputState
      outputState.converged = False
      P2MaterialPointUpdate = False
      Exit Function
    End If
    stepCount = stepCount * 2
    If stepCount > maxSubsteps Then stepCount = maxSubsteps
  Loop
  outputState.Stress(0) = currentSx
  outputState.Stress(1) = currentSy
  outputState.Stress(2) = currentTxy
  outputState.Stress(3) = currentSz
  outputState.PlasticStrain(0) = totalPlastic(0)
  outputState.PlasticStrain(1) = totalPlastic(1)
  outputState.PlasticStrain(2) = totalPlastic(2)
  outputState.PlasticStrain(3) = totalPlastic(3)
  outputState.PlasticMultiplier = totalLambda
  outputState.yielded = stepOutput.yielded
  outputState.PlasticOccurred = yieldedAny
  outputState.Elastic = Not stepOutput.yielded
  outputState.converged = True
  outputState.failureCode = specialFailureCode
  outputState.failureMessage = specialFailureMessage
  outputState.iterations = totalIterations
  outputState.UsedSubsteps = stepCount
  AccelNoteSubsteps stepCount
  outputState.Tangent(0, 0) = stepOutput.Tangent(0, 0)
  outputState.Tangent(0, 1) = stepOutput.Tangent(0, 1)
  outputState.Tangent(0, 2) = stepOutput.Tangent(0, 2)
  outputState.Tangent(1, 0) = stepOutput.Tangent(1, 0)
  outputState.Tangent(1, 1) = stepOutput.Tangent(1, 1)
  outputState.Tangent(1, 2) = stepOutput.Tangent(1, 2)
  outputState.Tangent(2, 0) = stepOutput.Tangent(2, 0)
  outputState.Tangent(2, 1) = stepOutput.Tangent(2, 1)
  outputState.Tangent(2, 2) = stepOutput.Tangent(2, 2)
  outputState.PrincipalStress(0) = stepOutput.PrincipalStress(0)
  outputState.PrincipalStress(1) = stepOutput.PrincipalStress(1)
  outputState.PrincipalStress(2) = stepOutput.PrincipalStress(2)
  outputState.principalAngle = stepOutput.principalAngle
  outputState.YieldFunction = stepOutput.YieldFunction
  ' サブステップ後は最後のalgorithmic tangentを使う。FDは列途中失敗で混成接線になるため既定では呼ばない。
  If Not yieldedAny Then
    outputState.Tangent(0, 0) = d00: outputState.Tangent(0, 1) = d01: outputState.Tangent(0, 2) = 0#
    outputState.Tangent(1, 0) = d01: outputState.Tangent(1, 1) = d00: outputState.Tangent(1, 2) = 0#
    outputState.Tangent(2, 0) = 0#: outputState.Tangent(2, 1) = 0#: outputState.Tangent(2, 2) = d22
    outputState.PrincipalStress(0) = sigmaMax
    outputState.PrincipalStress(1) = sigmaMid
    outputState.PrincipalStress(2) = sigmaMin
    outputState.principalAngle = angle * 180# / piValue
    outputState.YieldFunction = yieldValue
  End If
  P2MaterialPointUpdate = True
End Function

Public Function P2IsPlasticState(ByRef outputState As P2_MaterialPointOutput) As Boolean
  P2IsPlasticState = outputState.yielded
End Function

Private Sub P2SetSelfTestInput(ByRef inputState As P2_MaterialPointInput, ByVal ex As Double, ByVal ey As Double, ByVal gammaXY As Double, ByVal phi As Double, ByVal cohesionValue As Double, ByVal psi As Double)
  Dim i As Long
  inputState.Young = 1000#
  inputState.Poisson = 0.3
  inputState.frictionAngle = phi
  inputState.cohesion = cohesionValue
  inputState.dilationAngle = psi
  inputState.tolerance = 0.00000001
  inputState.maxIterations = P2_DEFAULT_MAX_ITERATIONS
  inputState.maxSubsteps = P2_DEFAULT_MAX_SUBSTEPS
  inputState.EnableSubstepping = True
  For i = 0 To 3
    inputState.PreviousStress(i) = 0#
  Next i
  inputState.StrainIncrement(0) = ex
  inputState.StrainIncrement(1) = ey
  inputState.StrainIncrement(2) = gammaXY
End Sub

Private Function P2RunSelfTestCase(ByRef ws As Worksheet, ByVal rowNumber As Long, ByVal startColumn As Long, ByVal caseName As String, ByRef inputState As P2_MaterialPointInput, ByVal expectedYielded As Boolean, ByVal expectedFailureCode As Long, ByRef outputState As P2_MaterialPointOutput) As Boolean
  Dim passed As Boolean, yieldScale As Double
  passed = P2MaterialPointUpdate(inputState, outputState)
  If passed Then
    If Not outputState.converged Then passed = False
  End If
  If passed Then
    If outputState.yielded <> expectedYielded Then passed = False
  End If
  If passed Then
    If expectedFailureCode <> 0 Then
      If outputState.failureCode <> expectedFailureCode Then passed = False
    End If
  End If
  If passed Then
    yieldScale = Abs(outputState.PrincipalStress(0)) + Abs(outputState.PrincipalStress(2)) + 2# * inputState.cohesion
    If yieldScale < 1# Then yieldScale = 1#
    If outputState.yielded Then
      If Abs(outputState.YieldFunction) > 0.000001 * yieldScale Then passed = False
    End If
  End If
  ws.Cells(rowNumber, startColumn).value2 = caseName
  ws.Cells(rowNumber, startColumn + 1).value2 = outputState.converged
  ws.Cells(rowNumber, startColumn + 2).value2 = outputState.yielded
  ws.Cells(rowNumber, startColumn + 3).value2 = outputState.Elastic
  ws.Cells(rowNumber, startColumn + 4).value2 = outputState.TrialYieldFunction
  ws.Cells(rowNumber, startColumn + 5).value2 = outputState.YieldFunction
  ws.Cells(rowNumber, startColumn + 6).value2 = outputState.PlasticMultiplier
  ws.Cells(rowNumber, startColumn + 7).value2 = outputState.failureCode
  ws.Cells(rowNumber, startColumn + 9).value2 = outputState.Stress(3)
  ws.Cells(rowNumber, startColumn + 10).value2 = outputState.PlasticStrain(3)
  If passed Then
    ws.Cells(rowNumber, startColumn + 8).value2 = "PASS"
  Else
    ws.Cells(rowNumber, startColumn + 8).value2 = "FAIL"
  End If
  P2RunSelfTestCase = passed
End Function

Public Function P2RunMaterialPointSelfTest() As Boolean
  Dim ws As Worksheet, inputState As P2_MaterialPointInput, outputState As P2_MaterialPointOutput
  Dim plasticOutput As P2_MaterialPointOutput, rotationOutput As P2_MaterialPointOutput
  Dim rowNumber As Long, allPassed As Boolean, passed As Boolean, testStartColumn As Long
  Dim principalDifference As Double, tangentDifference As Double
  On Error GoTo Failed
  Set ws = GetIntegratedResultSheet()
  testStartColumn = FEM_INTEGRATED_P2_TEST_START_COLUMN
  ws.range(ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW, testStartColumn), ws.Cells(FEM_INTEGRATED_P2_TEST_END_ROW, testStartColumn + 11)).ClearContents
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW, testStartColumn).value2 = "P2 Material Point Self Test"
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW + 1, testStartColumn).value2 = "Case"
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW + 1, testStartColumn + 1).value2 = "Converged"
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW + 1, testStartColumn + 2).value2 = "Yielded"
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW + 1, testStartColumn + 3).value2 = "Elastic"
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW + 1, testStartColumn + 4).value2 = "TrialF"
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW + 1, testStartColumn + 5).value2 = "FinalF"
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW + 1, testStartColumn + 6).value2 = "PlasticMultiplier"
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW + 1, testStartColumn + 7).value2 = "FailureCode"
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW + 1, testStartColumn + 8).value2 = "Status"
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW + 1, testStartColumn + 9).value2 = "SigmaZ"
  ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW + 1, testStartColumn + 10).value2 = "PlasticStrainZ"
  rowNumber = FEM_INTEGRATED_P2_TEST_START_ROW + 2
  allPassed = True

  P2SetSelfTestInput inputState, 0.0001, 0#, 0#, 30#, 100#, 0#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "elastic", inputState, False, 0, outputState): If Not passed Then allPassed = False
  rowNumber = rowNumber + 1
  P2SetSelfTestInput inputState, 0#, 0#, 0#, 30#, 1#, 0#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "zero_stress", inputState, False, 0, outputState): If Not passed Then allPassed = False
  rowNumber = rowNumber + 1
  P2SetSelfTestInput inputState, 0.005, 0.005, 0#, 30#, 1#, 0#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "equal_principal_stress", inputState, False, 0, outputState): If Not passed Then allPassed = False
  rowNumber = rowNumber + 1
  P2SetSelfTestInput inputState, 0.0001, 0.0002, 0#, 30#, 100#, 0#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "plane_strain_sigma_z", inputState, False, 0, outputState)
  If Abs(outputState.Stress(3) - inputState.Young * inputState.Poisson / ((1# + inputState.Poisson) * (1# - 2# * inputState.Poisson)) * 0.0003) > 0.00000001 Then passed = False
  If Abs(outputState.PlasticStrain(3)) > 0.000000000001 Then passed = False
  If Not passed Then allPassed = False
  rowNumber = rowNumber + 1
  P2SetSelfTestInput inputState, -0.1, 0#, 0#, 30#, 1#, 0#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "plastic_loading_psi0", inputState, True, 0, plasticOutput): If Not passed Then allPassed = False
  rowNumber = rowNumber + 1
  inputState.PreviousStress(0) = plasticOutput.Stress(0)
  inputState.PreviousStress(1) = plasticOutput.Stress(1)
  inputState.PreviousStress(2) = plasticOutput.Stress(2)
  inputState.PreviousStress(3) = plasticOutput.Stress(3)
  inputState.StrainIncrement(0) = 0.0005
  inputState.StrainIncrement(1) = 0.0005
  inputState.StrainIncrement(2) = 0#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "unloading", inputState, False, 0, outputState): If Not passed Then allPassed = False
  rowNumber = rowNumber + 1
  P2SetSelfTestInput inputState, -0.05, 0#, 0#, 30#, 0#, 0#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "zero_cohesion", inputState, True, 0, outputState): If Not passed Then allPassed = False
  rowNumber = rowNumber + 1
  P2SetSelfTestInput inputState, 0#, 0#, 0.02, 0#, 5#, 0#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "zero_friction", inputState, True, 0, outputState): If Not passed Then allPassed = False
  rowNumber = rowNumber + 1
  P2SetSelfTestInput inputState, -0.002, 0#, 0#, 30#, 1#, 30#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "associated_flow_psi_phi", inputState, True, 0, outputState): If Not passed Then allPassed = False
  rowNumber = rowNumber + 1
  P2SetSelfTestInput inputState, -0.002, 0#, 0#, 30#, 1#, 0#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "nonassociated_flow_psi0", inputState, True, 0, outputState): If Not passed Then allPassed = False
  rowNumber = rowNumber + 1
  P2SetSelfTestInput inputState, -0.002, 0#, 0#, 30#, 1#, 0#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "rotation_reference_x", inputState, True, 0, outputState): If Not passed Then allPassed = False
  P2SetSelfTestInput inputState, 0#, -0.002, 0#, 30#, 1#, 0#
  passed = P2RunSelfTestCase(ws, rowNumber + 1, testStartColumn, "rotation_reference_y", inputState, True, 0, rotationOutput): If Not passed Then allPassed = False
  principalDifference = Abs(outputState.PrincipalStress(0) - rotationOutput.PrincipalStress(0)) + Abs(outputState.PrincipalStress(1) - rotationOutput.PrincipalStress(1))
  If principalDifference > 0.000001 Then allPassed = False
  rowNumber = rowNumber + 2
  P2SetSelfTestInput inputState, 0#, 0#, 0#, 30#, 0#, 30#
  inputState.PreviousStress(0) = -10#: inputState.PreviousStress(1) = -10#: inputState.PreviousStress(2) = 0#: inputState.PreviousStress(3) = -10#
  passed = P2RunSelfTestCase(ws, rowNumber, testStartColumn, "mohr_coulomb_vertex_3d", inputState, True, P2_FAILURE_VERTEX, outputState): If Not passed Then allPassed = False
  rowNumber = rowNumber + 1
  ws.Cells(rowNumber, testStartColumn).value2 = "rotation_principal_difference"
  ws.Cells(rowNumber, testStartColumn + 1).value2 = principalDifference
  rowNumber = rowNumber + 1
  tangentDifference = Abs(plasticOutput.Tangent(0, 1) - plasticOutput.Tangent(1, 0))
  ws.Cells(rowNumber, testStartColumn).value2 = "nonassociated_tangent_asymmetry"
  ws.Cells(rowNumber, testStartColumn + 1).value2 = tangentDifference
  ws.Cells(rowNumber + 1, testStartColumn).value2 = "Overall"
  If allPassed Then
    ws.Cells(rowNumber + 1, testStartColumn + 1).value2 = "PASS"
  Else
    ws.Cells(rowNumber + 1, testStartColumn + 1).value2 = "FAIL"
  End If
  ws.range(ws.columns(testStartColumn), ws.columns(testStartColumn + 10)).EntireColumn.AutoFit
  P2RunMaterialPointSelfTest = allPassed
  Exit Function
Failed:
  If Not ws Is Nothing Then
    ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW, testStartColumn + 9).value2 = "ERROR"
    ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW, testStartColumn + 10).value2 = Err.Number
    ws.Cells(FEM_INTEGRATED_P2_TEST_START_ROW, testStartColumn + 11).value2 = Err.Description
  End If
  P2RunMaterialPointSelfTest = False
End Function

Public Function P2RunEdgeReturnSelfTest() As Boolean
  Dim inputState As P2_MaterialPointInput, outputState As P2_MaterialPointOutput
  Dim exValues(0 To 6) As Double, eyValues(0 To 6) As Double, gammaValues(0 To 4) As Double
  Dim i As Long, j As Long, k As Long
  On Error GoTo Failed
  exValues(0) = -0.02: exValues(1) = -0.01: exValues(2) = -0.005: exValues(3) = -0.002: exValues(4) = 0.002: exValues(5) = 0.005: exValues(6) = 0.01
  eyValues(0) = -0.02: eyValues(1) = -0.01: eyValues(2) = -0.005: eyValues(3) = -0.002: eyValues(4) = 0.002: eyValues(5) = 0.005: eyValues(6) = 0.01
  gammaValues(0) = -0.01: gammaValues(1) = -0.002: gammaValues(2) = 0#: gammaValues(3) = 0.002: gammaValues(4) = 0.01
  P2SetSelfTestInput inputState, -0.1, 0.098, -0.1, 30#, 1#, 0#
  If P2MaterialPointUpdate(inputState, outputState) Then
    If outputState.yielded And outputState.converged And outputState.failureCode <> P2_FAILURE_VERTEX And Abs(outputState.PrincipalStress(0) - outputState.PrincipalStress(1)) <= 0.000001 And Abs(outputState.PrincipalStress(1) - outputState.PrincipalStress(2)) > 0.0001 Then
      P2RunEdgeReturnSelfTest = True
      Exit Function
    End If
  End If
  P2SetSelfTestInput inputState, -0.1, 0.098, 0#, 30#, 1#, 0#
  If P2MaterialPointUpdate(inputState, outputState) Then
    If outputState.yielded And outputState.converged And outputState.failureCode <> P2_FAILURE_VERTEX And Abs(outputState.PrincipalStress(0) - outputState.PrincipalStress(1)) <= 0.000001 And Abs(outputState.PrincipalStress(1) - outputState.PrincipalStress(2)) > 0.0001 Then
      P2RunEdgeReturnSelfTest = True
      Exit Function
    End If
  End If
  For i = 0 To 6
    For j = 0 To 6
      For k = 0 To 4
        P2SetSelfTestInput inputState, exValues(i), eyValues(j), gammaValues(k), 30#, 1#, 0#
        If P2MaterialPointUpdate(inputState, outputState) Then
          If outputState.yielded And outputState.converged And outputState.failureCode <> P2_FAILURE_VERTEX And Abs(outputState.PrincipalStress(0) - outputState.PrincipalStress(1)) <= 0.000001 And Abs(outputState.PrincipalStress(1) - outputState.PrincipalStress(2)) > 0.0001 Then
            P2RunEdgeReturnSelfTest = True
            Exit Function
          End If
        End If
      Next k
    Next j
  Next i
  P2RunEdgeReturnSelfTest = False
  Exit Function
Failed:
  P2RunEdgeReturnSelfTest = False
End Function

Public Function P2RunEdgeReturnCaseReport() As String
  Dim inputState As P2_MaterialPointInput, outputState As P2_MaterialPointOutput
  On Error GoTo Failed
  P2SetSelfTestInput inputState, -0.1, 0.098, -0.1, 30#, 1#, 0#
  If P2MaterialPointUpdate(inputState, outputState) Then
    P2RunEdgeReturnCaseReport = "ok=True|message=" & outputState.failureMessage & "|failure=" & CStr(outputState.failureCode) & "|yield=" & CStr(outputState.YieldFunction) & "|s0=" & CStr(outputState.PrincipalStress(0)) & "|s1=" & CStr(outputState.PrincipalStress(1)) & "|s2=" & CStr(outputState.PrincipalStress(2))
  Else
    P2RunEdgeReturnCaseReport = "ok=False|message=" & outputState.failureMessage & "|failure=" & CStr(outputState.failureCode) & "|yield=" & CStr(outputState.YieldFunction) & "|s0=" & CStr(outputState.PrincipalStress(0)) & "|s1=" & CStr(outputState.PrincipalStress(1)) & "|s2=" & CStr(outputState.PrincipalStress(2))
  End If
  Exit Function
Failed:
  P2RunEdgeReturnCaseReport = "error=" & CStr(Err.Number) & "|" & Err.Description
End Function



Private Sub P3ApplyIncrementPredictor(ByVal deltaLambda As Double, ByVal prescribedFactor As Double, ByRef used As Boolean)
  Dim i As Long, ratio As Double
  used = False
  If Not AccelV1 Or AccelBaselineRetry Or Not AccelHistoryReady Then Exit Sub
  If Not AdaptiveAllowMethod(1) Then Exit Sub
  If P3AfterCutback Or AccelPreviousStep <= 0# Or P6MixedUP Or P3HasJointElements Then Exit Sub
  ratio = deltaLambda / AccelPreviousStep
  If ratio <= 0# Or ratio > AccelPredictMaxRatio Then Exit Sub
  ratio = ratio * AccelPredictBeta
  If ratio <= 0# Then Exit Sub
  For i = 0 To lastDof
    If NodeCond(i) = 0 Then
      TDisp(i) = P3CommittedDisp(i) + ratio * AccelDelta(i)
    Else
      TDisp(i) = P3TargetBoundaryDisp(i, prescribedFactor)
    End If
    UDisp(i) = TDisp(i)
  Next i
  P3TrialStateValid = False
  used = True
  AccelPredictCount = AccelPredictCount + 1
  AdaptiveUseMethod 1
End Sub

Private Sub P3AccelResetAA()
  mAAReady = False
End Sub

Private Function P3AccelTryAA(ByRef incrementBase() As Double, ByRef targetForce() As Double, ByRef internalForce() As Double, ByRef correctionRatio As Double, ByVal residualBefore As Double) As Boolean
  Dim x0() As Double, f0() As Double, candidate() As Double
  Dim i As Long, num As Double, den As Double, corrScale As Double, d As Double, gamma As Double
  Dim plainRes As Double, candidateRes As Double, valid As Boolean, adaptiveStart As Double, adaptiveRejectReason As String
  P3AccelTryAA = False
  If Not AccelV4 Or AccelBaselineRetry Or P6MixedUP Or P3HasJointElements Then
    mAAReady = False: Exit Function
  End If
  If P3TangentAge = 0 Or P3TangentJustRebuilt Or P3ForceTangentRebuild Or P3Ls50Pending Or P3Ls50Watch Or P3AfterCutback Then
    mAAReady = False: Exit Function
  End If
  If Not AdaptiveAllowAA(RelativeResidualFree, P3_ENGINEERING_RESIDUAL) Then mAAReady = False: Exit Function
  If Not AccelSmallTangentChange() Then mAAReady = False: Exit Function
  x0 = TDisp: f0 = Disp
  If mAAReady Then
    If mAAFactor <> P6FactorizationCount Then mAAReady = False
  End If
  If mAAReady Then
    For i = 0 To lastDof
      If NodeCond(i) = 0 Then
        d = f0(i) - mAAF(i)
        num = num + d * f0(i): den = den + d * d: corrScale = corrScale + f0(i) * f0(i)
      End If
    Next i
    If den <= 1E-24 * (corrScale + 1E-30) Then mAAReady = False
  End If
  If Not mAAReady Then
    ' Seed only from the existing, undamped Newton/line-search path.
    mAAX = x0: mAAF = f0: mAAFactor = P6FactorizationCount
    mAAReady = True
    Exit Function
  End If
  If AdaptiveEnabled Then adaptiveStart = Timer
  AdaptiveUseMethod 4
  plainRes = 1E+30: candidateRes = 1E+30
  gamma = num / den
  If Not P2IsFinite(gamma) Or Abs(gamma) > 1# Then adaptiveRejectReason = "MIXING_LIMIT": GoTo RejectAA
  ReDim candidate(lastDof)
  For i = 0 To lastDof
    candidate(i) = x0(i) + f0(i)
    If NodeCond(i) = 0 Then candidate(i) = candidate(i) - gamma * ((x0(i) - mAAX(i)) + (f0(i) - mAAF(i)))
    If Not P2IsFinite(candidate(i)) Then adaptiveRejectReason = "NONFINITE": GoTo RejectAA
    TDisp(i) = x0(i) + f0(i): UDisp(i) = TDisp(i)
  Next i
  P3TrialStateValid = False
  If Not P3EvaluateTrialState(incrementBase, internalForce) Then adaptiveRejectReason = "PLAIN_MATERIAL": GoTo RejectAA
  P3UpdateResidualNormOnly targetForce, internalForce
  plainRes = ResidualNormFree
  P3RestoreCommittedMaterialState False
  For i = 0 To lastDof: TDisp(i) = candidate(i): UDisp(i) = candidate(i): Next i
  If Not P3EvaluateTrialState(incrementBase, internalForce) Then adaptiveRejectReason = "CANDIDATE_MATERIAL": GoTo RejectAA
  P3UpdateResidualMetrics targetForce, internalForce
  candidateRes = ResidualNormFree
  If candidateRes >= plainRes Or candidateRes >= residualBefore * 0.9999 Then adaptiveRejectReason = "NO_RESIDUAL_GAIN": GoTo RejectAA
  If P3FactorPlasticXorCount() <> 0 Then adaptiveRejectReason = "ACTIVE_SET": GoTo RejectAA
  If Not AccelSmallTangentChange() Then adaptiveRejectReason = "TANGENT_OR_BRANCH": GoTo RejectAA
  For i = 0 To lastDof: Disp(i) = candidate(i) - x0(i): Next i
  P3ApplyCorrectionRatioFromDir 1#, x0, correctionRatio
  P6LsNoteAccept 1#, False
  P6LsNoteQ residualBefore, ResidualNormFree
  mAAX = x0: mAAF = f0: mAAFactor = P6FactorizationCount
  AccelAACount = AccelAACount + 1
  If AdaptiveEnabled Then AdaptiveNoteAA plainRes, candidateRes, P6ElapsedMs(adaptiveStart), True, "RESIDUAL_GAIN"
  If AccelTrace Then P6SolverEvent "ANDERSON_ACCEPT", "gamma=" & Format$(gamma, "0.000") & ";plain=" & Format$(plainRes, "0.000E+00") & ";candidate=" & Format$(candidateRes, "0.000E+00")
  P3AccelTryAA = True
  Exit Function
RejectAA:
  If AdaptiveEnabled Then AdaptiveNoteAA plainRes, candidateRes, P6ElapsedMs(adaptiveStart), False, adaptiveRejectReason
  AccelAAReject = AccelAAReject + 1
  mAAReady = False
  P3RestoreCommittedMaterialState False
  For i = 0 To lastDof: TDisp(i) = x0(i): UDisp(i) = x0(i): Disp(i) = f0(i): Next i
  P3TrialStateValid = False
  ' Fatal errors are handled by the outer retry/abort path; transient candidate failure is discarded.
  If Not P3IsFatalFailure() Then P3ClearTransientFailure
End Function

Private Function P3AccelApplyCorrection(ByRef incrementBase() As Double, ByRef targetForce() As Double, ByRef internalForce() As Double, ByRef correctionRatio As Double, ByVal residualBefore As Double) As Boolean
  Dim cutsBefore As Long, ok As Boolean
  mAccelAttemptFailure = ""
  mLineSearchFailure = "": mLineSearchTries = 0: mLineSearchMaterialRejects = 0: mLineSearchBestQ = -1#
  cutsBefore = P3LineSearchCuts
  If P3AccelTryAA(incrementBase, targetForce, internalForce, correctionRatio, residualBefore) Then
    P3AccelApplyCorrection = True: Exit Function
  End If
  If P3IsFatalFailure() Then Exit Function
  ok = P3ApplyLineSearch(incrementBase, targetForce, internalForce, correctionRatio, residualBefore)
  If Not ok Or P3LineSearchCuts <> cutsBefore Then mAAReady = False
  P3AccelApplyCorrection = ok
End Function

Private Function V3ElapsedSeconds(ByVal startedAt As Double) As Double
  V3ElapsedSeconds = Timer - startedAt
  If V3ElapsedSeconds < 0# Then V3ElapsedSeconds = V3ElapsedSeconds + 86400#
End Function

Private Sub V3BeginTrialCost()
  V3TrialStartedAt = Timer
End Sub

Private Sub V3RecordTrialCost(ByVal passed As Boolean)
  Dim elapsed As Double
  elapsed = V3ElapsedSeconds(V3TrialStartedAt)
  If elapsed <= 0# Then Exit Sub
  If passed Then
    V3LastPassCostSec = elapsed
    V3PassCostSamples = V3PassCostSamples + 1
  Else
    V3LastFailCostSec = elapsed
    V3FailCostSamples = V3FailCostSamples + 1
  End If
End Sub

Public Function P3AccelFailureReason() As String
  P3AccelFailureReason = mAccelAttemptFailure
  If Len(P3AccelFailureReason) = 0 Then
    If AnalysisErrorNumber <> 0 Then
      P3AccelFailureReason = "ANALYSIS_" & ResultStatus & "_" & CStr(AnalysisErrorNumber)
    Else
      P3AccelFailureReason = "ATTEMPT_NOT_CONVERGED"
    End If
  End If
End Function








