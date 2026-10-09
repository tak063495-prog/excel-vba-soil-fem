Option Explicit
' Optional V0--V5 experiments. No changes to equilibrium or material tolerances.
Public AccelV1 As Boolean, AccelV2a As Boolean, AccelV2b As Boolean
Public AccelV4 As Boolean, AccelV5 As Boolean, AccelTrace As Boolean
Public AccelBaselineRetry As Boolean
Public AccelPredictBeta As Double, AccelPredictMaxRatio As Double, AccelTangentLimit As Double
Public AccelHistoryReady As Boolean, AccelPreviousStep As Double
Public AccelDelta() As Double
Public AccelFirstReusePending As Boolean, AccelIncrementEligible As Boolean
Public AccelV2aPreviousIterations As Long
Public AccelV2aPreviousStep As Double
Public AccelV2aLongCandidates As Long, AccelV2aLongReuses As Long
Public AccelPredictCount As Long, AccelFallbackCount As Long, AccelFirstReuseCount As Long
Public AccelCostReuseCount As Long, AccelAACount As Long, AccelAAReject As Long
Public AccelGMRESCount As Long, AccelGMRESFallback As Long
Public AccelSrmWidthTarget As Double
Public AccelSubstepCalls(0 To 8) As Long
Public AccelMaterialCalls As Long, AccelMaterialSuccess As Long, AccelBranchGateReject As Long
Private mFactorSp() As Double, mFactorElem As Long, mFactorReady As Boolean
Private mFactorBranch() As Byte
Private mLastQ As Double, mBuildQ As Double
Private mEvalStart As Double, mLastIterationMs As Double, mIterationFactorMs As Double
Private mReasons As Object

Public Sub AccelResetRun()
  Dim i As Long
  AccelV1 = (P6ReadSetting("ACCEL_V1_PREDICTOR", 0#) = 1#)
  AccelV2a = (P6ReadSetting("ACCEL_V2A_REUSE", 0#) = 1#)
  AccelV2b = (P6ReadSetting("ACCEL_V2B_COST", 0#) = 1#)
  AccelV4 = (P6ReadSetting("ACCEL_V4_ANDERSON", 0#) = 1#)
  AccelV5 = (P6ReadSetting("ACCEL_V5_GMRES_LU", 0#) = 1#)
  AccelTrace = (P6ReadSetting("ACCEL_TRACE", 0#) = 1#)
  If P3FlowPolicyIsInconsistent() Then
    ' Full nonsymmetric Newton cannot use factors from the symmetric BAND path.
    AccelV2a = False: AccelV2b = False: AccelV5 = False
  End If
  AccelPredictBeta = P6ReadSetting("ACCEL_PREDICT_BETA", 1#)
  If AccelPredictBeta < 0# Or AccelPredictBeta > 1# Then AccelPredictBeta = 1#
  AccelPredictMaxRatio = P6ReadSetting("ACCEL_PREDICT_MAX_RATIO", 1.5)
  If AccelPredictMaxRatio <= 0# Or AccelPredictMaxRatio > 2# Then AccelPredictMaxRatio = 1.5
  AccelTangentLimit = P6ReadSetting("ACCEL_TANGENT_CHANGE", 0.05)
  If AccelTangentLimit <= 0# Or AccelTangentLimit > 0.1 Then AccelTangentLimit = 0.05
  Set mReasons = CreateObject("Scripting.Dictionary")
  AccelPredictCount = 0: AccelFallbackCount = 0: AccelFirstReuseCount = 0
  AccelCostReuseCount = 0: AccelAACount = 0: AccelAAReject = 0
  AccelV2aLongCandidates = 0: AccelV2aLongReuses = 0
  AccelGMRESCount = 0: AccelGMRESFallback = 0
  AccelSrmWidthTarget = 0#
  AccelMaterialCalls = 0: AccelMaterialSuccess = 0: AccelBranchGateReject = 0
  For i = 0 To 8: AccelSubstepCalls(i) = 0: Next i
  PlasticLogResetRun
  AdaptiveResetRun
  AccelResetStage
End Sub

Public Sub AccelResetStage()
  AdaptiveResetStage
  IncrementLogResetStage
  AccelHistoryReady = False: AccelPreviousStep = 0#
  AccelFirstReusePending = False: AccelIncrementEligible = False
  AccelV2aPreviousIterations = 0: AccelV2aPreviousStep = 0#
  AccelBaselineRetry = False: mFactorReady = False
  mLastQ = 0#: mBuildQ = 0#: mLastIterationMs = 0#
  P6AccelResetPreconditioner
End Sub

Public Sub AccelRecordReason(ByVal reason As String)
  If mReasons Is Nothing Then Set mReasons = CreateObject("Scripting.Dictionary")
  IncrementLogNoteRebuild reason
  AdaptiveV2aRecordRebuild reason
  If Not mReasons.Exists(reason) Then mReasons.Add reason, 0&
  mReasons(reason) = CLng(mReasons(reason)) + 1
  If AccelTrace Then P6SolverEvent "TANGENT_REBUILD", "reason=" & reason & ";iter=" & CStr(CurrentIteration)
End Sub

Public Sub AccelCaptureFactorTangent()
  Dim k As Long, gp As Long, i As Long, j As Long
  mFactorReady = False
  On Error GoTo CaptureSkip
  If Not AccelV2a And Not AccelV2b And Not AccelV4 Then Exit Sub
  If NumberOfElement < 1 Or P3HasJointElements Or P6MixedUP Then Exit Sub
  If mFactorElem <> NumberOfElement Then
    ReDim mFactorSp(0 To 2, 0 To 15, 0 To 3, 0 To NumberOfElement - 1)
    ReDim mFactorBranch(0 To 3, 0 To NumberOfElement - 1)
  End If
  mFactorElem = NumberOfElement
  For k = 0 To NumberOfElement - 1
    For gp = 0 To 3
      mFactorBranch(gp, k) = AccelMaterialBranch(k, gp)
      For i = 0 To 2
        For j = 0 To 15
          If Elem(k).SpmatValid(gp) Then
            mFactorSp(i, j, gp, k) = Elem(k).Spmat(i, j, gp)
          Else
            mFactorSp(i, j, gp, k) = Elem(k).Smat(i, j, gp)
          End If
        Next j
      Next i
    Next gp
  Next k
  mFactorReady = True
  Exit Sub
CaptureSkip:
  mFactorReady = False
  If Err.Number = 18 Then Err.Raise 18
  Err.Clear
End Sub

Public Function AccelSmallTangentChange(Optional ByRef gateReason As String = "") As Boolean
  Dim k As Long, gp As Long, i As Long, j As Long
  Dim num As Double, den As Double, v As Double, d As Double
  AccelSmallTangentChange = False: gateReason = "FACTOR_SNAPSHOT"
  If Not mFactorReady Or mFactorElem <> NumberOfElement Then Exit Function
  If P3HasJointElements Or P6MixedUP Then Exit Function
  If Not P6CanReuseAssembledFactor() Then gateReason = "SIGNATURE": Exit Function
  If P3FactorPlasticXorCount() <> 0 Then gateReason = "PLASTIC_XOR": Exit Function
  For k = 0 To NumberOfElement - 1
    If P3IsElementActive(k) Then
      num = 0#: den = 0#
      For gp = 0 To 3
        If AccelMaterialBranch(k, gp) <> mFactorBranch(gp, k) Then
          gateReason = "MATERIAL_BRANCH"
          AccelBranchGateReject = AccelBranchGateReject + 1
          Exit Function
        End If
        For i = 0 To 2
          For j = 0 To 15
            If Elem(k).SpmatValid(gp) Then v = Elem(k).Spmat(i, j, gp) Else v = Elem(k).Smat(i, j, gp)
            d = v - mFactorSp(i, j, gp, k)
            num = num + d * d: den = den + mFactorSp(i, j, gp, k) ^ 2
          Next j
        Next i
      Next gp
      If den <= 1E-30 Then
        If num > 1E-30 Then gateReason = "TANGENT_CHANGE": Exit Function
      ElseIf num > AccelTangentLimit * AccelTangentLimit * den Then
        gateReason = "TANGENT_CHANGE": Exit Function
      End If
    End If
  Next k
  gateReason = "OK": AccelSmallTangentChange = True
End Function

Private Function AccelMaterialBranch(ByVal k As Long, ByVal gp As Long) As Byte
  Dim stressScale As Double, eps As Double, upperEdge As Boolean, lowerEdge As Boolean
  ' Diagnostic only: principal-stress degeneracy guards same-count face/edge changes.
  If Not Elem(k).P2Yielded(gp) Then Exit Function
  stressScale = Abs(Elem(k).P2PrincipalStress(0, gp)) + Abs(Elem(k).P2PrincipalStress(1, gp)) + Abs(Elem(k).P2PrincipalStress(2, gp))
  If stressScale < 1# Then stressScale = 1#
  eps = 0.0000001 * stressScale
  upperEdge = (Abs(Elem(k).P2PrincipalStress(0, gp) - Elem(k).P2PrincipalStress(1, gp)) <= eps)
  lowerEdge = (Abs(Elem(k).P2PrincipalStress(1, gp) - Elem(k).P2PrincipalStress(2, gp)) <= eps)
  AccelMaterialBranch = 1
  If upperEdge Then AccelMaterialBranch = 2
  If lowerEdge Then AccelMaterialBranch = 3
  If upperEdge And lowerEdge Then AccelMaterialBranch = 4
End Function

Public Sub AccelRememberIncrement(ByRef base() As Double, ByVal deltaLambda As Double, ByVal iterations As Long, ByVal retries As Long)
  Dim i As Long
  AccelHistoryReady = False: AccelIncrementEligible = False
  AccelV2aPreviousIterations = 0: AccelV2aPreviousStep = 0#
  ' Both methods require a clean, normally committed increment in supported models.
  If deltaLambda <= 0# Or retries <> 0 Or AccelBaselineRetry Or P6MixedUP Or P3HasJointElements Then Exit Sub
  ' V2a eligibility is independent of predictor quality. Matrix/state guards run at use.
  If Not P3AfterCutback Then
    AccelIncrementEligible = True
    AccelV2aPreviousIterations = iterations: AccelV2aPreviousStep = deltaLambda
    If AccelV2a And iterations > 5 Then AccelV2aLongCandidates = AccelV2aLongCandidates + 1
  End If
  ' Keep the existing V1 limit and leave old predictor arrays inaccessible when rejected.
  If iterations > 5 Then Exit Sub
  ReDim AccelDelta(lastDof)
  For i = 0 To lastDof
    If NodeCond(i) = 0 Then AccelDelta(i) = TDisp(i) - base(i)
  Next i
  AccelPreviousStep = deltaLambda
  AccelHistoryReady = True
End Sub

Public Sub AccelBeginIteration()
  AdaptiveBeginIteration
  mEvalStart = Timer: mIterationFactorMs = P6ProfFactorMs
End Sub

Public Sub AccelNoteCorrection(ByVal beforeResidual As Double, ByVal afterResidual As Double, ByVal rebuilt As Boolean)
  Dim q As Double
  If beforeResidual <= 1E-30 Then Exit Sub
  q = afterResidual / beforeResidual
  If rebuilt Then mBuildQ = q
  mLastQ = q
  mLastIterationMs = P6ElapsedMs(mEvalStart)
  AdaptiveNoteCorrection beforeResidual, afterResidual, rebuilt, mLastIterationMs, P6ProfFactorMs - mIterationFactorMs
End Sub

Public Function AccelCostAllowsReuse(ByVal q As Double) As Boolean
  Dim FactorMs As Double, extraIters As Double
  AccelCostAllowsReuse = False
  If Not AccelV2b Or AccelBaselineRetry Then Exit Function
  If Not AdaptiveAllowMethod(3) Then Exit Function
  If q <= 0# Or q > 0.85 Or mLastQ <= 0# Or mLastQ > 0.85 Then Exit Function
  If Abs(q - mLastQ) > 0.1 Or mBuildQ <= 0# Or mBuildQ >= 0.7 Then Exit Function
  If P3TangentStaleIters >= 2 Or P3AfterCutback Then Exit Function
  If Not AccelSmallTangentChange() Then Exit Function
  If P6FactorizationCount < 2 Or mLastIterationMs <= 0# Then Exit Function
  FactorMs = P6ProfFactorMs / P6FactorizationCount
  ' Extra work to gain the same logarithmic reduction as the last fresh tangent.
  extraIters = Log(mBuildQ) / Log(q) - 1#
  If extraIters < 0# Then extraIters = 0#
  If (1# + extraIters) * mLastIterationMs >= FactorMs Then Exit Function
  AdaptiveUseMethod 3
  AccelCostAllowsReuse = True
  AccelCostReuseCount = AccelCostReuseCount + 1
  If AccelTrace Then P6SolverEvent "V2B_REUSE", "q=" & Format$(q, "0.000") & ";factor_ms=" & Format$(FactorMs, "0.000")
End Function

Public Sub AccelNoteSubsteps(ByVal n As Long)
  Dim bucket As Long
  bucket = 0
  Do While n > 1 And bucket < 8
    n = (n + 1) \ 2: bucket = bucket + 1
  Loop
  AccelSubstepCalls(bucket) = AccelSubstepCalls(bucket) + 1
  AccelMaterialSuccess = AccelMaterialSuccess + 1
End Sub

Public Function AccelSummary() As String
  Dim s As String, key As Variant, i As Long
  s = "AccelFlags=" & CStr(AccelV1) & "," & CStr(AccelV2a) & "," & CStr(AccelV2b) & "," & CStr(AccelV4) & "," & CStr(AccelV5) & vbCrLf
  s = s & "AccelV3CostSearch=" & CStr(P6ReadSetting("ACCEL_V3_COST_SEARCH", 0#) = 1#) & vbCrLf
  s = s & "SRM_CONFIGURED_TOL=" & Format$(P6ReadSetting("SRM_TOL", 0.025), "0.000000000000") & vbCrLf
  If AccelSrmWidthTarget > 0# Then
    s = s & "SRM_COMPARISON_WIDTH_TARGET=" & Format$(AccelSrmWidthTarget, "0.000000000000") & vbCrLf
  Else
    s = s & "SRM_COMPARISON_WIDTH_TARGET=NA" & vbCrLf
  End If
  s = s & "FOS_PASS_EXACT=" & Format$(P3FosPass, "0.000000000000") & vbCrLf
  If P3FosBracket Then
    s = s & "FOS_FAIL_EXACT=" & Format$(P3FosFail, "0.000000000000") & vbCrLf
    s = s & "FOS_WIDTH_EXACT=" & Format$(P3FosWidth, "0.000000000000") & vbCrLf
  Else
    s = s & "FOS_FAIL_EXACT=NA" & vbCrLf & "FOS_WIDTH_EXACT=NA" & vbCrLf
  End If
  s = s & "PredictorCount=" & CStr(AccelPredictCount) & vbCrLf & "AccelerationFallbackCount=" & CStr(AccelFallbackCount) & vbCrLf
  s = s & "V2aLongIncrementCandidateCount=" & CStr(AccelV2aLongCandidates) & vbCrLf
  s = s & "V2aLongIncrementReuseCount=" & CStr(AccelV2aLongReuses) & vbCrLf
  s = s & "FirstApproximateReuseCount=" & CStr(AccelFirstReuseCount) & vbCrLf & "CostReuseCount=" & CStr(AccelCostReuseCount) & vbCrLf
  s = s & "AndersonAccept=" & CStr(AccelAACount) & vbCrLf & "AndersonReject=" & CStr(AccelAAReject) & vbCrLf
  s = s & "GMRES_LU_Accept=" & CStr(AccelGMRESCount) & vbCrLf & "GMRES_LU_Fallback=" & CStr(AccelGMRESFallback) & vbCrLf
  s = s & "MaterialUpdateCalls=" & CStr(AccelMaterialCalls) & vbCrLf & "MaterialUpdateFailed=" & CStr(AccelMaterialCalls - AccelMaterialSuccess) & vbCrLf
  s = s & "MaterialBranchGateReject=" & CStr(AccelBranchGateReject) & vbCrLf
  For i = 0 To 8: s = s & "MaterialSubstepBucket_" & CStr(2 ^ i) & "=" & CStr(AccelSubstepCalls(i)) & vbCrLf: Next i
  If Not mReasons Is Nothing Then
    For Each key In mReasons.Keys: s = s & "REBUILD_EXACT_" & CStr(key) & "=" & CStr(mReasons(key)) & vbCrLf: Next key
  End If
  s = s & AdaptiveSummary()
  AccelSummary = s
End Function
