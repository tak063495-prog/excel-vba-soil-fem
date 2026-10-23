Option Explicit
Private mParent As Object, mElements() As Element_Data
Private mStageHydraulics As Variant, mStageCount As Long, mStageGuard As Boolean
Private mPlastic As Variant
Public FSSrmBranch As Boolean

Public Sub FSCaptureSrmParent()
  Set mParent = CreateObject("Scripting.Dictionary")
  mElements = Elem
  mParent.Add "elementActive", P3ElementActive
  mParent.Add "T", TDisp: mParent.Add "U", UDisp: mParent.Add "C", P3CommittedDisp
  mParent.Add "internal", P3FinalInternalForce
  mParent.Add "plastic", P3CommittedPlasticStrain: mParent.Add "mult", P3CommittedPlasticMultiplier
  mParent.Add "yield", P3CommittedYieldFunction: mParent.Add "yielded", P3CommittedYielded
  mParent.Add "principal", P3CommittedPrincipalStress: mParent.Add "spmat", P3CommittedSpmat
  mParent.Add "dpmat", P3CommittedDpmat: mParent.Add "spvalid", P3CommittedSpmatValid
  mParent.Add "activePlastic", P3ActivePlasticPoint: mParent.Add "activeCount", P3ActivePlasticPointCount
  mParent.Add "nodeCond", NodeCond: mParent.Add "boundary", P3BoundaryDisp
  mParent.Add "self", P3SelfWeightForce: mParent.Add "applied", P3AppliedForce
  mParent.Add "lockedSelf", P3LockedSelfWeight: mParent.Add "lockedApplied", P3LockedApplied
  mParent.Add "committedCond", P3CommittedNodeCond: mParent.Add "committedBoundary", P3CommittedBoundaryDisp
  mParent.Add "successInc", P3SuccessfulIncrementCount: mParent.Add "lastInc", P3LastConvergedIncrement
  mParent.Add "maxDisp", P3MaxTrialDisp: mParent.Add "maxCorr", P3MaxCorrection
  mParent.Add "rx", P6ReactionSumX: mParent.Add "ry", P6ReactionSumY
  mParent.Add "residual", RelativeResidualFree: mParent.Add "fullResidual", ResidualNormFull
  mParent.Add "freeResidual", ResidualNormFree: mParent.Add "maxResidual", MaxAbsResidualFree
  mParent.Add "energy", EnergyError
  mParent.Add "increment", FEMIncrementBranchState()
  mPlastic = PlasticLogCaptureBranch()
  FSSrmBranch = True
End Sub

Public Sub FSRestoreSrmParent()
  Dim active() As Boolean, e As Long, changed As Boolean
  If mParent Is Nothing Then FSCoupledInputError "SRMの親状態がありません。"
  Elem = mElements
  active = mParent("elementActive")
  For e = 0 To NumberOfElement - 1
    If P3ElementActive(e) <> active(e) Then changed = True
  Next e
  If changed Then
    P3ElementActive = active
    P3ActiveSetGen = P3ActiveSetGen + 1
    P6InvalidateActiveDependentCaches
  End If
  TDisp = mParent("T"): UDisp = mParent("U"): P3CommittedDisp = mParent("C")
  P3FinalInternalForce = mParent("internal")
  P3CommittedPlasticStrain = mParent("plastic"): P3CommittedPlasticMultiplier = mParent("mult")
  P3CommittedYieldFunction = mParent("yield"): P3CommittedYielded = mParent("yielded")
  P3CommittedPrincipalStress = mParent("principal"): P3CommittedSpmat = mParent("spmat")
  P3CommittedDpmat = mParent("dpmat"): P3CommittedSpmatValid = mParent("spvalid")
  P3ActivePlasticPoint = mParent("activePlastic"): P3ActivePlasticPointCount = mParent("activeCount")
  NodeCond = mParent("nodeCond"): P3BoundaryDisp = mParent("boundary")
  P3SelfWeightForce = mParent("self"): P3AppliedForce = mParent("applied")
  P3LockedSelfWeight = mParent("lockedSelf"): P3LockedApplied = mParent("lockedApplied")
  P3CommittedNodeCond = mParent("committedCond"): P3CommittedBoundaryDisp = mParent("committedBoundary")
  P3SuccessfulIncrementCount = mParent("successInc"): P3LastConvergedIncrement = mParent("lastInc")
  P3MaxTrialDisp = mParent("maxDisp"): P3MaxCorrection = mParent("maxCorr")
  P6ReactionSumX = mParent("rx"): P6ReactionSumY = mParent("ry")
  RelativeResidualFree = mParent("residual"): ResidualNormFull = mParent("fullResidual")
  ResidualNormFree = mParent("freeResidual"): MaxAbsResidualFree = mParent("maxResidual")
  EnergyError = mParent("energy")
  PlasticLogRestoreBranch mPlastic
  FEMIncrementRestoreBranch mParent("increment")
  P3TrialStateValid = False
  P3MarkAllTangentDirty
  P6InvalidateNumericTangent
End Sub

Public Function FSRunSrmTrial() As Boolean
  Dim target() As Double, base() As Double, guess() As Double, internal() As Double, i As Long
  base = P3CommittedDisp: guess = base
  ReDim target(lastDof)
  For i = 0 To lastDof: target(i) = P3LockedSelfWeight(i) + P3LockedApplied(i): Next i
  If Not P3FixedStressTrial(target, base, guess, internal, P3_RESIDUAL_TOLERANCE) Then
    If ResultStatus = RESULT_PASS Or ResultStatus = RESULT_NOT_RUN Or Len(ResultStatus) = 0 Then SetAnalysisFailure RESULT_NONCONVERGED, "親水圧保持のSRM力学試行が収束しません。", vbObjectError + 8122, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
    Exit Function
  End If
  P3FixedStressAccept internal
  FSRunSrmTrial = True
End Function

Public Sub FSBeginStageGuard()
  FSCaptureSrmParent
  FSSrmBranch = False
  mStageHydraulics = FSHydraulicMemory()
  FSCaptureStageWater
  mStageCount = FSAccepted: mStageGuard = True
End Sub
Public Sub FSEndStageGuard(ByVal accepted As Boolean)
  If Not mStageGuard Then Exit Sub
  mStageGuard = False
  If Not accepted And FSAccepted = mStageCount Then
    FSRestoreSrmParent
    FSRestoreHydraulicMemory mStageHydraulics
  End If
  FSSrmBranch = False
End Sub
