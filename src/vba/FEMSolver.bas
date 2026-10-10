Option Explicit
Private mAdaptiveGMRESBudgetStop As Boolean
' Global matrix assembly, RCM, BAND_LU / CSR iterative solve.
Private P6LdltDim As Long
Private P6LdltBw As Long
Private P6LdltRhsReady As Boolean
Private P6LdltBand() As Double
Private P6LdltHold() As Double
Private P6LdltSnapReady As Boolean
Private P6LdltRhs() As Double
Private P6LdltX() As Double
Private P6LdltAx() As Double
Private P6LdltMult() As Double
Private P6BandBenchDone As Boolean
Private P6BandBenchCached As Boolean
Private P6BandBenchOn As Boolean
Private P6BandBenchPending As Boolean
Private P6BandBenchN As Long
Private P6BandBenchBw As Long
Private P6BandBenchSnap() As Double
Private P6BandBenchFree() As Boolean
Private P6KernelCached As Boolean
Private P6KernelNew As Boolean
Private P6PackN As Long
Private P6PackBw As Long
Private P6Packed() As Double
Private P6PackRowOff() As Long
Private P6BandFactorFailK As Long
Private P6BandSolveFailK As Long
Private P6BandRejectReason As String
Private mBandCacheReady As Boolean, mBandCacheN As Long, mBandCacheBw As Long
Private mBandCacheA() As Double, mBandCacheLower() As Double
Private mBandCachePivots() As Long, mBandCacheRows() As Long, mBandCacheColumns() As Long
Private mBandCacheValues() As Double
Private P6CSRPreconditionerMode As String
Private P6CSRILUFailRow As Long
Private P6CSRILUFailColumn As Long
Private P6CSRILUMinAbsDiag As Double
Private mPreReady As Boolean, mPreUses As Long, mPreN As Long, mPreBw As Long
Private mPreActive As Long, mPreConstraint As Long
Private mPreMaterial As Long, mPreStrength As Long, mPreGeometry As Long
Private mPreFs As Double, mPreTol As Double
Private mPrePacked() As Double, mPreOffsets() As Long
Public Sub P6InvalidateCSRConstraintCache()
    P6CSRConstraintVersion = P6CSRConstraintVersion + 1
    If P6CSRConstraintVersion <= 0 Then P6CSRConstraintVersion = 1
    P6CSRBoundaryConstraintVersion = -1
    P6CSRBoundaryApplied = False
End Sub

Public Function P6BandArrayReady(ByRef mat() As Double) As Boolean
  P6BandArrayReady = False
  If lastDof < 0 Or BandWidth < 0 Then Exit Function
  On Error Resume Next
  P6BandArrayReady = (UBound(mat, 1) = lastDof And UBound(mat, 2) = BandWidth)
  If Err.Number <> 0 Then P6BandArrayReady = False
  Err.Clear
  On Error GoTo 0
End Function

Private Sub P6EnsureBandArray(ByRef mat() As Double)
  If Not P6BandArrayReady(mat) Then ReDim mat(lastDof, BandWidth)
End Sub

Private Sub P6CopyBand(ByRef src() As Double, ByRef dst() As Double)
  P6EnsureBandArray dst
  dst = src
End Sub

Function SetBandWidth()
Dim ii As Long, jj As Long
Dim i As Long, j As Long, k As Long
Dim bandValue As Long
'バンド幅の設定
   bandValue = 0
   For k = 0 To NumberOfElement - 1
     For i = 0 To 15
        If Elem(k).ElNode(i) < 0 Then GoTo NextBandI
        For j = 0 To 15
           If Elem(k).ElNode(j) < 0 Then GoTo NextBandJ
           ii = Elem(k).ElNode(i)
           jj = Elem(k).ElNode(j) - ii
           If jj >= 0 Then
              If jj > bandValue Then bandValue = jj
           End If
NextBandJ:
        Next j
NextBandI:
     Next i
   Next k
   SetBandWidth = bandValue
End Function

Private Sub P6AddRCMNeighbor(ByRef neighborList As Collection, ByVal nodeId As Long, ByVal nodeCount As Long)
  Dim itemIndex As Long
  If nodeId < 0 Then Err.Raise vbObjectError + 3460, "FEM.P6AddRCMNeighbor", "RCM節点番号が負です。"
  If nodeId >= nodeCount Then Err.Raise vbObjectError + 3461, "FEM.P6AddRCMNeighbor", "RCM節点番号が範囲外です。"
  For itemIndex = 1 To neighborList.count
    If CLng(neighborList.item(itemIndex)) = nodeId Then Exit Sub
  Next itemIndex
  neighborList.Add nodeId
End Sub

Private Sub P6SortRCMNeighbors(ByRef neighborNodes() As Long, ByVal neighborCount As Long, ByRef degree() As Long)
  Dim outerIndex As Long, innerIndex As Long
  Dim keyNode As Long, previousNode As Long
  Dim movePrevious As Boolean
  For outerIndex = 1 To neighborCount - 1
    keyNode = neighborNodes(outerIndex)
    innerIndex = outerIndex - 1
    Do While innerIndex >= 0
      previousNode = neighborNodes(innerIndex)
      movePrevious = False
      If degree(previousNode) > degree(keyNode) Then
        movePrevious = True
      ElseIf degree(previousNode) = degree(keyNode) Then
        If previousNode > keyNode Then movePrevious = True
      End If
      If Not movePrevious Then Exit Do
      neighborNodes(innerIndex + 1) = previousNode
      innerIndex = innerIndex - 1
    Loop
    neighborNodes(innerIndex + 1) = keyNode
  Next outerIndex
End Sub

Private Function P6BuildRCMOrdering(ByRef oldToNew() As Long, ByRef newToOld() As Long) As Boolean
  Dim adjacency() As Collection
  Dim degree() As Long, visited() As Boolean
  Dim bfsQueue() As Long, cmOrder() As Long, neighborNodes() As Long
  Dim nodeCount As Long, nodeId As Long, elementId As Long
  Dim localI As Long, localJ As Long, nodeA As Long, nodeB As Long
  Dim componentStart As Long, componentDegree As Long
  Dim queueHead As Long, queueTail As Long, currentNode As Long
  Dim neighborCount As Long, itemIndex As Long, neighborNode As Long
  Dim cmCount As Long, newIndex As Long

  P6BuildRCMOrdering = False
  nodeCount = NumberOfFreeNode
  If nodeCount <= 0 Then Exit Function
  If NumberOfElement <= 0 Then Exit Function

  ReDim adjacency(0 To nodeCount - 1)
  ReDim degree(0 To nodeCount - 1)
  ReDim visited(0 To nodeCount - 1)
  ReDim bfsQueue(0 To nodeCount - 1)
  ReDim cmOrder(0 To nodeCount - 1)
  ReDim oldToNew(0 To nodeCount - 1)
  ReDim newToOld(0 To nodeCount - 1)
  For nodeId = 0 To nodeCount - 1
    Set adjacency(nodeId) = New Collection
  Next nodeId

  For elementId = 0 To NumberOfElement - 1
    For localI = 0 To 7
      nodeA = Elem(elementId).node(localI)
      If nodeA < 0 Then GoTo NextRcmAdjI
      For localJ = localI + 1 To 7
        nodeB = Elem(elementId).node(localJ)
        If nodeB < 0 Then GoTo NextRcmAdjJ
        P6AddRCMNeighbor adjacency(nodeA), nodeB, nodeCount
        P6AddRCMNeighbor adjacency(nodeB), nodeA, nodeCount
NextRcmAdjJ:
      Next localJ
NextRcmAdjI:
    Next localI
  Next elementId

  For nodeId = 0 To nodeCount - 1
    degree(nodeId) = adjacency(nodeId).count
  Next nodeId

  cmCount = 0
  Do While cmCount < nodeCount
    componentStart = -1
    componentDegree = nodeCount + 1
    For nodeId = 0 To nodeCount - 1
      If Not visited(nodeId) Then
        If degree(nodeId) < componentDegree Then
          componentStart = nodeId
          componentDegree = degree(nodeId)
        End If
      End If
    Next nodeId
    If componentStart < 0 Then Exit Do

    queueHead = 0
    queueTail = 0
    bfsQueue(0) = componentStart
    visited(componentStart) = True
    Do While queueHead <= queueTail
      currentNode = bfsQueue(queueHead)
      queueHead = queueHead + 1
      cmOrder(cmCount) = currentNode
      cmCount = cmCount + 1
      neighborCount = adjacency(currentNode).count
      If neighborCount > 0 Then
        ReDim neighborNodes(0 To neighborCount - 1)
        For itemIndex = 1 To neighborCount
          neighborNodes(itemIndex - 1) = CLng(adjacency(currentNode).item(itemIndex))
        Next itemIndex
        P6SortRCMNeighbors neighborNodes, neighborCount, degree
        For itemIndex = 0 To neighborCount - 1
          neighborNode = neighborNodes(itemIndex)
          If Not visited(neighborNode) Then
            If queueTail >= nodeCount - 1 Then
              Err.Raise vbObjectError + 3462, "FEM.P6BuildRCMOrdering", "RCMキュー容量を超えました。"
            End If
            queueTail = queueTail + 1
            bfsQueue(queueTail) = neighborNode
            visited(neighborNode) = True
          End If
        Next itemIndex
      End If
    Loop
  Loop

  If cmCount <> nodeCount Then
    Err.Raise vbObjectError + 3463, "FEM.P6BuildRCMOrdering", "RCM順序の生成数が節点数と一致しません。"
  End If
  For newIndex = 0 To nodeCount - 1
    nodeId = cmOrder(nodeCount - 1 - newIndex)
    newToOld(newIndex) = nodeId
    oldToNew(nodeId) = newIndex
  Next newIndex
  P6BuildRCMOrdering = True
End Function

Private Function P6ComputeRCMCandidateBandwidth(ByRef oldToNew() As Long) As Long
  Dim elementId As Long, localI As Long, localJ As Long
  Dim nodeA As Long, nodeB As Long
  Dim dofA As Long, dofB As Long, bandDifference As Long
  Dim candidateBandwidth As Long

  candidateBandwidth = 0
  For elementId = 0 To NumberOfElement - 1
    For localI = 0 To 15
      If Elem(elementId).ElNode(localI) < 0 Then GoTo NextRcmCandI
      nodeA = Elem(elementId).node(localI \ 2)
      If nodeA < 0 Then GoTo NextRcmCandI
      dofA = 2 * oldToNew(nodeA) + localI - (localI \ 2) * 2
      For localJ = 0 To 15
        If Elem(elementId).ElNode(localJ) < 0 Then GoTo NextRcmCandJ
        nodeB = Elem(elementId).node(localJ \ 2)
        If nodeB < 0 Then GoTo NextRcmCandJ
        dofB = 2 * oldToNew(nodeB) + localJ - (localJ \ 2) * 2
        bandDifference = Abs(dofB - dofA)
        If bandDifference > candidateBandwidth Then candidateBandwidth = bandDifference
NextRcmCandJ:
      Next localJ
NextRcmCandI:
    Next localI
  Next elementId
  P6ComputeRCMCandidateBandwidth = candidateBandwidth
End Function

Private Sub P6ValidateRCMPermutation(ByRef oldToNew() As Long, ByRef newToOld() As Long)
  Dim nodeCount As Long, oldNode As Long, newNode As Long
  Dim seenNew() As Boolean

  nodeCount = NumberOfFreeNode
  If nodeCount <= 0 Then Err.Raise vbObjectError + 3510, "FEM.P6ValidateRCMPermutation", "RCM対象節点がありません。"
  If LBound(oldToNew) <> 0 Or UBound(oldToNew) <> nodeCount - 1 Then _
    Err.Raise vbObjectError + 3511, "FEM.P6ValidateRCMPermutation", "oldToNewの範囲が節点数と一致しません。"
  If LBound(newToOld) <> 0 Or UBound(newToOld) <> nodeCount - 1 Then _
    Err.Raise vbObjectError + 3512, "FEM.P6ValidateRCMPermutation", "newToOldの範囲が節点数と一致しません。"
  ReDim seenNew(0 To nodeCount - 1)
  For oldNode = 0 To nodeCount - 1
    newNode = oldToNew(oldNode)
    If newNode < 0 Or newNode >= nodeCount Then _
      Err.Raise vbObjectError + 3513, "FEM.P6ValidateRCMPermutation", "oldToNewが範囲外です。old=" & CStr(oldNode)
    If seenNew(newNode) Then _
      Err.Raise vbObjectError + 3514, "FEM.P6ValidateRCMPermutation", "oldToNewが重複しています。new=" & CStr(newNode)
    If newToOld(newNode) <> oldNode Then _
      Err.Raise vbObjectError + 3515, "FEM.P6ValidateRCMPermutation", "oldToNewとnewToOldが逆写像になっていません。"
    seenNew(newNode) = True
  Next oldNode
End Sub

Private Sub P6InvalidateCSROrdering()
  ' CSR positions use global DOF numbers; element-local elastic caches do not.
  P6InvalidateNumericTangent
  P6InvalidateCSRConstraintCache
  P6CSRPatternReady = False: P6CSRPatternSignature = vbNullString
  P6CSRMaterialGeneration = -1: P6CSRDeviatoricReady = False
  Erase P6CSRRowPtr: Erase P6CSRColumnIndex: Erase P6CSRDiagonalPosition
  Erase P6CSRElementPosition
  Erase P6CSRBoundaryZeroPosition: Erase P6CSRBoundaryDiagonalPosition
  Erase P6CSRBoundaryRHSPosition: Erase P6CSRBoundaryRHSRow: Erase P6CSRBoundaryRHSColumn
  P6CSRBoundaryZeroCount = 0: P6CSRBoundaryDiagonalCount = 0: P6CSRBoundaryRHSCount = 0
  P6CSRScatterCount = 0: P6CSRScatterElementCount = -1
  P6CSRScatterTangentGeneration = -1: P6CSRScatterMaterialGeneration = -1
  P6CSRScatterPatternSignature = vbNullString
  Erase P6CSRScatterStart: Erase P6CSRScatterLocalRow: Erase P6CSRScatterLocalColumn
  Erase P6CSRScatterPosition: Erase P6CSRScatterValueIndex
  P6CSRNNZ = 0: P6CSRValueCapacity = 0: P6CSRStorageBytes = 0#
  Erase P6CSRDiagonalInverse: Erase P6CSRILUValues
  Erase P6CSROriginalValues: Erase P6CSRValues
  MatrixFactored = False
End Sub

Private Sub P6ApplyRCMOrdering(ByRef oldToNew() As Long, ByRef newToOld() As Long)
  Dim vectorSnapshot() As Double, conditionSnapshot() As Long
  Dim mapOldToNew() As Long, mapNewToOld() As Long
  Dim oldDof As Long, newDof As Long, oldNode As Long, newNode As Long
  Dim elementId As Long, localNode As Long, localDof As Long

  ' ByRef引数がP6RCMOldToNew自身のとき、後段のReDimで写像が消える。
  ' 適用と保存の前に必ずローカルへ退避する。
  ReDim mapOldToNew(0 To NumberOfFreeNode - 1)
  ReDim mapNewToOld(0 To NumberOfFreeNode - 1)
  For oldNode = 0 To NumberOfFreeNode - 1
    mapOldToNew(oldNode) = oldToNew(oldNode)
    mapNewToOld(oldNode) = newToOld(oldNode)
  Next oldNode
  P6ValidateRCMPermutation mapOldToNew, mapNewToOld

  ReDim vectorSnapshot(0 To 9, 0 To lastDof)
  ReDim conditionSnapshot(0 To lastDof)
  For oldDof = 0 To lastDof
    vectorSnapshot(0, oldDof) = Force(oldDof)
    vectorSnapshot(1, oldDof) = Disp(oldDof)
    vectorSnapshot(2, oldDof) = UDisp(oldDof)
    vectorSnapshot(3, oldDof) = XXX(oldDof)
    vectorSnapshot(4, oldDof) = RForce(oldDof)
    vectorSnapshot(5, oldDof) = eForce(oldDof)
    vectorSnapshot(6, oldDof) = TDisp(oldDof)
    vectorSnapshot(7, oldDof) = NForce(oldDof)
    vectorSnapshot(8, oldDof) = mRisp(oldDof)
    vectorSnapshot(9, oldDof) = iNForce(oldDof)
    conditionSnapshot(oldDof) = NodeCond(oldDof)
  Next oldDof

  For oldNode = 0 To NumberOfFreeNode - 1
    newNode = mapOldToNew(oldNode)
    For localDof = 0 To 1
      oldDof = 2 * oldNode + localDof
      newDof = 2 * newNode + localDof
      Force(newDof) = vectorSnapshot(0, oldDof)
      Disp(newDof) = vectorSnapshot(1, oldDof)
      UDisp(newDof) = vectorSnapshot(2, oldDof)
      XXX(newDof) = vectorSnapshot(3, oldDof)
      RForce(newDof) = vectorSnapshot(4, oldDof)
      eForce(newDof) = vectorSnapshot(5, oldDof)
      TDisp(newDof) = vectorSnapshot(6, oldDof)
      NForce(newDof) = vectorSnapshot(7, oldDof)
      mRisp(newDof) = vectorSnapshot(8, oldDof)
      iNForce(newDof) = vectorSnapshot(9, oldDof)
      NodeCond(newDof) = conditionSnapshot(oldDof)
    Next localDof
  Next oldNode
  P3BumpConstraintGeneration

  For elementId = 0 To NumberOfElement - 1
    For localNode = 0 To 7
      oldNode = Elem(elementId).node(localNode)
      If oldNode >= 0 Then
        newNode = mapOldToNew(oldNode)
        Elem(elementId).node(localNode) = newNode
      End If
    Next localNode
    For localDof = 0 To 15
      If Elem(elementId).node(localDof \ 2) < 0 Then
        Elem(elementId).ElNode(localDof) = -1
      Else
        newNode = Elem(elementId).node(localDof \ 2)
        Elem(elementId).ElNode(localDof) = 2 * newNode + (localDof Mod 2)
      End If
    Next localDof
  Next elementId

  P6InvalidateCSROrdering

  ReDim P6RCMOldToNew(0 To NumberOfFreeNode - 1)
  ReDim P6RCMNewToOld(0 To NumberOfFreeNode - 1)
  For oldNode = 0 To NumberOfFreeNode - 1
    P6RCMOldToNew(oldNode) = mapOldToNew(oldNode)
    P6RCMNewToOld(oldNode) = mapNewToOld(oldNode)
  Next oldNode
  P6RCMMapReady = True
  P6RenumberingApplied = True
End Sub

Private Sub P6EvaluateRCMBandwidth(ByVal originalBandwidth As Long)
  Dim oldToNew() As Long, newToOld() As Long
  Dim failureNumber As Long, failureDescription As String
  Dim rcmPolicy As String

  P6RCMOriginalBandwidth = originalBandwidth
  P6RCMCandidateBandwidth = P6RCMOriginalBandwidth
  P6RCMStatus = "RCM候補: 評価未完了"
  P6NodeOrderingPolicy = "原節点番号"
  P6RenumberingApplied = False
  P6RCMMapReady = False
  rcmPolicy = UCase$(Trim$(P6ReadTextSetting("RCM_POLICY", "AUTO")))
  If rcmPolicy <> "AUTO" And rcmPolicy <> "ON" And rcmPolicy <> "OFF" Then rcmPolicy = "AUTO"
  P6NodeOrderingPolicy = P6NodeOrderingPolicy & "; RCM_POLICY=" & rcmPolicy
  On Error GoTo RCMFailed
  If NumberOfFreeNode <= 0 Then
    P6RCMStatus = "RCM候補: 節点がないため未評価"
    Exit Sub
  End If
  If NumberOfElement <= 0 Then
    P6RCMStatus = "RCM候補: 要素がないため未評価"
    Exit Sub
  End If
  If rcmPolicy = "OFF" Then
    P6RCMStatus = "RCM候補: 設定により無効"
    Exit Sub
  End If
  If P6RCMCacheReady Then
    P6RCMCandidateBandwidth = P6RCMCacheCandidateBandwidth
    If P6RCMCandidateBandwidth < P6RCMOriginalBandwidth Then
      P6ApplyRCMOrdering P6RCMOldToNew, P6RCMNewToOld
      P6RCMStatus = "RCM再利用: DOF帯域幅 " & CStr(P6RCMOriginalBandwidth) & " -> " & CStr(P6RCMCandidateBandwidth)
      P6NodeOrderingPolicy = "保存済みRCM内部自由度順序を再利用。結果出力は原自由度順へ逆変換; RCM_POLICY=" & rcmPolicy
    Else
      P6RCMStatus = "RCM再利用: 縮小なし（原順序を維持）"
    End If
    Exit Sub
  End If
  If Not P6BuildRCMOrdering(oldToNew, newToOld) Then
    P6RCMStatus = "RCM候補: 順序を生成できないため未評価"
    Exit Sub
  End If
  P6RCMCandidateBandwidth = P6ComputeRCMCandidateBandwidth(oldToNew)
  P6RCMCacheOriginalBandwidth = P6RCMOriginalBandwidth
  P6RCMCacheCandidateBandwidth = P6RCMCandidateBandwidth
  ReDim P6RCMOldToNew(0 To NumberOfFreeNode - 1)
  ReDim P6RCMNewToOld(0 To NumberOfFreeNode - 1)
  For failureNumber = 0 To NumberOfFreeNode - 1
    P6RCMOldToNew(failureNumber) = oldToNew(failureNumber)
    P6RCMNewToOld(failureNumber) = newToOld(failureNumber)
  Next failureNumber
  P6RCMCacheReady = True
  If P6RCMCandidateBandwidth < P6RCMOriginalBandwidth Then
    P6ApplyRCMOrdering oldToNew, newToOld
    P6RCMStatus = "RCM適用: DOF帯域幅 " & CStr(P6RCMOriginalBandwidth) & " -> " & CStr(P6RCMCandidateBandwidth)
    P6NodeOrderingPolicy = "RCM内部自由度順序を適用。結果出力は原自由度順へ逆変換; RCM_POLICY=" & rcmPolicy
  Else
    P6RCMStatus = "RCM候補: 縮小なし（原順序を維持）"
  End If
  Exit Sub
RCMFailed:
  failureNumber = Err.Number
  failureDescription = Err.Description
  P6RCMCandidateBandwidth = P6RCMOriginalBandwidth
  P6RenumberingApplied = False
  P6RCMMapReady = False
  P6RCMStatus = "RCM候補評価失敗 [" & CStr(failureNumber) & "] " & failureDescription
End Sub

Private Function P6ChooseGMRESRestart(ByVal requestedRestart As Long) As Long
  Dim restart As Long, basisBudgetBytes As Double, basisBytes As Double
  restart = requestedRestart
  If restart < 8 Then restart = 8
  If restart > 200 Then restart = 200
  basisBudgetBytes = P6SolverMemoryLimitBytes * 0.25
  If basisBudgetBytes <= 0# Then basisBudgetBytes = 128# * 1024# * 1024#
  Do While restart > 8
    basisBytes = CDbl(nDof) * CDbl(restart + 1) * 8#
    If basisBytes <= basisBudgetBytes Then Exit Do
    restart = restart - 5
    If restart < 8 Then restart = 8
  Loop
  P6ChooseGMRESRestart = restart
End Function

Private Sub P6SelectSolverMode()
  Dim memoryLimitMB As Double, symmetryTolerance As Double
  Dim iterativeMethod As String, policy As String
  Dim requestedRestart As Long
  Dim isLargeProblem As Boolean

  ' 解法選択は小規模問題では初回評価を保持し、大規模問題だけ再評価する。
  ' しきい値は設定シートのE列キー/C列値から読み、設定変更時に無効化する。
  If P6SolverReevaluationThreshold <= 0 Then
    P6SolverReevaluationThreshold = CLng(P6ReadSetting("SOLVER_REEVAL_DOF_THRESHOLD", 2000#))
    If P6SolverReevaluationThreshold < 1 Then P6SolverReevaluationThreshold = 2000
  End If
  isLargeProblem = (nDof >= P6SolverReevaluationThreshold)
  If P6SolverSelectionReady And Not isLargeProblem And Not P3FlowPolicyIsInconsistent() And Not P6UseCSR Then
    P6SolverEvaluationStatus = "reuse_small"
    Exit Sub
  End If
  P6SolverEvaluationCount = P6SolverEvaluationCount + 1
  If isLargeProblem Then
    P6SolverEvaluationStatus = "reevaluate_large"
  Else
    P6SolverEvaluationStatus = "initial_small"
  End If

  ' Nonassociated flow retains the full nonsymmetric tangent. Other modes use BAND.
  P6SolverPolicy = "BAND"
  P6UseCSR = False
  P6SolverMode = "BAND_LU"
  If P3FlowPolicyIsInconsistent() Then
    P6UseCSR = True
    P6SolverPolicy = "NONSYMMETRIC"
    P6SolverMode = "NONSYM_BAND_LU"
  End If
  P6MixedCoupledKrylov = False

  memoryLimitMB = P6ReadSetting("SOLVER_MEMORY_LIMIT_MB", 512#)
  If memoryLimitMB < 1# Then memoryLimitMB = 512#
  P6SolverMemoryLimitBytes = memoryLimitMB * 1024# * 1024#
  P6IterativeTolerance = P6ReadSetting("SOLVER_TOLERANCE", 0.00000001)
  If P6IterativeTolerance <= 0# Then P6IterativeTolerance = 0.00000001
  P6IterativeMaxIterations = CLng(P6ReadSetting("SOLVER_MAX_ITERATIONS", 2000#))
  If P6IterativeMaxIterations < 1 Then P6IterativeMaxIterations = 2000
  requestedRestart = CLng(P6ReadSetting("SOLVER_GMRES_RESTART", 30#))
  P6GMRESRestartRequested = requestedRestart
  P6GMRESRestart = P6ChooseGMRESRestart(requestedRestart)
  P6GMRESAutoAdjusted = (P6GMRESRestart <> P6GMRESRestartRequested)

  P6BandStorageEntries = CDbl(nDof) * (CDbl(BandWidth) + 1#)
  P6BandStorageBytes = P6BandStorageEntries * 8#
  P6FactorStorageBytes = P6BandStorageBytes
  P6DenseStorageBytes = CDbl(nDof) * CDbl(nDof) * 8#
  P6EstimatedWorkMemoryBytes = P6BandStorageBytes * 2#
  P6MatrixStoragePolicy = "帯域格納＋分解済み帯域因子; mode=BAND_LU; solver=固定"
  If P6UseCSR Then
    P6MatrixStoragePolicy = "非対称帯域LU（行ピボット）＋CSR真残差照合; total-stress displacement"
  End If
  P6SolverSelectionReady = True
End Sub

Public Sub P6InvalidateNumericTangent()
  ' Only the assembled tangent changes; preserve geometry and initial elastic caches.
  P6TangentGeneration = P6TangentGeneration + 1
  If P6TangentGeneration <= 0 Then P6TangentGeneration = 1
  P6CSRNumericGeneration = -1
  P6CSRReady = False
  P6CSRBoundaryApplied = False
  P6CSRDiagonalInverseReady = False
  P6CSRILUReady = False
  P6FactorReady = False
End Sub

Public Function P6GetInternalFreeNode(ByVal originalNode As Long) As Long
  P6GetInternalFreeNode = originalNode
  If originalNode < 0 Or originalNode >= NumberOfFreeNode Then Exit Function
  If P6RCMMapReady Then P6GetInternalFreeNode = P6RCMOldToNew(originalNode)
End Function

Public Function P6GetOriginalFreeNode(ByVal internalNode As Long) As Long
  P6GetOriginalFreeNode = internalNode
  If internalNode < 0 Or internalNode >= NumberOfFreeNode Then Exit Function
  If P6RCMMapReady Then P6GetOriginalFreeNode = P6RCMNewToOld(internalNode)
End Function

Private Sub P6EnsureBandScatter()
  Dim k As Long, i As Long, j As Long, ii As Long, jj As Long, n As Long
  If P6ScatterReady And P6ScatterLastDof = lastDof And P6ScatterBand = BandWidth _
     And P6ScatterActiveGen = P3ActiveSetGen And P6ScatterRcm = P6RenumberingApplied Then
    Exit Sub
  End If
  n = 0
  For k = 0 To NumberOfElement - 1
    If P3IsElementActive(k) Then
      For i = 0 To 15
        If Elem(k).ElNode(i) < 0 Then GoTo NextScatterPairA
        For j = 0 To 15
          If Elem(k).ElNode(j) < 0 Then GoTo NextScatterColA
          ii = Elem(k).ElNode(i)
          jj = Elem(k).ElNode(j) - ii
          If jj >= 0 Then
            If ii < 0 Or ii > lastDof Then Err.Raise vbObjectError + 3307, "FEM.P6EnsureBandScatter", "剛性組立の行番号が範囲外です。要素=" & CStr(k + 1) & " ii=" & CStr(ii)
            If jj > BandWidth Then Err.Raise vbObjectError + 3307, "FEM.P6EnsureBandScatter", "剛性組立が帯域幅を超えました。要素=" & CStr(k + 1) & " jj=" & CStr(jj) & " Band=" & CStr(BandWidth)
            n = n + 1
          End If
NextScatterColA:
        Next j
NextScatterPairA:
      Next i
    End If
  Next k
  If n < 1 Then
    P6ScatterN = 0
    P6ScatterReady = True
    P6ScatterLastDof = lastDof
    P6ScatterBand = BandWidth
    P6ScatterActiveGen = P3ActiveSetGen
    P6ScatterRcm = P6RenumberingApplied
    Exit Sub
  End If
  ReDim P6ScatterRow(0 To n - 1)
  ReDim P6ScatterCol(0 To n - 1)
  ReDim P6ScatterElem(0 To n - 1)
  ReDim P6ScatterI(0 To n - 1)
  ReDim P6ScatterJ(0 To n - 1)
  n = 0
  For k = 0 To NumberOfElement - 1
    If P3IsElementActive(k) Then
      For i = 0 To 15
        If Elem(k).ElNode(i) < 0 Then GoTo NextScatterPairB
        For j = 0 To 15
          If Elem(k).ElNode(j) < 0 Then GoTo NextScatterColB
          ii = Elem(k).ElNode(i)
          jj = Elem(k).ElNode(j) - ii
          If jj >= 0 Then
            P6ScatterRow(n) = ii
            P6ScatterCol(n) = jj
            P6ScatterElem(n) = k
            P6ScatterI(n) = i
            P6ScatterJ(n) = j
            n = n + 1
          End If
NextScatterColB:
        Next j
NextScatterPairB:
      Next i
    End If
  Next k
  P6ScatterN = n
  P6ScatterReady = True
  P6ScatterLastDof = lastDof
  P6ScatterBand = BandWidth
  P6ScatterActiveGen = P3ActiveSetGen
  P6ScatterRcm = P6RenumberingApplied
End Sub

Sub SetTotalMat()
Dim ii As Long, jj As Long
Dim matrixMax As Double: Dim symmetryMax As Double
Dim i As Long, j As Long, k As Long
Dim rcmPolicy As String, meshCacheKey As String
Dim t0 As Double, s As Long
Dim errNum As Long, errSrc As String, errDesc As String
'全体剛性行列の設定
 t0 = Timer
 On Error GoTo AssembleFail
 rcmPolicy = UCase$(Trim$(P6ReadTextSetting("RCM_POLICY", "AUTO")))
 If rcmPolicy <> "AUTO" And rcmPolicy <> "ON" And rcmPolicy <> "OFF" Then rcmPolicy = "AUTO"
 meshCacheKey = P6RCMCacheSignature & "|GEOMETRY=" & P6ComputeGeometrySignature() & "|RCM_POLICY=" & rcmPolicy
 If Not P6MeshInvariantReady Or P6MeshInvariantSignature <> meshCacheKey Then
   P6SolverSelectionReady = False
   BandWidth = SetBandWidth()
   P6EvaluateRCMBandwidth BandWidth
   If P6RenumberingApplied Then
     BandWidth = P6RCMCandidateBandwidth
   Else
     BandWidth = P6RCMOriginalBandwidth
   End If
   P6MeshInvariantOriginalBandwidth = P6RCMOriginalBandwidth
   P6MeshInvariantCandidateBandwidth = P6RCMCandidateBandwidth
   P6MeshInvariantRCMApplied = P6RenumberingApplied
   P6MeshInvariantSignature = meshCacheKey
   P6MeshInvariantReady = True
 Else
   P6RCMOriginalBandwidth = P6MeshInvariantOriginalBandwidth
   P6RCMCandidateBandwidth = P6MeshInvariantCandidateBandwidth
   P6RenumberingApplied = P6MeshInvariantRCMApplied
   If P6MeshInvariantRCMApplied Then
     If Not P6RCMMapReady Then P6ApplyRCMOrdering P6RCMOldToNew, P6RCMNewToOld
     BandWidth = P6MeshInvariantCandidateBandwidth
     P6RCMStatus = "RCM再利用: DOF帯域幅 " & CStr(P6MeshInvariantOriginalBandwidth) & " -> " & CStr(P6MeshInvariantCandidateBandwidth)
     P6NodeOrderingPolicy = "保存済みRCM内部自由度順序を再利用。結果出力は原自由度順へ逆変換; RCM_POLICY=" & rcmPolicy
   Else
     BandWidth = P6MeshInvariantOriginalBandwidth
     P6RCMStatus = "RCM再利用: 縮小なし（原順序を維持）"
     P6NodeOrderingPolicy = "原節点番号; RCM_POLICY=" & rcmPolicy
   End If
 End If

   matrixMax = 0#
   symmetryMax = 0#
   For k = 0 To NumberOfElement - 1
     If P3IsElementActive(k) Then
     For i = 0 To 15
       For j = 0 To 15
         If Abs(Elem(k).kmat(i, j)) > matrixMax Then matrixMax = Abs(Elem(k).kmat(i, j))
         If Abs(Elem(k).kmat(i, j) - Elem(k).kmat(j, i)) > symmetryMax Then symmetryMax = Abs(Elem(k).kmat(i, j) - Elem(k).kmat(j, i))
       Next j
     Next i
     End If
   Next k
   If matrixMax <= 1E-30 Then
     MatrixSymmetryError = 0#
   Else
     MatrixSymmetryError = symmetryMax / matrixMax
   End If

 P6SelectSolverMode
 If Not P6UseCSR Then
   If P6EstimatedWorkMemoryBytes > P6SolverMemoryLimitBytes And P6SolverMemoryLimitBytes > 0# Then
     Err.Raise vbObjectError + 3510, "FEM.SetTotalMat", "BAND作業メモリ概算=" & Format$(P6EstimatedWorkMemoryBytes / 1024# / 1024#, "0.0") & "MB が上限 " & Format$(P6SolverMemoryLimitBytes / 1024# / 1024#, "0.0") & "MB を超えます。設定キー SOLVER_MEMORY_LIMIT_MB を確認してください。"
   End If
 End If
If P6UseCSR Then
  If Not P6BuildCSRFromElements() Then Err.Raise vbObjectError + 3501, "FEM.SetTotalMat", "CSR疎行列の組立に失敗しました。"
  Erase TotalMat
  Erase TotalMat2
  Erase OriginalMat
Else
  ReDim OriginalMat(lastDof, BandWidth)
  P6EnsureBandScatter
  For s = 0 To P6ScatterN - 1
    OriginalMat(P6ScatterRow(s), P6ScatterCol(s)) = OriginalMat(P6ScatterRow(s), P6ScatterCol(s)) + Elem(P6ScatterElem(s)).kmat(P6ScatterI(s), P6ScatterJ(s))
  Next s
  If P6PrepareBiotMixed() Then
    If P6MixedUP Then
      P6CopyBand OriginalMat, TotalMat
      P6AddSchurContribution
      P6MatrixStoragePolicy = "帯域格納＋混合Schur; mode=BAND_LU; Biot u-p; p=" & CStr(P6PressureCount) & "; drain=" & CStr(P6DrainCount)
    End If
  End If
  Erase TotalMat2
End If
P6FactorReady = False
P6FactorPivotTolerance = 0#
MatrixFactored = False
 GoTo AssembleDone
AssembleFail:
 errNum = Err.Number
 errSrc = Err.source
 errDesc = Err.Description
 Resume AssembleDone
AssembleDone:
 On Error GoTo 0
 P6ProfAssembleCount = P6ProfAssembleCount + 1
 P6ProfAssembleMs = P6ProfAssembleMs + P6ElapsedMs(t0)
 If errNum <> 0 Then Err.Raise errNum, errSrc, errDesc
End Sub
Sub SetBoundaryCondition()
Dim i As Long
'変位境界条件の設定
   If P6ForceBeforeBoundaryN <> lastDof Then
     ReDim P6ForceBeforeBoundary(lastDof)
     P6ForceBeforeBoundaryN = lastDof
   End If
   For i = 0 To lastDof
     P6ForceBeforeBoundary(i) = Force(i)
   Next i
   P6ForceBeforeBoundaryReady = True
   If P6UseCSR Then
       If Not P6CSRBoundaryApplied Then
           If Not ApplyCSRBoundaryToMatrix() Then Err.Raise vbObjectError + 3502, "FEM.SetBoundaryCondition", "CSR境界条件の適用に失敗しました。"
           P6CSRBoundaryApplied = True
       End If
       ApplyCSRBoundaryRHS
       For i = 0 To lastDof
           If NodeCond(i) <> 0 Then
               Force(i) = Disp(i)
           End If
       Next i
       Exit Sub
   End If
   For i = 0 To lastDof
       If NodeCond(i) <> 0 Then
           Force(i) = 1E+30 * Disp(i)
       End If
   Next i
   If P6MixedUP Then P6ApplySchurRHS
   If Not P6FactorReady Then
     If P6MixedUP Then
       For i = 0 To lastDof
         If NodeCond(i) <> 0 Then TotalMat(i, 0) = 1E+30
       Next i
       P6CopyBand TotalMat, P6FactoredBand
     Else
       P6CopyBand OriginalMat, P6FactoredBand
       For i = 0 To lastDof
         If NodeCond(i) <> 0 Then P6FactoredBand(i, 0) = 1E+30
       Next i
     End If
   End If
End Sub

Private Function P6CSRFindColumnPosition(ByVal rowNo As Long, ByVal columnId As Long) As Long
  Dim position As Long
  P6CSRFindColumnPosition = -1
  If rowNo < 0 Or rowNo >= nDof Then Exit Function
  For position = P6CSRRowPtr(rowNo) To P6CSRRowPtr(rowNo + 1) - 1
    If P6CSRColumnIndex(position) = columnId Then
      P6CSRFindColumnPosition = position
      Exit Function
    End If
  Next position
End Function

Private Function P6BuildCSRPattern() As Boolean
  Dim rowCounts() As Long, writePositions() As Long, markers() As Long, fillMarkers() As Long
  Dim rowHead() As Long, occurrenceElement() As Long, occurrenceLocalRow() As Long, occurrenceNext() As Long
  Dim elementId As Long, localRow As Long, localCol As Long, rowId As Long, columnId As Long
  Dim rowNo As Long, position As Long, totalNNZ As Long, keyColumn As Long, keyPosition As Long
  Dim occurrenceCount As Long, occurrenceIndex As Long

  P6BuildCSRPattern = False
  If nDof <= 0 Or lastDof < 0 Then Exit Function
  ReDim rowCounts(0 To lastDof)
  ReDim rowHead(0 To lastDof)
  For rowNo = 0 To lastDof
    rowHead(rowNo) = -1
  Next rowNo
  occurrenceCount = NumberOfElement * 16
  If occurrenceCount <= 0 Then Exit Function
  ReDim occurrenceElement(0 To occurrenceCount - 1)
  ReDim occurrenceLocalRow(0 To occurrenceCount - 1)
  ReDim occurrenceNext(0 To occurrenceCount - 1)

  ' 同じ全体行が複数要素に分散して現れるため、要素走査順の一時マーカーだけでは
  ' 重複列を除去できない。先に全体行ごとの要素局所行リストを作り、行単位で列を一度だけ登録する。
  occurrenceIndex = 0
  ReDim markers(0 To lastDof)
  For elementId = 0 To NumberOfElement - 1
    For localRow = 0 To 15
      rowId = Elem(elementId).ElNode(localRow)
      If rowId < 0 Then GoTo NextCsrOccRow
      If rowId > lastDof Then Err.Raise vbObjectError + 3505, "FEM.P6BuildCSRPattern", "CSR行番号が範囲外です。"
      occurrenceElement(occurrenceIndex) = elementId
      occurrenceLocalRow(occurrenceIndex) = localRow
      occurrenceNext(occurrenceIndex) = rowHead(rowId)
      rowHead(rowId) = occurrenceIndex
      occurrenceIndex = occurrenceIndex + 1
NextCsrOccRow:
    Next localRow
  Next elementId

  For rowNo = 0 To lastDof
    occurrenceIndex = rowHead(rowNo)
    Do While occurrenceIndex >= 0
      elementId = occurrenceElement(occurrenceIndex)
      localRow = occurrenceLocalRow(occurrenceIndex)
      For localCol = 0 To 15
        columnId = Elem(elementId).ElNode(localCol)
        If columnId < 0 Then GoTo NextCsrMarkCol
        If columnId > lastDof Then Err.Raise vbObjectError + 3506, "FEM.P6BuildCSRPattern", "CSR列番号が範囲外です。"
        If markers(columnId) <> rowNo + 1 Then
          markers(columnId) = rowNo + 1
          rowCounts(rowNo) = rowCounts(rowNo) + 1
        End If
NextCsrMarkCol:
      Next localCol
      occurrenceIndex = occurrenceNext(occurrenceIndex)
    Loop
  Next rowNo

  ReDim P6CSRRowPtr(0 To nDof)
  totalNNZ = 0
  For rowNo = 0 To lastDof
    P6CSRRowPtr(rowNo) = totalNNZ
    totalNNZ = totalNNZ + rowCounts(rowNo)
    P6CSRRowPtr(rowNo + 1) = totalNNZ
  Next rowNo
  If totalNNZ <= 0 Then Exit Function
  P6CSRNNZ = totalNNZ
  ReDim P6CSRColumnIndex(0 To P6CSRNNZ - 1)
  ReDim P6CSRDiagonalPosition(0 To lastDof)
  ReDim writePositions(0 To lastDof)
  ReDim fillMarkers(0 To lastDof)
  For rowNo = 0 To lastDof
    writePositions(rowNo) = P6CSRRowPtr(rowNo)
    P6CSRDiagonalPosition(rowNo) = -1
  Next rowNo

  For rowNo = 0 To lastDof
    occurrenceIndex = rowHead(rowNo)
    Do While occurrenceIndex >= 0
      elementId = occurrenceElement(occurrenceIndex)
      localRow = occurrenceLocalRow(occurrenceIndex)
      For localCol = 0 To 15
        columnId = Elem(elementId).ElNode(localCol)
        If columnId < 0 Then GoTo NextCsrFillCol
        If fillMarkers(columnId) <> rowNo + 1 Then
          fillMarkers(columnId) = rowNo + 1
          position = writePositions(rowNo)
          If position >= P6CSRRowPtr(rowNo + 1) Then Err.Raise vbObjectError + 3507, "FEM.P6BuildCSRPattern", "CSR行の書込み位置が範囲外です。"
          P6CSRColumnIndex(position) = columnId
          writePositions(rowNo) = position + 1
        End If
NextCsrFillCol:
      Next localCol
      occurrenceIndex = occurrenceNext(occurrenceIndex)
    Loop
  Next rowNo

  ' 行内の列番号を数値配列上で昇順にし、結果を再現可能にする。
  For rowNo = 0 To lastDof
    For position = P6CSRRowPtr(rowNo) + 1 To P6CSRRowPtr(rowNo + 1) - 1
      keyColumn = P6CSRColumnIndex(position)
      keyPosition = position - 1
      Do While keyPosition >= P6CSRRowPtr(rowNo)
        If P6CSRColumnIndex(keyPosition) <= keyColumn Then Exit Do
        P6CSRColumnIndex(keyPosition + 1) = P6CSRColumnIndex(keyPosition)
        keyPosition = keyPosition - 1
      Loop
      P6CSRColumnIndex(keyPosition + 1) = keyColumn
    Next position
    For position = P6CSRRowPtr(rowNo) To P6CSRRowPtr(rowNo + 1) - 1
      If P6CSRColumnIndex(position) = rowNo Then P6CSRDiagonalPosition(rowNo) = position
    Next position
    If P6CSRDiagonalPosition(rowNo) < 0 Then
      Err.Raise vbObjectError + 3503, "FEM.P6BuildCSRPattern", "CSR行に対角成分がありません。行=" & CStr(rowNo)
    End If
  Next rowNo

  ' 数値組立時にDictionaryや列検索を繰り返さないため、局所項からCSR位置を保存する。
  ReDim P6CSRElementPosition(0 To NumberOfElement - 1, 0 To 15, 0 To 15)
  For elementId = 0 To NumberOfElement - 1
    For localRow = 0 To 15
      rowId = Elem(elementId).ElNode(localRow)
      For localCol = 0 To 15
        columnId = Elem(elementId).ElNode(localCol)
        If rowId < 0 Or columnId < 0 Then
          P6CSRElementPosition(elementId, localRow, localCol) = -1
        Else
          position = P6CSRFindColumnPosition(rowId, columnId)
          If position < 0 Then Err.Raise vbObjectError + 3508, "FEM.P6BuildCSRPattern", "要素項からCSR位置を特定できません。"
          P6CSRElementPosition(elementId, localRow, localCol) = position
        End If
      Next localCol
    Next localRow
  Next elementId
  P6CSRPatternReady = True
  P6CSRPatternSignature = P6RCMCacheSignature
  P6BuildCSRPattern = True
End Function

Private Sub P6EnsureCSRValueWorkspace()
  If P6CSRNNZ <= 0 Then Exit Sub
  If P6CSRValueCapacity < P6CSRNNZ Then
    ReDim P6CSROriginalValues(0 To P6CSRNNZ - 1)
    ReDim P6CSRValues(0 To P6CSRNNZ - 1)
    P6CSRValueCapacity = P6CSRNNZ
  End If
End Sub

Private Function P6EnsureCSRScatterList() As Boolean
  Dim elementId As Long, localRow As Long, localCol As Long
  Dim totalCount As Long, writeIndex As Long
  Dim stiffnessValue As Double, cacheBase As Long
  Dim useFlatKCache As Boolean

  P6EnsureCSRScatterList = False
  If P6CSRScatterElementCount = NumberOfElement _
     And P6CSRScatterTangentGeneration = P6TangentGeneration _
     And P6CSRScatterMaterialGeneration = P6MaterialGeneration _
     And P6CSRScatterPatternSignature = P6RCMCacheSignature Then
    P6EnsureCSRScatterList = True
    Exit Function
  End If

  On Error GoTo ScatterBuildFailed
  useFlatKCache = False ' Always assemble the current element tangent, including plastic and joint terms.
  ReDim P6CSRScatterStart(0 To NumberOfElement)
  P6CSRScatterStart(0) = 0
  totalCount = 0
  For elementId = 0 To NumberOfElement - 1
    cacheBase = elementId * 256
    For localRow = 0 To 15
      For localCol = 0 To 15
        If useFlatKCache Then
          stiffnessValue = P6ElasticKCache(cacheBase + localRow * 16 + localCol)
        Else
          stiffnessValue = Elem(elementId).kmat(localRow, localCol)
        End If
        If stiffnessValue <> 0# Then totalCount = totalCount + 1
      Next localCol
    Next localRow
    P6CSRScatterStart(elementId + 1) = totalCount
  Next elementId

  P6CSRScatterCount = totalCount
  If totalCount > 0 Then
    ReDim P6CSRScatterLocalRow(0 To totalCount - 1)
    ReDim P6CSRScatterLocalColumn(0 To totalCount - 1)
    ReDim P6CSRScatterPosition(0 To totalCount - 1)
    ReDim P6CSRScatterValueIndex(0 To totalCount - 1)
    writeIndex = 0
    For elementId = 0 To NumberOfElement - 1
      cacheBase = elementId * 256
      For localRow = 0 To 15
        For localCol = 0 To 15
          If useFlatKCache Then
            stiffnessValue = P6ElasticKCache(cacheBase + localRow * 16 + localCol)
          Else
            stiffnessValue = Elem(elementId).kmat(localRow, localCol)
          End If
          If stiffnessValue <> 0# Then
            P6CSRScatterLocalRow(writeIndex) = localRow
            P6CSRScatterLocalColumn(writeIndex) = localCol
            P6CSRScatterPosition(writeIndex) = P6CSRElementPosition(elementId, localRow, localCol)
            P6CSRScatterValueIndex(writeIndex) = cacheBase + localRow * 16 + localCol
            writeIndex = writeIndex + 1
          End If
        Next localCol
      Next localRow
    Next elementId
  Else
    Erase P6CSRScatterLocalRow
    Erase P6CSRScatterLocalColumn
    Erase P6CSRScatterPosition
    Erase P6CSRScatterValueIndex
  End If
  P6CSRScatterElementCount = NumberOfElement
  P6CSRScatterTangentGeneration = P6TangentGeneration
  P6CSRScatterMaterialGeneration = P6MaterialGeneration
  P6CSRScatterPatternSignature = P6RCMCacheSignature
  P6EnsureCSRScatterList = True
  Exit Function

ScatterBuildFailed:
  P6CSRScatterCount = 0
  P6CSRScatterElementCount = -1
  P6CSRScatterTangentGeneration = -1
  P6CSRScatterMaterialGeneration = -1
  P6CSRScatterPatternSignature = vbNullString
  Erase P6CSRScatterValueIndex
  P6EnsureCSRScatterList = False
  Err.Clear
End Function

Private Function P6AssembleCSRValuesDetailed() As Boolean
  Dim elementId As Long, scatterIndex As Long, position As Long
  Dim localRow As Long, localCol As Long
  Dim useFlatKCache As Boolean

  P6AssembleCSRValuesDetailed = False
  Erase P6CSRDiagonalInverse
  For position = 0 To P6CSRNNZ - 1
    P6CSROriginalValues(position) = 0#
  Next position

  ' CSR再組立時は、生成済み散布表にある有効非ゼロ項だけを加算する。
  useFlatKCache = False ' Always assemble the current element tangent, including plastic and joint terms.
  For elementId = 0 To NumberOfElement - 1
    If Not P3IsElementActive(elementId) Then GoTo NextCsrDetailedElement
    For scatterIndex = P6CSRScatterStart(elementId) To P6CSRScatterStart(elementId + 1) - 1
      position = P6CSRScatterPosition(scatterIndex)
      If useFlatKCache Then
        P6CSROriginalValues(position) = P6CSROriginalValues(position) + P6ElasticKCache(P6CSRScatterValueIndex(scatterIndex))
      Else
        localRow = P6CSRScatterLocalRow(scatterIndex)
        localCol = P6CSRScatterLocalColumn(scatterIndex)
        P6CSROriginalValues(position) = P6CSROriginalValues(position) + Elem(elementId).kmat(localRow, localCol)
      End If
    Next scatterIndex
NextCsrDetailedElement:
  Next elementId
  P6AssembleCSRValuesDetailed = True
End Function

Private Function P6AssembleCSRValuesDirect() As Boolean
  Dim elementId As Long, localRow As Long, localCol As Long
  Dim position As Long, rowId As Long, columnId As Long

  P6AssembleCSRValuesDirect = False
  If P6CSRNNZ <= 0 Then Exit Function
  For position = 0 To P6CSRNNZ - 1
    P6CSROriginalValues(position) = 0#
  Next position

  ' 散布キャッシュに異常がある場合の検証用・復旧用経路。
  ' 要素の局所剛性をCSRの行・列位置へ直接加算し、Kキャッシュの参照を介さない。
  For elementId = 0 To NumberOfElement - 1
    If Not P3IsElementActive(elementId) Then GoTo NextCsrDirectElement
    For localRow = 0 To 15
      rowId = Elem(elementId).ElNode(localRow)
      If rowId < 0 Or rowId > lastDof Then Exit Function
      For localCol = 0 To 15
        columnId = Elem(elementId).ElNode(localCol)
        If columnId < 0 Or columnId > lastDof Then Exit Function
        ' 散布位置キャッシュ自体が壊れているケースも復旧できるよう、
        ' フォールバックではCSR行内を直接検索して位置を再解決する。
        position = P6CSRFindColumnPosition(rowId, columnId)
        If position < 0 Or position >= P6CSRNNZ Then Exit Function
        P6CSROriginalValues(position) = P6CSROriginalValues(position) + Elem(elementId).kmat(localRow, localCol)
      Next localCol
    Next localRow
NextCsrDirectElement:
  Next elementId
  P6AssembleCSRValuesDirect = True
End Function

Private Function P6CheckCSRDiagonalValues(ByRef matrixValues() As Double) As Boolean
  Dim rowNo As Long, diagonalPosition As Long
  Dim diagonalValue As Double

  P6CSRDiagonalZeroCount = 0
  P6CSRFirstZeroDiagonalDof = -1
  P6CSRMinimumDiagonal = 1E+308
  For rowNo = 0 To lastDof
    diagonalPosition = P6CSRDiagonalPosition(rowNo)
    If diagonalPosition < 0 Or diagonalPosition >= P6CSRNNZ Then
      P6CSRDiagonalZeroCount = P6CSRDiagonalZeroCount + 1
      If P6CSRFirstZeroDiagonalDof < 0 Then P6CSRFirstZeroDiagonalDof = rowNo
    Else
      diagonalValue = matrixValues(diagonalPosition)
      If Not P2IsFinite(diagonalValue) Or Abs(diagonalValue) <= 1E-30 Then
        P6CSRDiagonalZeroCount = P6CSRDiagonalZeroCount + 1
        If P6CSRFirstZeroDiagonalDof < 0 Then P6CSRFirstZeroDiagonalDof = rowNo
      ElseIf diagonalValue < P6CSRMinimumDiagonal Then
        P6CSRMinimumDiagonal = diagonalValue
      End If
    End If
  Next rowNo
  If P6CSRMinimumDiagonal = 1E+308 Then P6CSRMinimumDiagonal = 0#
  P6CheckCSRDiagonalValues = (P6CSRDiagonalZeroCount = 0)
End Function

Private Function P6BuildCSRFromElements() As Boolean
  Dim position As Long
  Dim errorNumber As Long, errorDescription As String

  P6BuildCSRFromElements = False
  P6CSRReady = False
  P6CSRILUReady = False
  P6CSRBoundaryApplied = False
  P6CSRDiagonalInverseReady = False
  P6CSRDirectFallbackUsed = False
  P6CSRDiagonalZeroCount = 0
  P6CSRFirstZeroDiagonalDof = -1
  P6CSRMinimumDiagonal = 0#
  On Error GoTo CSRBuildFailed
  If nDof <= 0 Or lastDof < 0 Then Exit Function
  If Not P6CSRPatternReady Or P6CSRPatternSignature <> P6RCMCacheSignature Then
    If Not P6BuildCSRPattern() Then Exit Function
  End If
  If P6CSRNNZ <= 0 Then Exit Function
  P6EnsureCSRValueWorkspace
  If P6CSRPatternReady And P6CSRNumericGeneration = P6TangentGeneration _
     And P6CSRMaterialGeneration = P6MaterialGeneration _
     And P6CSRNNZ > 0 And P6CSRValueCapacity >= P6CSRNNZ Then
    P6CSRReady = True
    P6CSRNumericAssemblyStatus = "reuse"
    P6CSRReuseCount = P6CSRReuseCount + 1
    P6CSRLastAssemblyPath = "reuse_no_detail"
    P6BuildCSRFromElements = True
     Exit Function
  End If
  P6CSRDeviatoricReady = False
  If Not P6EnsureCSRScatterList() Then Exit Function
  If Not P6AssembleCSRValuesDetailed() Then Exit Function
  If Not P6CheckCSRDiagonalValues(P6CSROriginalValues) Then
    If Not P6AssembleCSRValuesDirect() Then Exit Function
    If Not P6CheckCSRDiagonalValues(P6CSROriginalValues) Then
      SetAnalysisFailure RESULT_GLOBAL_SINGULAR, _
        "CSR数値組立後の対角成分が0または極小です。ゼロ対角数=" & CStr(P6CSRDiagonalZeroCount) & _
        "、最初の自由度=" & CStr(P6CSRFirstZeroDiagonalDof), vbObjectError + 3509, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
    P6CSRDirectFallbackUsed = True
    P6CSRLastAssemblyPath = "detailed_direct_fallback"
  End If
  For position = 0 To P6CSRNNZ - 1
    P6CSRValues(position) = P6CSROriginalValues(position)
  Next position
  P6CSRStorageBytes = CDbl(P6CSRNNZ) * 16# + CDbl(P6CSRNNZ + nDof + 1 + nDof) * 4# + CDbl(NumberOfElement) * 256# * 4#
  P6CSRNumericGeneration = P6TangentGeneration
  P6CSRMaterialGeneration = P6MaterialGeneration
  P6CSRNumericAssemblyStatus = "build"
  P6CSRRebuildCount = P6CSRRebuildCount + 1
  P6CSRLastAssemblyPath = "detailed_nonzero"
  P6CSRReady = True
  P6BuildCSRFromElements = True
  Exit Function
CSRBuildFailed:
  errorNumber = Err.Number
  errorDescription = Err.Description
  If errorNumber = 0 Then errorNumber = vbObjectError + 3504
  If Len(errorDescription) = 0 Then errorDescription = "CSR疎行列の組立で不明なエラーが発生しました。"
  SetAnalysisFailure RESULT_CAPACITY_ERROR, errorDescription, errorNumber, -1, -1, CurrentIncrement, CurrentIteration
End Function

Private Function P6BuildCSRBoundaryExtraction() As Boolean
  Dim rowNo As Long, position As Long, columnId As Long
  Dim zeroCount As Long, diagonalCount As Long, rhsCount As Long
  P6BuildCSRBoundaryExtraction = False
  If Not P6CSRReady Then Exit Function
  If P6CSRBoundaryConstraintVersion = P6CSRConstraintVersion Then
    P6BuildCSRBoundaryExtraction = True
    Exit Function
  End If

  For rowNo = 0 To lastDof
    If NodeCond(rowNo) <> 0 Then diagonalCount = diagonalCount + 1
    For position = P6CSRRowPtr(rowNo) To P6CSRRowPtr(rowNo + 1) - 1
      columnId = P6CSRColumnIndex(position)
      If NodeCond(rowNo) <> 0 Or NodeCond(columnId) <> 0 Then zeroCount = zeroCount + 1
      If NodeCond(rowNo) = 0 And NodeCond(columnId) <> 0 Then rhsCount = rhsCount + 1
    Next position
  Next rowNo

  P6CSRBoundaryZeroCount = zeroCount
  P6CSRBoundaryDiagonalCount = diagonalCount
  P6CSRBoundaryRHSCount = rhsCount
  If zeroCount > 0 Then ReDim P6CSRBoundaryZeroPosition(0 To zeroCount - 1) Else Erase P6CSRBoundaryZeroPosition
  If diagonalCount > 0 Then ReDim P6CSRBoundaryDiagonalPosition(0 To diagonalCount - 1) Else Erase P6CSRBoundaryDiagonalPosition
  If rhsCount > 0 Then
    ReDim P6CSRBoundaryRHSPosition(0 To rhsCount - 1)
    ReDim P6CSRBoundaryRHSRow(0 To rhsCount - 1)
    ReDim P6CSRBoundaryRHSColumn(0 To rhsCount - 1)
  Else
    Erase P6CSRBoundaryRHSPosition
    Erase P6CSRBoundaryRHSRow
    Erase P6CSRBoundaryRHSColumn
  End If

  zeroCount = 0: diagonalCount = 0: rhsCount = 0
  For rowNo = 0 To lastDof
    If NodeCond(rowNo) <> 0 Then
      P6CSRBoundaryDiagonalPosition(diagonalCount) = P6CSRDiagonalPosition(rowNo)
      diagonalCount = diagonalCount + 1
    End If
    For position = P6CSRRowPtr(rowNo) To P6CSRRowPtr(rowNo + 1) - 1
      columnId = P6CSRColumnIndex(position)
      If NodeCond(rowNo) <> 0 Or NodeCond(columnId) <> 0 Then
        P6CSRBoundaryZeroPosition(zeroCount) = position
        zeroCount = zeroCount + 1
      End If
      If NodeCond(rowNo) = 0 And NodeCond(columnId) <> 0 Then
        P6CSRBoundaryRHSPosition(rhsCount) = position
        P6CSRBoundaryRHSRow(rhsCount) = rowNo
        P6CSRBoundaryRHSColumn(rhsCount) = columnId
        rhsCount = rhsCount + 1
      End If
    Next position
  Next rowNo
  P6CSRBoundaryConstraintVersion = P6CSRConstraintVersion
  P6BuildCSRBoundaryExtraction = True
End Function

Private Function ApplyCSRBoundaryToMatrix() As Boolean
  Dim i As Long, position As Long
  ApplyCSRBoundaryToMatrix = False
  If Not P6CSRReady Then Exit Function
  If Not P6BuildCSRBoundaryExtraction() Then Exit Function
  P6EnsureCSRValueWorkspace
  For position = 0 To P6CSRNNZ - 1
    P6CSRValues(position) = P6CSROriginalValues(position)
  Next position
  P6CSRDiagonalInverseReady = False
  P6CSRILUReady = False
  For i = 0 To P6CSRBoundaryZeroCount - 1
    position = P6CSRBoundaryZeroPosition(i)
    P6CSRValues(position) = 0#
  Next i
  For i = 0 To P6CSRBoundaryDiagonalCount - 1
    position = P6CSRBoundaryDiagonalPosition(i)
    If position < 0 Then Exit Function
    P6CSRValues(position) = 1#
  Next i
  If Not P6CheckCSRDiagonalValues(P6CSRValues) Then
    SetAnalysisFailure RESULT_GLOBAL_SINGULAR, _
      "境界条件適用後のCSR対角成分が0または極小です。ゼロ対角数=" & CStr(P6CSRDiagonalZeroCount) & _
      "、最初の自由度=" & CStr(P6CSRFirstZeroDiagonalDof), vbObjectError + 3510, -1, -1, CurrentIncrement, CurrentIteration
    Exit Function
  End If
  ApplyCSRBoundaryToMatrix = True
End Function

Private Sub ApplyCSRBoundaryRHS()
  Dim i As Long
  If P6CSRBoundaryConstraintVersion <> P6CSRConstraintVersion Then
    If Not P6BuildCSRBoundaryExtraction() Then Exit Sub
  End If
  For i = 0 To P6CSRBoundaryRHSCount - 1
    Force(P6CSRBoundaryRHSRow(i)) = Force(P6CSRBoundaryRHSRow(i)) - _
        P6CSROriginalValues(P6CSRBoundaryRHSPosition(i)) * Disp(P6CSRBoundaryRHSColumn(i))
  Next i
End Sub

Private Sub P6RecordPivot(ByVal pivot As Double, ByVal constrained As Boolean)
  Dim absPivot As Double, nearZero As Double
  If constrained Then Exit Sub
  absPivot = Abs(pivot)
  If absPivot < P6PivotMin Then P6PivotMin = absPivot
  If absPivot < P6PivotMinTrial Then P6PivotMinTrial = absPivot
  If absPivot > P6PivotMax Then P6PivotMax = absPivot
  If absPivot > P6PivotMaxTrial Then P6PivotMaxTrial = absPivot
  If pivot < 0# Then
    P6PivotNegCount = P6PivotNegCount + 1
    P6PivotNegCountTrial = P6PivotNegCountTrial + 1
  End If
  nearZero = 0.00000001 * P6PivotScale
  If nearZero < 0.000000000001 Then nearZero = 0.000000000001
  If absPivot <= nearZero Then
    P6PivotNearZeroCount = P6PivotNearZeroCount + 1
    P6PivotNearZeroCountTrial = P6PivotNearZeroCountTrial + 1
  End If
End Sub

Private Function P6BuildBandFactorization() As Boolean
  Dim k As Long, i As Long, j As Long, it As Long, jt As Long
  Dim m As Long, n2 As Long, im1 As Long, workN As Long
  Dim pivot As Double, pivotTolerance As Double, matrixScale As Double
  Dim t0 As Double, alreadySized As Boolean
  Dim benchMs As Double, tBench As Double, FactorMs As Double

  t0 = Timer
  P6BuildBandFactorization = False
  P6FactorReady = False
  MatrixFactored = False
  n2 = nDof
  If n2 <= 0 Then
    SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "自由度数が0です。", vbObjectError + 3301, -1, -1, CurrentIncrement, CurrentIteration
    P6ProfFactorMs = P6ProfFactorMs + P6ElapsedMs(t0)
    Exit Function
  End If
  If lastDof < 0 Then
    SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "自由度の最終番号が負です。", vbObjectError + 3301, -1, -1, CurrentIncrement, CurrentIteration
    P6ProfFactorMs = P6ProfFactorMs + P6ElapsedMs(t0)
    Exit Function
  End If
  If BandWidth < 0 Then
    SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "バンド幅が負です。", vbObjectError + 3302, -1, -1, CurrentIncrement, CurrentIteration
    P6ProfFactorMs = P6ProfFactorMs + P6ElapsedMs(t0)
    Exit Function
  End If
  If BandWidth > lastDof Then
    SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "バンド幅が自由度範囲外です。", vbObjectError + 3302, -1, -1, CurrentIncrement, CurrentIteration
    P6ProfFactorMs = P6ProfFactorMs + P6ElapsedMs(t0)
    Exit Function
  End If

  alreadySized = P6BandArrayReady(P6FactoredBand)
  If Not alreadySized Then
    If P6MixedUP Then
      P6CopyBand TotalMat, P6FactoredBand
    Else
      P6CopyBand OriginalMat, P6FactoredBand
    End If
    For i = 0 To lastDof
      If NodeCond(i) <> 0 Then P6FactoredBand(i, 0) = 1E+30
    Next i
  End If
  workN = BandWidth
  If workN < 1 Then workN = 1
  If P6BandWorkN < workN Then
    ReDim P6BandWorkAkj(0 To workN)
    ReDim P6BandWorkAik(0 To workN)
    P6BandWorkN = workN
  End If
  matrixScale = 0#
  For i = 0 To lastDof
    If NodeCond(i) = 0 Then
      If Abs(P6FactoredBand(i, 0)) > matrixScale Then matrixScale = Abs(P6FactoredBand(i, 0))
    End If
  Next i
  If matrixScale < 1# Then matrixScale = 1#
  P6PivotScale = matrixScale
  pivotTolerance = 0.000000000001 * matrixScale
  For i = 0 To lastDof
    If NodeCond(i) = 0 Then
      If P6FactoredBand(i, 0) <= pivotTolerance Then
        P6RecordPivot P6FactoredBand(i, 0), False
        SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "自由自由度の対角ピボットが正でないか小さすぎます。自由度=" & CStr(i), vbObjectError + 3303, -1, -1, CurrentIncrement, CurrentIteration
        P6ProfFactorMs = P6ProfFactorMs + P6ElapsedMs(t0)
        Exit Function
      End If
    End If
  Next i
  If P6LdltShadowDue() Then
    P6SnapshotLdltHold
  Else
    P6LdltSnapReady = False
  End If

  benchMs = 0#
  If P6BandBenchDue() Then
    P6BandBenchDone = True
    tBench = Timer
    P6ScheduleBandBench
    benchMs = P6ElapsedMs(tBench)
  End If

  If P6BandKernelIsNew() Then
    P6BuildBandFactorization = P6FactorPackedProduction(n2, pivotTolerance, t0, benchMs)
    Exit Function
  End If

  For k = 0 To n2 - 1
    m = BandWidth + 1
    If (k + BandWidth + 1) > n2 Then m = n2 - k
    pivot = P6FactoredBand(k, 0)
    If NodeCond(k) = 0 Then P6RecordPivot pivot, False
    If Abs(pivot) <= pivotTolerance Then
      SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "バンド分解中にゼロまたは小さすぎるピボットが発生しました。自由度=" & CStr(k), vbObjectError + 3304, -1, -1, CurrentIncrement, CurrentIteration
      FactorMs = P6ElapsedMs(t0) - benchMs
      If FactorMs < 0# Then FactorMs = 0#
      P6ProfFactorMs = P6ProfFactorMs + FactorMs
      Exit Function
    End If
    For j = 1 To m - 1
      P6BandWorkAkj(j) = P6FactoredBand(k, j) / pivot
      P6BandWorkAik(j) = P6FactoredBand(k, j)
    Next j
    For i = 1 To m - 1
      it = i + k
      im1 = i
      For j = i To m - 1
        jt = j - im1
        P6FactoredBand(it, jt) = P6FactoredBand(it, jt) - P6BandWorkAik(i) * P6BandWorkAkj(j)
      Next j
    Next i
  Next k
  P6FactorPivotTolerance = pivotTolerance
  P6FactorizationCount = P6FactorizationCount + 1
  P6RememberFactorKeys
  P6FactorReady = True
  P6BuildBandFactorization = True
  FactorMs = P6ElapsedMs(t0) - benchMs
  If FactorMs < 0# Then FactorMs = 0#
  P6ProfFactorMs = P6ProfFactorMs + FactorMs
End Function

Private Function P6SolveBandRHS() As Boolean
  Dim k As Long, i As Long, j As Long, it As Long, jt As Long
  Dim m As Long, n2 As Long, s As Double
  Dim t0 As Double
  Dim errNum As Long, errSrc As String, errDesc As String

  t0 = Timer
  P6SolveBandRHS = False
  On Error GoTo SolveFail
  n2 = nDof
  If n2 <= 0 Then
    SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "自由度数が0です。", vbObjectError + 3301, -1, -1, CurrentIncrement, CurrentIteration
    GoTo SolveDone
  End If
  If Not P6FactorReady Then
    SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "帯域分解が準備されていません。", vbObjectError + 3301, -1, -1, CurrentIncrement, CurrentIteration
    GoTo SolveDone
  End If
  If P6BandKernelIsNew() Then
    If Not P6BandSolve1D(P6Packed, P6PackRowOff, Force, Disp, n2, BandWidth, P6FactorPivotTolerance) Then
      k = P6BandSolveFailK
      If k < 0 Then k = 0
      SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "パックド帯域の求解でゼロまたは小さすぎるピボットです。自由度=" & CStr(k), vbObjectError + 3305, -1, -1, CurrentIncrement, CurrentIteration
      GoTo SolveDone
    End If
    P6SolveBandRHS = True
    GoTo SolveDone
  End If

  For k = 0 To n2 - 2
    If Abs(P6FactoredBand(k, 0)) <= P6FactorPivotTolerance Then
      SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "右辺前進消去中にゼロまたは小さすぎるピボットが発生しました。自由度=" & CStr(k), vbObjectError + 3305, -1, -1, CurrentIncrement, CurrentIteration
      GoTo SolveDone
    End If
    Force(k) = Force(k) / P6FactoredBand(k, 0)
    m = BandWidth + 1
    If (k + BandWidth + 1) > n2 Then m = n2 - k
    For i = 1 To m - 1
      it = i + k
      Force(it) = Force(it) - P6FactoredBand(k, i) * Force(k)
    Next i
  Next k
  If Abs(P6FactoredBand(n2 - 1, 0)) <= P6FactorPivotTolerance Then
    SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "後退代入中にゼロまたは小さすぎる最終ピボットが発生しました。", vbObjectError + 3306, -1, -1, CurrentIncrement, CurrentIteration
    GoTo SolveDone
  End If
  Disp(n2 - 1) = Force(n2 - 1) / P6FactoredBand(n2 - 1, 0)
  Force(n2 - 1) = 0#
  For k = n2 - 2 To 0 Step -1
    s = 0#
    m = BandWidth + 1
    If (k + BandWidth + 1) > n2 Then m = n2 - k
    For j = 1 To m - 1
      jt = j + k
      s = s + P6FactoredBand(k, j) * Disp(jt)
    Next j
    Disp(k) = Force(k) - s / P6FactoredBand(k, 0)
  Next k
  P6SolveBandRHS = True
  GoTo SolveDone
SolveFail:
  errNum = Err.Number
  errSrc = Err.source
  errDesc = Err.Description
  Resume SolveDone
SolveDone:
  On Error GoTo 0
  P6ProfSolveMs = P6ProfSolveMs + P6ElapsedMs(t0)
  If errNum <> 0 Then Err.Raise errNum, errSrc, errDesc
End Function

' Production factor entry. Reuse must not call this.
Private Function BandLUFactor() As Boolean
  BandLUFactor = P6BuildBandFactorization()
End Function

' Production solve entry. Does not factor.
Private Function BandLUSolve() As Boolean
  BandLUSolve = P6SolveBandRHS()
End Function

' NEW = packed 1D. OLD or BAND_LU_OLD keeps the 2D loop. Missing key is NEW.
Public Function P6BandKernelName() As String
  Dim keyText As String
  If Not P6KernelCached Then
    keyText = UCase$(Trim$(P6ReadTextSetting("BAND_LU_KERNEL", "NEW")))
    If keyText = "OLD" Or keyText = "BAND_LU_OLD" Then
      P6KernelNew = False
    Else
      P6KernelNew = True
    End If
    P6KernelCached = True
  End If
  If P6KernelNew Then
    P6BandKernelName = "NEW"
  Else
    P6BandKernelName = "OLD"
  End If
End Function

Private Function P6BandKernelIsNew() As Boolean
  P6BandKernelIsNew = (P6BandKernelName() = "NEW")
End Function

Private Sub P6EnsurePackStorage(ByVal n As Long, ByVal bw As Long)
  Dim stride As Long
  stride = bw + 1
  If n < 1 Or bw < 0 Then Exit Sub
  If P6PackN = n And P6PackBw = bw Then Exit Sub
  ReDim P6Packed(0 To n * stride - 1)
  ReDim P6PackRowOff(0 To n - 1)
  P6BandBuildRowOff P6PackRowOff, n, stride
  P6PackN = n
  P6PackBw = bw
End Sub

' Factor the already penalized unfactored band. Does not run the 2D loop.
Private Function P6FactorPackedProduction(ByVal n2 As Long, ByVal pivotTolerance As Double, ByVal t0 As Double, ByVal benchMs As Double) As Boolean
  Dim k As Long
  Dim FactorMs As Double
  P6FactorPackedProduction = False
  P6EnsurePackStorage n2, BandWidth
  P6PackBand P6FactoredBand, P6Packed, P6PackRowOff, n2, BandWidth
  If Not P6BandFactor1D(P6Packed, P6PackRowOff, P6BandWorkAkj, P6BandWorkAik, n2, BandWidth, pivotTolerance) Then
    k = P6BandFactorFailK
    If k < 0 Then k = 0
    If k >= 0 And k <= lastDof Then
      If NodeCond(k) = 0 Then P6RecordPivot P6Packed(P6PackRowOff(k)), False
    End If
    SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "パックドバンド分解中にゼロまたは小さすぎるピボットが発生しました。自由度=" & CStr(k), vbObjectError + 3304, -1, -1, CurrentIncrement, CurrentIteration
    FactorMs = P6ElapsedMs(t0) - benchMs
    If FactorMs < 0# Then FactorMs = 0#
    P6ProfFactorMs = P6ProfFactorMs + FactorMs
    Exit Function
  End If
  For k = 0 To n2 - 1
    If NodeCond(k) = 0 Then P6RecordPivot P6Packed(P6PackRowOff(k)), False
  Next k
  P6FactorPivotTolerance = pivotTolerance
  P6FactorizationCount = P6FactorizationCount + 1
  P6RememberFactorKeys
  P6FactorReady = True
  P6FactorPackedProduction = True
  FactorMs = P6ElapsedMs(t0) - benchMs
  If FactorMs < 0# Then FactorMs = 0#
  P6ProfFactorMs = P6ProfFactorMs + FactorMs
End Function

Public Sub P6ResetBandBench()
  P6BandBenchDone = False
  P6BandBenchCached = False
  P6BandBenchOn = False
  P6BandBenchPending = False
  P6BandBenchN = 0
  P6BandBenchBw = 0
  P6BandBenchOnly = False
  P6KernelCached = False
  Erase P6BandBenchSnap
  Erase P6BandBenchFree
End Sub

Private Sub P6ScheduleBandBench()
  Dim i As Long
  P6BandBenchPending = False
  If nDof < 1 Or BandWidth < 0 Or nDof <> lastDof + 1 Then
    P6BandBenchEmit 0, BandWidth, 0, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "DIM"
    Exit Sub
  End If
  If CDbl(nDof) * CDbl(BandWidth + 1) > 8000000# Then
    P6BandBenchEmit nDof, BandWidth, 0, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "SKIP_SIZE"
    Exit Sub
  End If
  P6BandBenchN = nDof
  P6BandBenchBw = BandWidth
  ReDim P6BandBenchSnap(0 To lastDof, 0 To BandWidth)
  P6BandBenchSnap = P6FactoredBand
  ReDim P6BandBenchFree(0 To lastDof)
  For i = 0 To lastDof
    P6BandBenchFree(i) = (NodeCond(i) = 0)
  Next i
  P6BandBenchPending = True
End Sub

Public Sub P6FlushBandBench()
  If Not P6BandBenchPending Then Exit Sub
  P6BandBenchPending = False
  P6RunBandLUBench
  Erase P6BandBenchSnap
  Erase P6BandBenchFree
End Sub

Private Function P6BandBenchDue() As Boolean
  P6BandBenchDue = False
  If P6BandBenchDone Then Exit Function
  If Not P6BandBenchCached Then
    P6BandBenchOn = (P6ReadSetting("BAND_LU_BENCH", 0#) >= 1#)
    P6BandBenchOnly = P6BandBenchOn
    P6BandBenchCached = True
  End If
  P6BandBenchDue = P6BandBenchOn
End Function

Private Sub P6BandBuildRowOff(ByRef rowOff() As Long, ByVal n As Long, ByVal stride As Long)
  Dim i As Long
  For i = 0 To n - 1
    rowOff(i) = i * stride
  Next i
End Sub

Private Sub P6PackBand(ByRef mat() As Double, ByRef packed() As Double, ByRef rowOff() As Long, ByVal n As Long, ByVal bw As Long)
  Dim i As Long, j As Long, base As Long, nBand As Long
  nBand = bw
  For i = 0 To n - 1
    base = rowOff(i)
    packed(base) = mat(i, 0)
    For j = 1 To nBand
      packed(base + j) = mat(i, j)
    Next j
  Next i
End Sub

Private Function P6BandFactor2D(ByRef mat() As Double, ByRef akj() As Double, ByRef aik() As Double, ByVal n As Long, ByVal bw As Long, ByVal tol As Double) As Boolean
  Dim k As Long, i As Long, j As Long, it As Long, jt As Long
  Dim m As Long, im1 As Long, nBand As Long
  Dim pivot As Double
  P6BandFactor2D = False
  nBand = bw
  For k = 0 To n - 1
    m = nBand + 1
    If (k + nBand + 1) > n Then m = n - k
    pivot = mat(k, 0)
    If Abs(pivot) <= tol Then Exit Function
    For j = 1 To m - 1
      akj(j) = mat(k, j) / pivot
      aik(j) = mat(k, j)
    Next j
    For i = 1 To m - 1
      it = i + k
      im1 = i
      For j = i To m - 1
        jt = j - im1
        mat(it, jt) = mat(it, jt) - aik(i) * akj(j)
      Next j
    Next i
  Next k
  P6BandFactor2D = True
End Function

Private Function P6BandFactor1D(ByRef packed() As Double, ByRef rowOff() As Long, ByRef akj() As Double, ByRef aik() As Double, ByVal n As Long, ByVal bw As Long, ByVal tol As Double) As Boolean
  Dim k As Long, i As Long, j As Long, t As Long
  Dim m As Long, m1 As Long, lim As Long, nBand As Long
  Dim rowK As Long, rowIt As Long, idx As Long
  Dim pivot As Double, aikv As Double
  P6BandFactor1D = False
  P6BandFactorFailK = -1
  nBand = bw
  For k = 0 To n - 1
    m = nBand + 1
    If (k + nBand + 1) > n Then m = n - k
    m1 = m - 1
    rowK = rowOff(k)
    pivot = packed(rowK)
    If Abs(pivot) <= tol Then
      P6BandFactorFailK = k
      Exit Function
    End If
    For j = 1 To m1
      akj(j) = packed(rowK + j) / pivot
      aik(j) = packed(rowK + j)
    Next j
    For i = 1 To m1
      rowIt = rowOff(k + i)
      aikv = aik(i)
      lim = m1 - i
      For t = 0 To lim
        idx = rowIt + t
        packed(idx) = packed(idx) - aikv * akj(t + i)
      Next t
    Next i
  Next k
  P6BandFactor1D = True
End Function

Private Function P6BandSolve2D(ByRef mat() As Double, ByRef rhs() As Double, ByRef x() As Double, ByVal n As Long, ByVal bw As Long, ByVal tol As Double) As Boolean
  Dim k As Long, i As Long, j As Long, it As Long, jt As Long
  Dim m As Long, nBand As Long
  Dim s As Double
  P6BandSolve2D = False
  If n <= 0 Then Exit Function
  nBand = bw
  For k = 0 To n - 2
    If Abs(mat(k, 0)) <= tol Then Exit Function
    rhs(k) = rhs(k) / mat(k, 0)
    m = nBand + 1
    If (k + nBand + 1) > n Then m = n - k
    For i = 1 To m - 1
      it = i + k
      rhs(it) = rhs(it) - mat(k, i) * rhs(k)
    Next i
  Next k
  If Abs(mat(n - 1, 0)) <= tol Then Exit Function
  x(n - 1) = rhs(n - 1) / mat(n - 1, 0)
  rhs(n - 1) = 0#
  For k = n - 2 To 0 Step -1
    s = 0#
    m = nBand + 1
    If (k + nBand + 1) > n Then m = n - k
    For j = 1 To m - 1
      jt = j + k
      s = s + mat(k, j) * x(jt)
    Next j
    x(k) = rhs(k) - s / mat(k, 0)
  Next k
  P6BandSolve2D = True
End Function

Private Function P6BandSolve1D(ByRef packed() As Double, ByRef rowOff() As Long, ByRef rhs() As Double, ByRef x() As Double, ByVal n As Long, ByVal bw As Long, ByVal tol As Double) As Boolean
  Dim k As Long, i As Long, j As Long, it As Long, jt As Long
  Dim m As Long, nBand As Long, rowK As Long
  Dim s As Double, piv As Double
  P6BandSolve1D = False
  P6BandSolveFailK = -1
  If n <= 0 Then Exit Function
  nBand = bw
  For k = 0 To n - 2
    rowK = rowOff(k)
    piv = packed(rowK)
    If Abs(piv) <= tol Then
      P6BandSolveFailK = k
      Exit Function
    End If
    rhs(k) = rhs(k) / piv
    m = nBand + 1
    If (k + nBand + 1) > n Then m = n - k
    For i = 1 To m - 1
      it = i + k
      rhs(it) = rhs(it) - packed(rowK + i) * rhs(k)
    Next i
  Next k
  rowK = rowOff(n - 1)
  piv = packed(rowK)
  If Abs(piv) <= tol Then
    P6BandSolveFailK = n - 1
    Exit Function
  End If
  x(n - 1) = rhs(n - 1) / piv
  rhs(n - 1) = 0#
  For k = n - 2 To 0 Step -1
    rowK = rowOff(k)
    s = 0#
    m = nBand + 1
    If (k + nBand + 1) > n Then m = n - k
    For j = 1 To m - 1
      jt = j + k
      s = s + packed(rowK + j) * x(jt)
    Next j
    x(k) = rhs(k) - s / packed(rowK)
  Next k
  P6BandSolve1D = True
End Function

Private Sub P6BandApply(ByRef mat() As Double, ByRef x() As Double, ByRef y() As Double, ByVal n As Long, ByVal bw As Long)
  Dim i As Long, j As Long, col As Long, m As Long, nBand As Long
  Dim s As Double
  nBand = bw
  For i = 0 To n - 1
    s = mat(i, 0) * x(i)
    m = nBand
    If i + m > n - 1 Then m = n - 1 - i
    For j = 1 To m
      s = s + mat(i, j) * x(i + j)
    Next j
    For j = 1 To nBand
      col = i - j
      If col < 0 Then Exit For
      s = s + mat(col, j) * x(col)
    Next j
    y(i) = s
  Next i
End Sub

Private Function P6BandVecRel(ByRef a() As Double, ByRef b() As Double, ByVal n As Long) As Double
  Dim i As Long
  Dim num As Double, den As Double, d As Double
  num = 0#
  den = 0#
  For i = 0 To n - 1
    d = a(i) - b(i)
    num = num + d * d
    den = den + b(i) * b(i)
  Next i
  P6BandVecRel = Sqr(num) / (Sqr(den) + 1E-30)
End Function

Private Function P6BandFreeRel(ByRef resid() As Double, ByRef rhs() As Double, ByVal n As Long) As Double
  Dim i As Long
  Dim num As Double, den As Double
  num = 0#
  den = 0#
  For i = 0 To n - 1
    If P6BandBenchFree(i) Then
      num = num + resid(i) * resid(i)
      den = den + rhs(i) * rhs(i)
    End If
  Next i
  P6BandFreeRel = Sqr(num) / (Sqr(den) + 1E-30)
End Function

Private Function P6BandPlain(ByVal v As Double, ByVal ok As Boolean) As String
  If Not ok Then
    P6BandPlain = "NA"
  Else
    P6BandPlain = Format$(v, "0.000")
  End If
End Function

Private Function P6BandSci(ByVal v As Double, ByVal ok As Boolean) As String
  If Not ok Then
    P6BandSci = "NA"
  Else
    P6BandSci = Format$(v, "0.000E+00")
  End If
End Function

Private Sub P6BandBenchEmit(ByVal n As Long, ByVal bw As Long, ByVal reps As Long, ByVal tOld As String, ByVal tNew As String, ByVal tOldFac As String, ByVal tNewFac As String, ByVal tPack As String, ByVal speedupLoop As String, ByVal speedupNet As String, ByVal exText As String, ByVal erOld As String, ByVal erNew As String, ByVal gateText As String, ByVal reasonText As String)
  Dim lineText As String, detailText As String
  lineText = FEM_BUILD_STAMP
  lineText = lineText & "," & CStr(n)
  lineText = lineText & "," & CStr(bw)
  lineText = lineText & "," & CStr(reps)
  lineText = lineText & "," & tOld
  lineText = lineText & "," & tNew
  lineText = lineText & "," & tOldFac
  lineText = lineText & "," & tNewFac
  lineText = lineText & "," & tPack
  lineText = lineText & "," & speedupLoop
  lineText = lineText & "," & speedupNet
  lineText = lineText & "," & exText
  lineText = lineText & "," & erOld
  lineText = lineText & "," & erNew
  lineText = lineText & "," & gateText
  lineText = lineText & "," & reasonText
  FEMIoAppendBenchRow lineText
  detailText = "n=" & CStr(n)
  detailText = detailText & ";bw=" & CStr(bw)
  detailText = detailText & ";reps=" & CStr(reps)
  detailText = detailText & ";old_factor_ms=" & tOld
  detailText = detailText & ";new_factor_ms=" & tNew
  detailText = detailText & ";old_solve_ms=" & tOldFac
  detailText = detailText & ";new_solve_ms=" & tNewFac
  detailText = detailText & ";pack_ms=" & tPack
  detailText = detailText & ";speedup_factor=" & speedupLoop
  detailText = detailText & ";speedup_net=" & speedupNet
  detailText = detailText & ";x_error=" & exText
  detailText = detailText & ";residual_old=" & erOld
  detailText = detailText & ";residual_new=" & erNew
  detailText = detailText & ";gate=" & gateText
  detailText = detailText & ";reason=" & reasonText
  detailText = detailText & ";switched=0"
  P6SolverEvent "BAND_LU_BENCH", detailText
  FEMAppendRunLog "BAND_LU_BENCH", gateText, detailText
End Sub

Private Sub P6RunBandLUBench()
  Dim n As Long, bw As Long, reps As Long, rep As Long, stride As Long, i As Long
  Dim tol As Double, matrixScale As Double, entries As Double
  Dim t0 As Double, tOldAll As Double, tOldCopy As Double, tOldFac As Double
  Dim tNewAll As Double, tPack As Double, tNewFac As Double
  Dim tSolveOld As Double, tSolveNew As Double
  Dim perOld As Double, perNew As Double, perPack As Double, perSolOld As Double, perSolNew As Double
  Dim speedupLoop As Double, speedupNet As Double, ex As Double, erOld As Double, erNew As Double
  Dim okOld As Boolean, okNew As Boolean, timed As Boolean, benchErrLogged As Boolean
  Dim snap() As Double, work2d() As Double, packed() As Double
  Dim akj() As Double, aik() As Double
  Dim rowOff() As Long
  Dim rhs() As Double, bOld() As Double, bNew() As Double
  Dim xOld() As Double, xNew() As Double, y() As Double
  Dim gateText As String, reasonText As String
  On Error GoTo BenchFail
  benchErrLogged = False
  n = P6BandBenchN
  bw = P6BandBenchBw
  If n < 1 Or bw < 0 Then
    P6BandBenchEmit 0, bw, 0, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "DIM"
    Exit Sub
  End If
  stride = bw + 1
  entries = CDbl(n) * CDbl(stride)
  If entries > 8000000# Then
    P6BandBenchEmit n, bw, 0, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "SKIP_SIZE"
    Exit Sub
  End If
  reps = CLng(P6ReadSetting("BAND_LU_BENCH_N", 20#))
  If reps < 1 Then reps = 1
  If reps > 30 Then reps = 30
  ReDim snap(0 To n - 1, 0 To bw)
  snap = P6BandBenchSnap
  matrixScale = 0#
  For i = 0 To n - 1
    If P6BandBenchFree(i) Then
      If Abs(snap(i, 0)) > matrixScale Then matrixScale = Abs(snap(i, 0))
    End If
  Next i
  If matrixScale < 1# Then matrixScale = 1#
  tol = 0.000000000001 * matrixScale
  ReDim akj(0 To bw)
  ReDim aik(0 To bw)
  ReDim work2d(0 To n - 1, 0 To bw)
  ReDim rowOff(0 To n - 1)
  P6BandBuildRowOff rowOff, n, stride
  ReDim packed(0 To n * stride - 1)
  For rep = 1 To 3
    work2d = snap
    okOld = P6BandFactor2D(work2d, akj, aik, n, bw, tol)
    If Not okOld Then
      P6BandBenchEmit n, bw, reps, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "FACTOR_FAIL"
      Exit Sub
    End If
    P6PackBand snap, packed, rowOff, n, bw
    okNew = P6BandFactor1D(packed, rowOff, akj, aik, n, bw, tol)
    If Not okNew Then
      P6BandBenchEmit n, bw, reps, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "FACTOR_FAIL"
      Exit Sub
    End If
  Next rep
  t0 = Timer
  For rep = 1 To reps
    work2d = snap
    okOld = P6BandFactor2D(work2d, akj, aik, n, bw, tol)
    If Not okOld Then Exit For
  Next rep
  tOldAll = P6ElapsedMs(t0)
  If Not okOld Then
    P6BandBenchEmit n, bw, reps, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "FACTOR_FAIL"
    Exit Sub
  End If
  t0 = Timer
  For rep = 1 To reps
    work2d = snap
  Next rep
  tOldCopy = P6ElapsedMs(t0)
  tOldFac = tOldAll - tOldCopy
  If tOldFac < 0# Then tOldFac = 0#
  t0 = Timer
  For rep = 1 To reps
    P6PackBand snap, packed, rowOff, n, bw
    okNew = P6BandFactor1D(packed, rowOff, akj, aik, n, bw, tol)
    If Not okNew Then Exit For
  Next rep
  tNewAll = P6ElapsedMs(t0)
  If Not okNew Then
    P6BandBenchEmit n, bw, reps, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "FACTOR_FAIL"
    Exit Sub
  End If
  t0 = Timer
  For rep = 1 To reps
    P6PackBand snap, packed, rowOff, n, bw
  Next rep
  tPack = P6ElapsedMs(t0)
  tNewFac = tNewAll - tPack
  If tNewFac < 0# Then tNewFac = 0#
  work2d = snap
  okOld = P6BandFactor2D(work2d, akj, aik, n, bw, tol)
  P6PackBand snap, packed, rowOff, n, bw
  okNew = P6BandFactor1D(packed, rowOff, akj, aik, n, bw, tol)
  If (Not okOld) Or (Not okNew) Then
    P6BandBenchEmit n, bw, reps, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "FACTOR_FAIL"
    Exit Sub
  End If
  ReDim rhs(0 To n - 1)
  ReDim bOld(0 To n - 1)
  ReDim bNew(0 To n - 1)
  ReDim xOld(0 To n - 1)
  ReDim xNew(0 To n - 1)
  ReDim y(0 To n - 1)
  For i = 0 To n - 1
    If P6BandBenchFree(i) Then
      rhs(i) = 1#
    Else
      rhs(i) = 0#
    End If
    bOld(i) = rhs(i)
    bNew(i) = rhs(i)
  Next i
  okOld = P6BandSolve2D(work2d, bOld, xOld, n, bw, tol)
  okNew = P6BandSolve1D(packed, rowOff, bNew, xNew, n, bw, tol)
  If (Not okOld) Or (Not okNew) Then
    P6BandBenchEmit n, bw, reps, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "SOLVE_FAIL"
    Exit Sub
  End If
  ex = P6BandVecRel(xNew, xOld, n)
  P6BandApply snap, xOld, y, n, bw
  For i = 0 To n - 1
    y(i) = y(i) - rhs(i)
  Next i
  erOld = P6BandFreeRel(y, rhs, n)
  P6BandApply snap, xNew, y, n, bw
  For i = 0 To n - 1
    y(i) = y(i) - rhs(i)
  Next i
  erNew = P6BandFreeRel(y, rhs, n)
  t0 = Timer
  For rep = 1 To reps
    For i = 0 To n - 1
      bOld(i) = rhs(i)
    Next i
    okOld = P6BandSolve2D(work2d, bOld, xOld, n, bw, tol)
    If Not okOld Then Exit For
  Next rep
  tSolveOld = P6ElapsedMs(t0)
  t0 = Timer
  For rep = 1 To reps
    For i = 0 To n - 1
      bNew(i) = rhs(i)
    Next i
    okNew = P6BandSolve1D(packed, rowOff, bNew, xNew, n, bw, tol)
    If Not okNew Then Exit For
  Next rep
  tSolveNew = P6ElapsedMs(t0)
  If (Not okOld) Or (Not okNew) Then
    P6BandBenchEmit n, bw, reps, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "SOLVE_FAIL"
    Exit Sub
  End If
  perOld = tOldFac / CDbl(reps)
  perNew = tNewFac / CDbl(reps)
  perPack = tPack / CDbl(reps)
  perSolOld = tSolveOld / CDbl(reps)
  perSolNew = tSolveNew / CDbl(reps)
  timed = (perNew > 0#)
  speedupLoop = 0#
  speedupNet = 0#
  If timed Then
    speedupLoop = perOld / perNew
    speedupNet = perOld / (perNew + perPack)
  End If
  gateText = "FAIL"
  If ex > 0.0000000001 Then
    reasonText = "EX"
  ElseIf Not timed Then
    reasonText = "TOO_FAST"
  ElseIf speedupLoop >= 1.1 Then
    gateText = "PASS"
    reasonText = "PASS_NOT_SWITCHED"
  ElseIf speedupLoop >= 1.05 Then
    gateText = "CONDITIONAL"
    reasonText = "MARGINAL"
  Else
    reasonText = "SLOW"
  End If
  P6BandBenchEmit n, bw, reps, P6BandPlain(perOld, True), P6BandPlain(perNew, timed), P6BandPlain(perSolOld, True), P6BandPlain(perSolNew, True), P6BandPlain(perPack, True), P6BandPlain(speedupLoop, timed), P6BandPlain(speedupNet, timed), P6BandSci(ex, True), P6BandSci(erOld, True), P6BandSci(erNew, True), gateText, reasonText
  Exit Sub
BenchFail:
  If benchErrLogged Then Exit Sub
  benchErrLogged = True
  On Error Resume Next
  P6BandBenchEmit n, bw, reps, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "FAIL", "ERR"
  Err.Clear
End Sub

Private Sub P6EnsureIterativeWorkspace(ByVal vectorCount As Long, ByVal restart As Long)
  If vectorCount < 1 Then Exit Sub
  If restart < 2 Then restart = 2
  If P6IterativeVectorCapacity >= vectorCount And P6IterativeRestartCapacity >= restart Then Exit Sub
  ReDim P6BiCGB(0 To vectorCount - 1)
  ReDim P6BiCGX(0 To vectorCount - 1)
  ReDim P6BiCGR(0 To vectorCount - 1)
  ReDim P6BiCGRHat(0 To vectorCount - 1)
  ReDim P6BiCGP(0 To vectorCount - 1)
  ReDim P6BiCGV(0 To vectorCount - 1)
  ReDim P6BiCGS(0 To vectorCount - 1)
  ReDim P6BiCGT(0 To vectorCount - 1)
  ReDim P6BiCGPHat(0 To vectorCount - 1)
  ReDim P6BiCGSHat(0 To vectorCount - 1)
  ReDim P6BiCGAX(0 To vectorCount - 1)
  ReDim P6GMRESB(0 To vectorCount - 1)
  ReDim P6GMRESX(0 To vectorCount - 1)
  ReDim P6GMRESR(0 To vectorCount - 1)
  ReDim P6GMRESZ(0 To vectorCount - 1)
  ReDim P6GMRESW(0 To vectorCount - 1)
  ReDim P6GMRESInput(0 To vectorCount - 1)
  ReDim P6GMRESBasis(0 To vectorCount - 1, 0 To restart)
  ReDim P6GMRESH(0 To restart, 0 To restart - 1)
  ReDim P6GMRESCosine(0 To restart - 1)
  ReDim P6GMRESSine(0 To restart - 1)
  ReDim P6GMRESG(0 To restart)
  ReDim P6GMRESY(0 To restart - 1)
  P6IterativeVectorCapacity = vectorCount
  P6IterativeRestartCapacity = restart
End Sub

Public Sub P6EnsureResidualVectorWorkspace()
  If nDof <= 0 Then Exit Sub
  If P6ResidualVectorCapacity <> nDof Then
    ReDim P6CSRResidualProduct(0 To lastDof)
    P6ResidualVectorCapacity = nDof
  End If
End Sub

Public Sub P6EnsureReactionWorkspace()
  If nDof <= 0 Or lastDof < 0 Then Exit Sub
  If P6ReactionCapacity < nDof Then
    ReDim Reaction(0 To lastDof)
    P6ReactionCapacity = nDof
  End If
End Sub

Private Function P6Q4Shape(ByVal gaussId As Long, ByVal cornerId As Long) As Double
  Dim xi As Double, eta As Double, g As Double
  g = 0.5773502692
  Select Case gaussId
    Case 0: xi = -g: eta = -g
    Case 1: xi = g: eta = -g
    Case 2: xi = g: eta = g
    Case Else: xi = -g: eta = g
  End Select
  Select Case cornerId
    Case 0: P6Q4Shape = 0.25 * (1# - xi) * (1# - eta)
    Case 1: P6Q4Shape = 0.25 * (1# + xi) * (1# - eta)
    Case 2: P6Q4Shape = 0.25 * (1# + xi) * (1# + eta)
    Case Else: P6Q4Shape = 0.25 * (1# - xi) * (1# + eta)
  End Select
End Function

Private Sub P6Q4Deriv(ByVal gaussId As Long, ByVal cornerId As Long, ByRef dNxi As Double, ByRef dNeta As Double)
  Dim xi As Double, eta As Double, xiC As Double, etaC As Double, g As Double
  g = 0.5773502692
  Select Case gaussId
    Case 0: xi = -g: eta = -g
    Case 1: xi = g: eta = -g
    Case 2: xi = g: eta = g
    Case Else: xi = -g: eta = g
  End Select
  Select Case cornerId
    Case 0: xiC = -1#: etaC = -1#
    Case 1: xiC = 1#: etaC = -1#
    Case 2: xiC = 1#: etaC = 1#
    Case Else: xiC = -1#: etaC = 1#
  End Select
  dNxi = 0.25 * xiC * (1# + etaC * eta)
  dNeta = 0.25 * etaC * (1# + xiC * xi)
End Sub

Private Sub P6BuildPressureMap()
  Dim nodeId As Long, elementId As Long, cornerId As Long
  ReDim P6NodePressure(0 To NumberOfFreeNode - 1)
  For nodeId = 0 To NumberOfFreeNode - 1
    P6NodePressure(nodeId) = -1
  Next nodeId
  For elementId = 0 To NumberOfElement - 1
    For cornerId = 0 To 3
      nodeId = Elem(elementId).node(cornerId)
      If nodeId >= 0 And nodeId < NumberOfFreeNode Then P6NodePressure(nodeId) = 0
    Next cornerId
  Next elementId
  P6PressureCount = 0
  For nodeId = 0 To NumberOfFreeNode - 1
    If P6NodePressure(nodeId) = 0 Then
      P6NodePressure(nodeId) = P6PressureCount
      P6PressureCount = P6PressureCount + 1
    End If
  Next nodeId
End Sub

Private Sub P6SubtractVolumetricFromCSR()
  Dim elementId As Long, localRow As Long, localCol As Long, gaussId As Long
  Dim position As Long
  Dim lambdaValue As Double, thicknessValue As Double, kVol As Double
  Dim rowVol As Double, colVol As Double
  If P6CSRDeviatoricReady Then Exit Sub
  For elementId = 0 To NumberOfElement - 1
    lambdaValue = Elem(elementId).Dmat(0, 1)
    thicknessValue = Material(Elem(elementId).MatNo).thickness
    If Abs(lambdaValue) <= 1E-30 Or thicknessValue <= 0# Then GoTo NextVolElement
    For localRow = 0 To 15
      For localCol = 0 To 15
        kVol = 0#
        For gaussId = 0 To 3
          rowVol = Elem(elementId).Bmat(0, localRow, gaussId) + Elem(elementId).Bmat(1, localRow, gaussId)
          colVol = Elem(elementId).Bmat(0, localCol, gaussId) + Elem(elementId).Bmat(1, localCol, gaussId)
          kVol = kVol + lambdaValue * rowVol * colVol * Elem(elementId).dj(gaussId) * thicknessValue
        Next gaussId
        position = P6CSRElementPosition(elementId, localRow, localCol)
        If position >= 0 And position < P6CSRNNZ Then
          P6CSROriginalValues(position) = P6CSROriginalValues(position) - kVol
        End If
      Next localCol
    Next localRow
NextVolElement:
  Next elementId
  P6CSRDeviatoricReady = True
End Sub

Private Sub P6AssembleMixedCoupling()
  Dim elementId As Long, localDof As Long, cornerId As Long, gaussId As Long
  Dim pressureIndex As Long, nodeId As Long
  Dim lambdaValue As Double, thicknessValue As Double, shapeValue As Double, volStrain As Double
  ReDim P6MixedQ(0 To NumberOfElement - 1, 0 To 15, 0 To 3)
  ReDim P6MixedCompress(0 To P6PressureCount - 1)
  ReDim P6Pressure(0 To P6PressureCount - 1)
  For elementId = 0 To NumberOfElement - 1
    lambdaValue = Elem(elementId).Dmat(0, 1)
    thicknessValue = Material(Elem(elementId).MatNo).thickness
    If lambdaValue < 1# Then lambdaValue = 1#
    For gaussId = 0 To 3
      For cornerId = 0 To 3
        shapeValue = P6Q4Shape(gaussId, cornerId)
        nodeId = Elem(elementId).node(cornerId)
        If nodeId < 0 Or nodeId >= NumberOfFreeNode Then GoTo NextMixMapCorner
        pressureIndex = P6NodePressure(nodeId)
        If pressureIndex < 0 Then GoTo NextMixMapCorner
        P6MixedCompress(pressureIndex) = P6MixedCompress(pressureIndex) + shapeValue * Elem(elementId).dj(gaussId) * thicknessValue / lambdaValue
        For localDof = 0 To 15
          volStrain = Elem(elementId).Bmat(0, localDof, gaussId) + Elem(elementId).Bmat(1, localDof, gaussId)
          P6MixedQ(elementId, localDof, cornerId) = P6MixedQ(elementId, localDof, cornerId) + volStrain * shapeValue * Elem(elementId).dj(gaussId) * thicknessValue
        Next localDof
NextMixMapCorner:
      Next cornerId
    Next gaussId
  Next elementId
End Sub

Private Sub P6BuildMixedSchur()
  Dim elementId As Long, localDof As Long, cornerId As Long
  Dim pressureIndex As Long, nodeId As Long, uDof As Long, diagPos As Long
  Dim diagValue As Double, couplingValue As Double
  ReDim P6MixedSchurInv(0 To P6PressureCount - 1)
  For pressureIndex = 0 To P6PressureCount - 1
    P6MixedSchurInv(pressureIndex) = P6MixedCompress(pressureIndex)
  Next pressureIndex
  For elementId = 0 To NumberOfElement - 1
    For cornerId = 0 To 3
      nodeId = Elem(elementId).node(cornerId)
      If nodeId < 0 Or nodeId >= NumberOfFreeNode Then GoTo NextSchurCorner
      pressureIndex = P6NodePressure(nodeId)
      If pressureIndex < 0 Then GoTo NextSchurCorner
      For localDof = 0 To 15
        uDof = Elem(elementId).ElNode(localDof)
        If uDof < 0 Or uDof > lastDof Then GoTo NextSchurLocal
        diagPos = P6CSRDiagonalPosition(uDof)
        diagValue = 1#
        If diagPos >= 0 And diagPos < P6CSRNNZ Then diagValue = Abs(P6CSROriginalValues(diagPos))
        If diagValue < 0.000000000001 Then diagValue = 0.000000000001
        couplingValue = P6MixedQ(elementId, localDof, cornerId)
        P6MixedSchurInv(pressureIndex) = P6MixedSchurInv(pressureIndex) + couplingValue * couplingValue / diagValue
NextSchurLocal:
      Next localDof
NextSchurCorner:
    Next cornerId
  Next elementId
  For pressureIndex = 0 To P6PressureCount - 1
    If Abs(P6MixedSchurInv(pressureIndex)) < 0.000000000001 Then
      P6MixedSchurInv(pressureIndex) = 1#
    Else
      P6MixedSchurInv(pressureIndex) = 1# / P6MixedSchurInv(pressureIndex)
    End If
  Next pressureIndex
End Sub

Private Function P6PrepareBiotMixed() As Boolean
  Dim storageValue As Double, porosityValue As Double, waterBulk As Double
  P6PrepareBiotMixed = False
  P6ConsolActive = False
  P6MixedUP = False
  FEMValidateAnalysisScope
  If P6ReadSetting("CONSOL_ENABLE", 0#) = 0# Then
    P6PrepareBiotMixed = True
    Exit Function
  End If
  If nDof <= 0 Or NumberOfElement <= 0 Then Exit Function
  If Not P6BiotGeometryReady Then
    P6BuildPressureMap
    If P6PressureCount <= 0 Then Exit Function
    P6AssembleBiotCoupling
    P6AssembleBiotPermeability
    P6BuildQColumns
    P6MarkDrainedPressures
    P6BiotGeometryReady = True
  End If
  If P6PressureCount <= 0 Then Exit Function
  storageValue = P6ReadSetting("CONSOL_STORAGE", 0#)
  If storageValue <= 0# Then
    porosityValue = P6ReadSetting("CONSOL_POROSITY", 0.4)
    waterBulk = P6ReadSetting("CONSOL_KW", 2200000#)
    If waterBulk < 1# Then waterBulk = 2200000#
    storageValue = porosityValue / waterBulk
  End If
  If storageValue <= 0# Then storageValue = 0.000000000001
  P6RescaleBiotStorage storageValue
  P6BuildBiotCapacity
  P6MixedUP = True
  P6ConsolActive = True
  P6PrepareBiotMixed = True
End Function

Private Sub P6RescaleBiotStorage(ByVal storageValue As Double)
  Dim pressureIndex As Long, elementId As Long, cornerId As Long, gaussId As Long
  Dim nodeId As Long, thicknessValue As Double, shapeValue As Double
  If P6PressureCount <= 0 Then Exit Sub
  ReDim P6MixedS(0 To P6PressureCount - 1)
  For elementId = 0 To NumberOfElement - 1
    thicknessValue = Material(Elem(elementId).MatNo).thickness
    If thicknessValue <= 0# Then thicknessValue = 1#
    For gaussId = 0 To 3
      For cornerId = 0 To 3
        nodeId = Elem(elementId).node(cornerId)
        If nodeId < 0 Or nodeId >= NumberOfFreeNode Then GoTo NextStorageCorner
        pressureIndex = P6NodePressure(nodeId)
        If pressureIndex < 0 Then GoTo NextStorageCorner
        shapeValue = P6Q4Shape(gaussId, cornerId)
        P6MixedS(pressureIndex) = P6MixedS(pressureIndex) + storageValue * shapeValue * Elem(elementId).dj(gaussId) * thicknessValue
NextStorageCorner:
      Next cornerId
    Next gaussId
  Next elementId
End Sub

Private Sub P6AssembleBiotCoupling()
  Dim elementId As Long, localDof As Long, cornerId As Long, gaussId As Long
  Dim pressureIndex As Long, nodeId As Long
  Dim thicknessValue As Double, shapeValue As Double, volStrain As Double
  ReDim P6MixedQ(0 To NumberOfElement - 1, 0 To 15, 0 To 3)
  ReDim P6MixedS(0 To P6PressureCount - 1)
  ReDim P6Pressure(0 To P6PressureCount - 1)
  ReDim P6CommittedPressure(0 To P6PressureCount - 1)
  ReDim P6ContinuityG(0 To P6PressureCount - 1)
  ReDim P6PressureResidual(0 To P6PressureCount - 1)
  For elementId = 0 To NumberOfElement - 1
    thicknessValue = Material(Elem(elementId).MatNo).thickness
    If thicknessValue <= 0# Then thicknessValue = 1#
    For gaussId = 0 To 3
      For cornerId = 0 To 3
        shapeValue = P6Q4Shape(gaussId, cornerId)
        nodeId = Elem(elementId).node(cornerId)
        If nodeId < 0 Or nodeId >= NumberOfFreeNode Then GoTo NextBiotCorner
        pressureIndex = P6NodePressure(nodeId)
        If pressureIndex < 0 Then GoTo NextBiotCorner
        For localDof = 0 To 15
          volStrain = Elem(elementId).Bmat(0, localDof, gaussId) + Elem(elementId).Bmat(1, localDof, gaussId)
          P6MixedQ(elementId, localDof, cornerId) = P6MixedQ(elementId, localDof, cornerId) + volStrain * shapeValue * Elem(elementId).dj(gaussId) * thicknessValue
        Next localDof
NextBiotCorner:
      Next cornerId
    Next gaussId
  Next elementId
End Sub

Private Sub P6AssembleBiotPermeability()
  Dim elementId As Long, gaussId As Long, cornerA As Long, cornerB As Long
  Dim nodeA As Long, nodeB As Long, pressureA As Long, pressureB As Long
  Dim thicknessValue As Double, permeability As Double, gammaW As Double, conductivity As Double
  Dim dNxiA As Double, dNetaA As Double, dNxiB As Double, dNetaB As Double
  Dim j11 As Double, j12 As Double, j21 As Double, j22 As Double, detJ As Double
  Dim dNxA As Double, dNyA As Double, dNxB As Double, dNyB As Double
  Dim cornerId As Long
  permeability = P6ReadSetting("CONSOL_K", 0.00000001)
  gammaW = P6ReadSetting("CONSOL_GAMMA_W", 9.81)
  If gammaW < 0.000000000001 Then gammaW = 9.81
  conductivity = permeability / gammaW
  If conductivity < 0# Then conductivity = 0#
  ReDim P6MixedH(0 To P6PressureCount - 1)
  For elementId = 0 To NumberOfElement - 1
    thicknessValue = Material(Elem(elementId).MatNo).thickness
    If thicknessValue <= 0# Then thicknessValue = 1#
    For gaussId = 0 To 3
      j11 = 0#: j12 = 0#: j21 = 0#: j22 = 0#
      For cornerId = 0 To 3
        P6Q4Deriv gaussId, cornerId, dNxiA, dNetaA
        j11 = j11 + dNxiA * Elem(elementId).x(cornerId)
        j12 = j12 + dNetaA * Elem(elementId).x(cornerId)
        j21 = j21 + dNxiA * Elem(elementId).y(cornerId)
        j22 = j22 + dNetaA * Elem(elementId).y(cornerId)
      Next cornerId
      detJ = j11 * j22 - j12 * j21
      If Abs(detJ) <= 0.000000000001 Then GoTo NextPermGauss
      For cornerA = 0 To 3
        nodeA = Elem(elementId).node(cornerA)
        If nodeA < 0 Or nodeA >= NumberOfFreeNode Then GoTo NextPermA
        pressureA = P6NodePressure(nodeA)
        If pressureA < 0 Then GoTo NextPermA
        P6Q4Deriv gaussId, cornerA, dNxiA, dNetaA
        dNxA = (j22 * dNxiA - j12 * dNetaA) / detJ
        dNyA = (-j21 * dNxiA + j11 * dNetaA) / detJ
        For cornerB = 0 To 3
          nodeB = Elem(elementId).node(cornerB)
          If nodeB < 0 Or nodeB >= NumberOfFreeNode Then GoTo NextPermB
          pressureB = P6NodePressure(nodeB)
          If pressureB < 0 Then GoTo NextPermB
          P6Q4Deriv gaussId, cornerB, dNxiB, dNetaB
          dNxB = (j22 * dNxiB - j12 * dNetaB) / detJ
          dNyB = (-j21 * dNxiB + j11 * dNetaB) / detJ
          P6MixedH(pressureA) = P6MixedH(pressureA) + conductivity * (dNxA * dNxB + dNyA * dNyB) * Abs(detJ) * thicknessValue
NextPermB:
        Next cornerB
NextPermA:
      Next cornerA
NextPermGauss:
    Next gaussId
  Next elementId
End Sub

Private Sub P6BuildQColumns()
  Dim elementId As Long, localDof As Long, cornerId As Long
  Dim pressureIndex As Long, nodeId As Long, cursor As Long
  Dim counts() As Long
  ReDim counts(0 To P6PressureCount)
  ReDim P6QStart(0 To P6PressureCount)
  For elementId = 0 To NumberOfElement - 1
    For cornerId = 0 To 3
      nodeId = Elem(elementId).node(cornerId)
      If nodeId < 0 Or nodeId >= NumberOfFreeNode Then GoTo NextQCountCorner
      pressureIndex = P6NodePressure(nodeId)
      If pressureIndex < 0 Then GoTo NextQCountCorner
      counts(pressureIndex) = counts(pressureIndex) + 16
NextQCountCorner:
    Next cornerId
  Next elementId
  P6QStart(0) = 0
  For pressureIndex = 0 To P6PressureCount - 1
    P6QStart(pressureIndex + 1) = P6QStart(pressureIndex) + counts(pressureIndex)
    counts(pressureIndex) = P6QStart(pressureIndex)
  Next pressureIndex
  P6QEntryCount = P6QStart(P6PressureCount)
  If P6QEntryCount <= 0 Then Exit Sub
  ReDim P6QUDof(0 To P6QEntryCount - 1)
  ReDim P6QVal(0 To P6QEntryCount - 1)
  For elementId = 0 To NumberOfElement - 1
    For cornerId = 0 To 3
      nodeId = Elem(elementId).node(cornerId)
      If nodeId < 0 Or nodeId >= NumberOfFreeNode Then GoTo NextQFillCorner
      pressureIndex = P6NodePressure(nodeId)
      If pressureIndex < 0 Then GoTo NextQFillCorner
      For localDof = 0 To 15
        cursor = counts(pressureIndex)
        P6QUDof(cursor) = Elem(elementId).ElNode(localDof)
        P6QVal(cursor) = P6MixedQ(elementId, localDof, cornerId)
        counts(pressureIndex) = cursor + 1
      Next localDof
NextQFillCorner:
    Next cornerId
  Next elementId
End Sub

Private Sub P6MarkDrainedPressures()
  Dim elementId As Long, cornerId As Long, nodeId As Long, pressureIndex As Long
  Dim maxY As Double, yValue As Double, yTol As Double
  ReDim P6PressureFixed(0 To P6PressureCount - 1)
  P6DrainCount = 0
  If P6ReadSetting("CONSOL_DRAIN_TOP", 1#) = 0# Then Exit Sub
  maxY = -1E+30
  For elementId = 0 To NumberOfElement - 1
    For cornerId = 0 To 3
      If Elem(elementId).y(cornerId) > maxY Then maxY = Elem(elementId).y(cornerId)
    Next cornerId
  Next elementId
  yTol = Abs(maxY) * 0.000000001
  If yTol < 0.000000001 Then yTol = 0.000000001
  For elementId = 0 To NumberOfElement - 1
    For cornerId = 0 To 3
      nodeId = Elem(elementId).node(cornerId)
      If nodeId < 0 Or nodeId >= NumberOfFreeNode Then GoTo NextDrainCorner
      pressureIndex = P6NodePressure(nodeId)
      If pressureIndex < 0 Then GoTo NextDrainCorner
      yValue = Elem(elementId).y(cornerId)
      If Abs(yValue - maxY) <= yTol Then
        If Not P6PressureFixed(pressureIndex) Then
          P6PressureFixed(pressureIndex) = True
          P6DrainCount = P6DrainCount + 1
          P6Pressure(pressureIndex) = 0#
          P6CommittedPressure(pressureIndex) = 0#
        End If
      End If
NextDrainCorner:
    Next cornerId
  Next elementId
End Sub

Public Sub P6BuildBiotCapacity()
  Dim pressureIndex As Long
  Dim floorValue As Double
  ReDim P6MixedC(0 To P6PressureCount - 1)
  floorValue = 0.000000000001
  For pressureIndex = 0 To P6PressureCount - 1
    P6MixedC(pressureIndex) = P6MixedS(pressureIndex) + P6ConsolDt * P6MixedH(pressureIndex)
    If P6PressureFixed(pressureIndex) Then
      P6MixedC(pressureIndex) = 1#
    ElseIf P6MixedC(pressureIndex) < floorValue Then
      P6MixedC(pressureIndex) = floorValue
    End If
  Next pressureIndex
End Sub

Private Sub P6AddSchurContribution()
  Dim pressureIndex As Long, entryIndex As Long, iIndex As Long, jIndex As Long
  Dim nzCount As Long, nzA As Long, nzB As Long
  Dim rowNo As Long, colNo As Long, offsetCol As Long
  Dim cInv As Double, rowValue As Double, colValue As Double
  Dim work() As Double, nzList() As Long
  If Not P6MixedUP Or P6PressureCount <= 0 Then Exit Sub
  ReDim work(0 To lastDof)
  ReDim nzList(0 To lastDof)
  For pressureIndex = 0 To P6PressureCount - 1
    If P6PressureFixed(pressureIndex) Then GoTo NextSchurPressure
    If Abs(P6MixedC(pressureIndex)) <= 0.000000000001 Then GoTo NextSchurPressure
    cInv = 1# / P6MixedC(pressureIndex)
    nzCount = 0
    For entryIndex = P6QStart(pressureIndex) To P6QStart(pressureIndex + 1) - 1
      rowNo = P6QUDof(entryIndex)
      If rowNo < 0 Or rowNo > lastDof Then GoTo NextSchurEntry
      If NodeCond(rowNo) <> 0 Then GoTo NextSchurEntry
      If work(rowNo) = 0# Then
        nzList(nzCount) = rowNo
        nzCount = nzCount + 1
      End If
      work(rowNo) = work(rowNo) + P6QVal(entryIndex)
NextSchurEntry:
    Next entryIndex
    For nzA = 0 To nzCount - 1
      rowNo = nzList(nzA)
      rowValue = work(rowNo)
      For nzB = 0 To nzCount - 1
        colNo = nzList(nzB)
        If colNo >= rowNo Then
          offsetCol = colNo - rowNo
          If offsetCol <= BandWidth Then
            TotalMat(rowNo, offsetCol) = TotalMat(rowNo, offsetCol) + cInv * rowValue * work(colNo)
          End If
        End If
      Next nzB
    Next nzA
    For nzA = 0 To nzCount - 1
      work(nzList(nzA)) = 0#
    Next nzA
NextSchurPressure:
  Next pressureIndex
End Sub

Public Sub P6RebuildBandWithSchur()
  Dim i As Long, j As Long
  If lastDof < 0 Or BandWidth < 0 Then Exit Sub
  ReDim TotalMat(lastDof, BandWidth)
  For i = 0 To lastDof
    For j = 0 To BandWidth
      TotalMat(i, j) = OriginalMat(i, j)
    Next j
  Next i
  P6AddSchurContribution
  P6FactorReady = False
  MatrixFactored = False
End Sub

Public Sub P6UpdateContinuityG()
  Dim pressureIndex As Long
  If Not P6MixedUP Then Exit Sub
  P6MixedApplyQTu P3CommittedDisp, P6ContinuityG
  For pressureIndex = 0 To P6PressureCount - 1
    If P6PressureFixed(pressureIndex) Then
      P6ContinuityG(pressureIndex) = 0#
    Else
      P6ContinuityG(pressureIndex) = P6ContinuityG(pressureIndex) + P6MixedS(pressureIndex) * P6CommittedPressure(pressureIndex)
    End If
  Next pressureIndex
End Sub

Public Sub P6ComputePressureResidual()
  Dim pressureIndex As Long
  If Not P6MixedUP Then Exit Sub
  P6MixedApplyQTu TDisp, P6PressureResidual
  For pressureIndex = 0 To P6PressureCount - 1
    If P6PressureFixed(pressureIndex) Then
      P6PressureResidual(pressureIndex) = 0#
    Else
      P6PressureResidual(pressureIndex) = P6PressureResidual(pressureIndex) + P6MixedC(pressureIndex) * P6Pressure(pressureIndex) - P6ContinuityG(pressureIndex)
    End If
  Next pressureIndex
End Sub

Private Sub P6ApplySchurRHS()
  Dim pressureIndex As Long, entryIndex As Long, uDof As Long
  Dim scaledP As Double
  If Not P6MixedUP Then Exit Sub
  P6ComputePressureResidual
  For pressureIndex = 0 To P6PressureCount - 1
    If P6PressureFixed(pressureIndex) Then GoTo NextRhsPressure
    If Abs(P6MixedC(pressureIndex)) <= 0.000000000001 Then GoTo NextRhsPressure
    scaledP = P6PressureResidual(pressureIndex) / P6MixedC(pressureIndex)
    For entryIndex = P6QStart(pressureIndex) To P6QStart(pressureIndex + 1) - 1
      uDof = P6QUDof(entryIndex)
      If uDof >= 0 And uDof <= lastDof Then
        If NodeCond(uDof) = 0 Then Force(uDof) = Force(uDof) + P6QVal(entryIndex) * scaledP
      End If
    Next entryIndex
NextRhsPressure:
  Next pressureIndex
End Sub

Private Sub P6RecoverPressureIncrement()
  Dim pressureIndex As Long, entryIndex As Long, uDof As Long
  Dim qtdu As Double
  If Not P6MixedUP Then Exit Sub
  For pressureIndex = 0 To P6PressureCount - 1
    If P6PressureFixed(pressureIndex) Then
      P6Pressure(pressureIndex) = 0#
      GoTo NextRecoverP
    End If
    qtdu = 0#
    For entryIndex = P6QStart(pressureIndex) To P6QStart(pressureIndex + 1) - 1
      uDof = P6QUDof(entryIndex)
      If uDof >= 0 And uDof <= lastDof Then qtdu = qtdu + P6QVal(entryIndex) * Disp(uDof)
    Next entryIndex
    P6Pressure(pressureIndex) = P6Pressure(pressureIndex) - (P6PressureResidual(pressureIndex) + qtdu) / P6MixedC(pressureIndex)
NextRecoverP:
  Next pressureIndex
End Sub

Public Sub P6AddPressureToInternalForce()
  Dim pressureIndex As Long, entryIndex As Long, uDof As Long
  If Not P6MixedUP Then Exit Sub
  For pressureIndex = 0 To P6PressureCount - 1
    For entryIndex = P6QStart(pressureIndex) To P6QStart(pressureIndex + 1) - 1
      uDof = P6QUDof(entryIndex)
      If uDof >= 0 And uDof <= lastDof Then
        If NodeCond(uDof) = 0 Then iNForce(uDof) = iNForce(uDof) + P6QVal(entryIndex) * P6Pressure(pressureIndex)
      End If
    Next entryIndex
  Next pressureIndex
End Sub

Private Function P6PrepareMixedUP() As Boolean
  P6PrepareMixedUP = False
  P6MixedUP = False
  P6KrylovLast = lastDof
  If Not P6UseCSR Or P3FlowPolicyIsInconsistent() Then
    ' Total-stress displacement formulation. CSR must not silently activate
    ' the dormant mixed pressure formulation through MIXED_UP's default.
    P6PrepareMixedUP = True
    Exit Function
  End If
  If P6ReadSetting("MIXED_UP", 1#) = 0# Then
    P6PrepareMixedUP = True
    Exit Function
  End If
  If nDof <= 0 Or NumberOfElement <= 0 Then Exit Function
  P6BuildPressureMap
  If P6PressureCount <= 0 Then Exit Function
  P6SubtractVolumetricFromCSR
  P6AssembleMixedCoupling
  P6BuildMixedSchur
  P6KrylovLast = lastDof
  P6MixedUP = True
  P6MixedCoupledKrylov = False
  P6CSRDiagonalInverseReady = False
  P6CSRILUReady = False
  P6CSRBoundaryApplied = False
  If Not ApplyCSRBoundaryToMatrix() Then Exit Function
  P6CSRBoundaryApplied = True
  P6EnsureIterativeWorkspace lastDof + 1, P6GMRESRestart
  P6PrepareMixedUP = True
End Function

Private Sub P6MixedCompleteMatVec(ByRef inputVector() As Double, ByRef outputVector() As Double)
  Dim elementId As Long, localDof As Long, cornerId As Long
  Dim pressureIndex As Long, nodeId As Long, uDof As Long
  Dim pressureValue As Double, couplingValue As Double
  If Not P6MixedUP Or Not P6MixedCoupledKrylov Then Exit Sub
  For pressureIndex = 0 To P6PressureCount - 1
    outputVector(nDof + pressureIndex) = 0#
  Next pressureIndex
  For elementId = 0 To NumberOfElement - 1
    For cornerId = 0 To 3
      nodeId = Elem(elementId).node(cornerId)
      If nodeId < 0 Or nodeId >= NumberOfFreeNode Then GoTo NextApplyCorner
      pressureIndex = P6NodePressure(nodeId)
      If pressureIndex < 0 Then GoTo NextApplyCorner
      pressureValue = inputVector(nDof + pressureIndex)
      For localDof = 0 To 15
        uDof = Elem(elementId).ElNode(localDof)
        couplingValue = P6MixedQ(elementId, localDof, cornerId)
        outputVector(uDof) = outputVector(uDof) + couplingValue * pressureValue
        outputVector(nDof + pressureIndex) = outputVector(nDof + pressureIndex) + couplingValue * inputVector(uDof)
      Next localDof
NextApplyCorner:
    Next cornerId
  Next elementId
  For pressureIndex = 0 To P6PressureCount - 1
    outputVector(nDof + pressureIndex) = outputVector(nDof + pressureIndex) - P6MixedCompress(pressureIndex) * inputVector(nDof + pressureIndex)
  Next pressureIndex
  For uDof = 0 To lastDof
    If NodeCond(uDof) <> 0 Then outputVector(uDof) = inputVector(uDof)
  Next uDof
End Sub

Public Sub P6CSRMatVec(ByRef inputVector() As Double, ByRef outputVector() As Double, ByRef matrixValues() As Double)
  Dim rowNo As Long, position As Long
  For rowNo = 0 To lastDof
    outputVector(rowNo) = 0#
    For position = P6CSRRowPtr(rowNo) To P6CSRRowPtr(rowNo + 1) - 1
      outputVector(rowNo) = outputVector(rowNo) + matrixValues(position) * inputVector(P6CSRColumnIndex(position))
    Next position
  Next rowNo
  P6MixedCompleteMatVec inputVector, outputVector
End Sub

Private Sub P6CSRMatVecWithDot(ByRef inputVector() As Double, ByRef outputVector() As Double, _
                               ByRef matrixValues() As Double, ByRef dotVector() As Double, _
                               ByRef dotResult As Double)
  Dim rowNo As Long
  P6CSRMatVec inputVector, outputVector, matrixValues
  dotResult = 0#
  For rowNo = 0 To P6KrylovLast
    dotResult = dotResult + dotVector(rowNo) * outputVector(rowNo)
  Next rowNo
End Sub

Private Function P6CSRMatVecResidualNorm(ByRef inputVector() As Double, ByRef baseVector() As Double, _
                                         ByRef residualVector() As Double, ByRef matrixValues() As Double) As Double
  Dim rowNo As Long
  Dim residualValue As Double, sumSquares As Double
  P6CSRMatVec inputVector, residualVector, matrixValues
  sumSquares = 0#
  For rowNo = 0 To P6KrylovLast
    residualValue = baseVector(rowNo) - residualVector(rowNo)
    residualVector(rowNo) = residualValue
    sumSquares = sumSquares + residualValue * residualValue
  Next rowNo
  If sumSquares <= 0# Then
    P6CSRMatVecResidualNorm = 0#
  Else
    P6CSRMatVecResidualNorm = Sqr(sumSquares)
  End If
End Function

Private Sub P6CSRMatVecWithCommonDots(ByRef inputVector() As Double, ByRef outputVector() As Double, _
                                      ByRef matrixValues() As Double, ByRef firstVector() As Double, _
                                      ByRef secondVector() As Double, ByRef firstDot As Double, _
                                      ByRef secondDot As Double)
  Dim rowNo As Long
  P6CSRMatVec inputVector, outputVector, matrixValues
  firstDot = 0#: secondDot = 0#
  For rowNo = 0 To P6KrylovLast
    firstDot = firstDot + outputVector(rowNo) * firstVector(rowNo)
    secondDot = secondDot + outputVector(rowNo) * secondVector(rowNo)
  Next rowNo
End Sub

Private Function P6CSRDot(ByRef leftVector() As Double, ByRef rightVector() As Double) As Double
  Dim i As Long
  Dim dotValue As Double
  dotValue = 0#
  For i = 0 To P6KrylovLast
    dotValue = dotValue + leftVector(i) * rightVector(i)
  Next i
  P6CSRDot = dotValue
End Function

Private Function P6CSRNorm2(ByRef vectorValues() As Double) As Double
  Dim i As Long
  Dim sumSquares As Double, currentValue As Double
  sumSquares = 0#
  For i = 0 To P6KrylovLast
    currentValue = vectorValues(i)
    sumSquares = sumSquares + currentValue * currentValue
  Next i
  If sumSquares <= 0# Then
    P6CSRNorm2 = 0#
  Else
    P6CSRNorm2 = Sqr(sumSquares)
  End If
End Function

Private Sub P6CSRNorm2Pair(ByRef firstVector() As Double, ByRef secondVector() As Double, _
                           ByRef firstNorm As Double, ByRef secondNorm As Double)
  Dim i As Long, firstValue As Double, secondValue As Double
  Dim firstSum As Double, secondSum As Double
  firstSum = 0#: secondSum = 0#
  For i = 0 To P6KrylovLast
    firstValue = firstVector(i)
    secondValue = secondVector(i)
    firstSum = firstSum + firstValue * firstValue
    secondSum = secondSum + secondValue * secondValue
  Next i
  If firstSum <= 0# Then firstNorm = 0# Else firstNorm = Sqr(firstSum)
  If secondSum <= 0# Then secondNorm = 0# Else secondNorm = Sqr(secondSum)
End Sub

Private Sub P6CSRCommonDotPair(ByRef commonVector() As Double, ByRef firstVector() As Double, _
                               ByRef secondVector() As Double, ByRef firstDot As Double, _
                               ByRef secondDot As Double)
  Dim i As Long, commonValue As Double
  firstDot = 0#: secondDot = 0#
  For i = 0 To P6KrylovLast
    commonValue = commonVector(i)
    firstDot = firstDot + commonValue * firstVector(i)
    secondDot = secondDot + commonValue * secondVector(i)
  Next i
End Sub

Private Function P6EnsureCSRDiagonalInverse() As Boolean
  Dim rowNo As Long, diagonalPosition As Long, diagonalValue As Double
  P6EnsureCSRDiagonalInverse = False
  If Not P6CSRReady Or nDof <= 0 Or lastDof < 0 Then Exit Function
  If P6CSRDiagonalInverseReady Then
    P6EnsureCSRDiagonalInverse = True
    Exit Function
  End If
  ReDim P6CSRDiagonalInverse(0 To P6KrylovLast)
  For rowNo = 0 To lastDof
    diagonalPosition = P6CSRDiagonalPosition(rowNo)
    If diagonalPosition < 0 Then
      P6CSRDiagonalZeroCount = P6CSRDiagonalZeroCount + 1
      If P6CSRFirstZeroDiagonalDof < 0 Then P6CSRFirstZeroDiagonalDof = rowNo
      Exit Function
    End If
    diagonalValue = P6CSRValues(diagonalPosition)
    If Not P2IsFinite(diagonalValue) Or Abs(diagonalValue) <= 1E-30 Then
      P6CSRDiagonalZeroCount = P6CSRDiagonalZeroCount + 1
      If P6CSRFirstZeroDiagonalDof < 0 Then P6CSRFirstZeroDiagonalDof = rowNo
      If P6CSRMinimumDiagonal = 0# Or Abs(diagonalValue) < P6CSRMinimumDiagonal Then P6CSRMinimumDiagonal = diagonalValue
      Exit Function
    End If
    If P6CSRMinimumDiagonal = 0# Or diagonalValue < P6CSRMinimumDiagonal Then P6CSRMinimumDiagonal = diagonalValue
    P6CSRDiagonalInverse(rowNo) = 1# / diagonalValue
  Next rowNo
  If P6MixedCoupledKrylov And P6PressureCount > 0 Then
    For rowNo = 0 To P6PressureCount - 1
      P6CSRDiagonalInverse(nDof + rowNo) = P6MixedSchurInv(rowNo)
    Next rowNo
  End If
  P6CSRDiagonalInverseReady = True
  P6EnsureCSRDiagonalInverse = True
End Function

Private Function P6CSRApplyJacobi(ByRef inputVector() As Double, ByRef outputVector() As Double) As Boolean
  Dim rowNo As Long
  P6CSRApplyJacobi = False
  If P6CSRILUReady Then
    P6CSRPreconditionerMode = "ILU0"
    P6CSRApplyJacobi = P6CSRApplyILU(inputVector, outputVector)
    If Not P6CSRApplyJacobi Then P6CSRPreconditionerMode = "ILU0_APPLY_FAILED"
    Exit Function
  End If
  If Not P6EnsureCSRDiagonalInverse() Then Exit Function
  If Len(P6CSRPreconditionerMode) = 0 Then P6CSRPreconditionerMode = "JACOBI"
  For rowNo = 0 To P6KrylovLast
    outputVector(rowNo) = inputVector(rowNo) * P6CSRDiagonalInverse(rowNo)
  Next rowNo
  P6CSRApplyJacobi = True
End Function

Private Function P6EngineeringAcceptLimit(ByVal normB As Double) As Double
  Dim relativeLimit As Double
  relativeLimit = P6IterativeTolerance
  If relativeLimit < 0.00001 Then relativeLimit = 0.00001
  P6EngineeringAcceptLimit = relativeLimit * normB
End Function

Private Function P6AcceptCSRSolution(ByRef x() As Double, ByVal residualNorm As Double, ByVal normB As Double, ByVal acceptLimit As Double, ByVal exactLimit As Double) As Boolean
  Dim i As Long
  P6AcceptCSRSolution = False
  If residualNorm > acceptLimit Then Exit Function
  For i = 0 To lastDof
    Disp(i) = x(i)
  Next i
  If P6MixedUP Then
    For i = 0 To P6PressureCount - 1
      P6Pressure(i) = x(nDof + i)
    Next i
  End If
  P6IterativeLastResidual = residualNorm / normB
  If residualNorm > exactLimit Then
    If Len(P6IterativeFallbackStatus) = 0 Then P6IterativeFallbackStatus = "iterative_stalled_accept"
  End If
  P6AcceptCSRSolution = True
End Function

Private Function P6SolveCSRBICGSTAB() As Boolean
  Dim i As Long, iteration As Long, restartCount As Long
  Dim normB As Double, toleranceValue As Double, residualNorm As Double, acceptLimit As Double
  Dim residualSquare As Double, normBSquare As Double
  Dim rho As Double, rhoOld As Double, alpha As Double, omega As Double, beta As Double
  Dim denominator As Double, tt As Double, SS As Double
  Dim resetSearch As Boolean

  P6SolveCSRBICGSTAB = False
  P6EnsureIterativeWorkspace P6KrylovLast + 1, P6GMRESRestart
  normBSquare = 0#
  For i = 0 To P6KrylovLast
    If i <= lastDof Then P6BiCGB(i) = Force(i) Else P6BiCGB(i) = 0#
    P6BiCGX(i) = 0#
    P6BiCGAX(i) = 0#
    P6BiCGR(i) = P6BiCGB(i)
    P6BiCGRHat(i) = P6BiCGR(i)
    normBSquare = normBSquare + P6BiCGB(i) * P6BiCGB(i)
  Next i
  If normBSquare <= 0# Then
    normB = 0#
  Else
    normB = Sqr(normBSquare)
  End If
  residualNorm = normB
  If normB < 1# Then normB = 1#
  toleranceValue = P6IterativeTolerance * normB
  acceptLimit = P6EngineeringAcceptLimit(normB)
  P6IterativeLastIterations = 0
  P6IterativeLastResidual = residualNorm / normB
  If residualNorm <= toleranceValue Then
    If P6AcceptCSRSolution(P6BiCGX, residualNorm, normB, acceptLimit, toleranceValue) Then
      P6SolveCSRBICGSTAB = True
    End If
    Exit Function
  End If

  rhoOld = 1#: alpha = 1#: omega = 1#
  restartCount = 0
  resetSearch = True
  For iteration = 1 To P6IterativeMaxIterations
    rho = P6CSRDot(P6BiCGRHat, P6BiCGR)
    If Abs(rho) <= 1E-30 Then
      If Not P6BiCGRestartShadow(restartCount) Then Exit For
      rho = P6CSRDot(P6BiCGRHat, P6BiCGR)
      If Abs(rho) <= 1E-30 Then Exit For
      rhoOld = 1#: alpha = 1#: omega = 1#
      resetSearch = True
    End If
    If resetSearch Then
      For i = 0 To P6KrylovLast
        P6BiCGP(i) = P6BiCGR(i)
      Next i
      resetSearch = False
    Else
      If Abs(omega) <= 1E-30 Then
        If Not P6BiCGRestartShadow(restartCount) Then Exit For
        rho = P6CSRDot(P6BiCGRHat, P6BiCGR)
        If Abs(rho) <= 1E-30 Then Exit For
        rhoOld = 1#: alpha = 1#: omega = 1#
        For i = 0 To P6KrylovLast
          P6BiCGP(i) = P6BiCGR(i)
        Next i
      Else
        beta = (rho / rhoOld) * (alpha / omega)
        For i = 0 To P6KrylovLast
          P6BiCGP(i) = P6BiCGR(i) + beta * (P6BiCGP(i) - omega * P6BiCGV(i))
        Next i
      End If
    End If
    If Not P6CSRApplyJacobi(P6BiCGP, P6BiCGPHat) Then Exit For
    P6CSRMatVecWithDot P6BiCGPHat, P6BiCGV, P6CSRValues, P6BiCGRHat, denominator
    If Abs(denominator) <= 1E-30 Then
      If Not P6BiCGRestartShadow(restartCount) Then Exit For
      resetSearch = True
      GoTo NextBiCGIteration
    End If
    alpha = rho / denominator
    residualSquare = 0#
    For i = 0 To P6KrylovLast
      P6BiCGS(i) = P6BiCGR(i) - alpha * P6BiCGV(i)
      residualSquare = residualSquare + P6BiCGS(i) * P6BiCGS(i)
    Next i
    If residualSquare <= 0# Then residualNorm = 0# Else residualNorm = Sqr(residualSquare)
    If residualNorm <= acceptLimit Then
      For i = 0 To P6KrylovLast
        P6BiCGX(i) = P6BiCGX(i) + alpha * P6BiCGPHat(i)
      Next i
      P6IterativeLastIterations = iteration
      If P6AcceptCSRSolution(P6BiCGX, residualNorm, normB, acceptLimit, toleranceValue) Then
        If residualNorm > toleranceValue Then P6IterativeFallbackStatus = "bicgstab_stalled_accept"
        P6SolveCSRBICGSTAB = True
        Exit Function
      End If
    End If
    If Not P6CSRApplyJacobi(P6BiCGS, P6BiCGSHat) Then Exit For
    P6CSRMatVecWithCommonDots P6BiCGSHat, P6BiCGT, P6CSRValues, P6BiCGT, P6BiCGS, tt, SS
    If Abs(tt) <= 1E-30 Then
      If Not P6BiCGRestartShadow(restartCount) Then Exit For
      resetSearch = True
      GoTo NextBiCGIteration
    End If
    omega = SS / tt
    residualSquare = 0#
    For i = 0 To P6KrylovLast
      P6BiCGX(i) = P6BiCGX(i) + alpha * P6BiCGPHat(i) + omega * P6BiCGSHat(i)
      P6BiCGR(i) = P6BiCGS(i) - omega * P6BiCGT(i)
      residualSquare = residualSquare + P6BiCGR(i) * P6BiCGR(i)
    Next i
    If residualSquare <= 0# Then residualNorm = 0# Else residualNorm = Sqr(residualSquare)
    P6IterativeLastIterations = iteration
    P6IterativeLastResidual = residualNorm / normB
    If residualNorm <= acceptLimit Then
      If P6AcceptCSRSolution(P6BiCGX, residualNorm, normB, acceptLimit, toleranceValue) Then
        If residualNorm > toleranceValue Then P6IterativeFallbackStatus = "bicgstab_stalled_accept"
        P6SolveCSRBICGSTAB = True
        Exit Function
      End If
    End If
    rhoOld = rho
NextBiCGIteration:
  Next iteration

  residualNorm = P6CSRMatVecResidualNorm(P6BiCGX, P6BiCGB, P6BiCGR, P6CSRValues)
  If P6AcceptCSRSolution(P6BiCGX, residualNorm, normB, acceptLimit, toleranceValue) Then
    If residualNorm > toleranceValue Then P6IterativeFallbackStatus = "bicgstab_stalled_accept"
    P6SolveCSRBICGSTAB = True
    Exit Function
  End If
  If P6CSRDiagonalZeroCount > 0 Then Exit Function
End Function

Private Function P6BiCGRestartShadow(ByRef restartCount As Long) As Boolean
  Dim i As Long
  P6BiCGRestartShadow = False
  restartCount = restartCount + 1
  If restartCount > 3 Then Exit Function
  For i = 0 To P6KrylovLast
    P6BiCGRHat(i) = P6BiCGR(i)
  Next i
  P6IterativeFallbackStatus = "bicgstab_restart_" & CStr(restartCount)
  P6BiCGRestartShadow = True
End Function

Private Function P6SolveCSRGMRES() As Boolean
  Dim i As Long, j As Long, k As Long, pass As Long, used As Long, total As Long
  Dim restart As Long, normB As Double, beta As Double, nextNorm As Double
  Dim inner As Double, rot As Double, tmp As Double, residual As Double, tol As Double
  Dim zBasis() As Double
  P6SolveCSRGMRES = False
  P6IterativeLastResidual = 1E+100
  If P6CSRILUReady Then
    P6CSRPreconditionerMode = "ILU0"
  ElseIf Len(P6CSRPreconditionerMode) = 0 Then
    P6CSRPreconditionerMode = "JACOBI"
  End If
  On Error GoTo GMRESFailed
  If nDof < 1 Or P6KrylovLast <> lastDof Then Exit Function
  restart = P6GMRESRestart
  If restart < 2 Then restart = 30
  If restart > nDof Then restart = nDof
  If P6SolverMemoryLimitBytes > 0# Then
    ' Global Krylov workspace, local right-preconditioned basis and optional ILU.
    tmp = P6CSRStorageBytes + CDbl(P6CSRNNZ) * 8# + CDbl(nDof) * (23# + 2# * restart) * 8# + CDbl(restart + 1) ^ 2 * 8#
    If tmp > P6SolverMemoryLimitBytes Then
      SetAnalysisFailure RESULT_CAPACITY_ERROR, "非対称GMRESの作業メモリが設定上限を超えます。", vbObjectError + 3511, -1, -1, CurrentIncrement, CurrentIteration
      Exit Function
    End If
  End If
  P6EnsureIterativeWorkspace P6KrylovLast + 1, restart
  ReDim zBasis(0 To P6KrylovLast, 0 To restart - 1)
  For i = 0 To P6KrylovLast
    If Not P2IsFinite(Force(i)) Then Exit Function
    P6GMRESB(i) = Force(i): P6GMRESX(i) = 0#
    normB = normB + Force(i) * Force(i)
  Next i
  normB = Sqr(normB)
  If Not P2IsFinite(normB) Then Exit Function
  P6IterativeLastIterations = 0
  If normB <= 1E-30 Then
    For i = 0 To lastDof: Disp(i) = 0#: Next i
    P6IterativeLastResidual = 0#: P6SolveCSRGMRES = True: Exit Function
  End If
  tol = P6IterativeTolerance * normB
  Do While total < P6IterativeMaxIterations
    residual = P6CSRMatVecResidualNorm(P6GMRESX, P6GMRESB, P6GMRESR, P6CSRValues)
    If Not P2IsFinite(residual) Then
      P6SolverEvent "GMRES_BREAKDOWN", "reason=nonfinite_true_residual;iterations=" & CStr(total)
      Exit Function
    End If
    P6IterativeLastResidual = residual / normB
    P6SolverEvent "GMRES_RESTART", "restart=" & CStr(total \ restart) & ";iterations=" & CStr(total) & ";preconditioner=" & P6CSRPreconditionerMode & ";true_relres=" & Format$(P6IterativeLastResidual, "0.000E+00")
    If residual <= tol Then Exit Do
    beta = residual
    For i = 0 To P6KrylovLast: P6GMRESBasis(i, 0) = P6GMRESR(i) / beta: Next i
    For i = 0 To restart: P6GMRESG(i) = 0#: Next i
    P6GMRESG(0) = beta: used = 0
    For j = 0 To restart - 1
      For i = 0 To P6KrylovLast: P6GMRESInput(i) = P6GMRESBasis(i, j): Next i
      If Not P6CSRApplyJacobi(P6GMRESInput, P6GMRESZ) Then
        P6SolverEvent "GMRES_BREAKDOWN", "reason=preconditioner_apply;preconditioner=" & P6CSRPreconditionerMode & ";iterations=" & CStr(total)
        Exit Function
      End If
      For i = 0 To P6KrylovLast: zBasis(i, j) = P6GMRESZ(i): Next i
      P6CSRMatVec P6GMRESZ, P6GMRESW, P6CSRValues
      For k = 0 To restart: P6GMRESH(k, j) = 0#: Next k
      ' Right preconditioning preserves the true RHS/residual scale.
      ' Two MGS passes protect orthogonality near a plastic limit.
      For pass = 1 To 2
        For k = 0 To j
          inner = 0#
          For i = 0 To P6KrylovLast: inner = inner + P6GMRESW(i) * P6GMRESBasis(i, k): Next i
          P6GMRESH(k, j) = P6GMRESH(k, j) + inner
          For i = 0 To P6KrylovLast: P6GMRESW(i) = P6GMRESW(i) - inner * P6GMRESBasis(i, k): Next i
        Next k
      Next pass
      nextNorm = P6CSRNorm2(P6GMRESW): P6GMRESH(j + 1, j) = nextNorm
      If Not P2IsFinite(nextNorm) Then
        P6SolverEvent "GMRES_BREAKDOWN", "reason=nonfinite_krylov_norm;iterations=" & CStr(total)
        Exit Function
      End If
      If nextNorm > 1E-30 Then
        For i = 0 To P6KrylovLast: P6GMRESBasis(i, j + 1) = P6GMRESW(i) / nextNorm: Next i
      End If
      For k = 0 To j - 1
        tmp = P6GMRESCosine(k) * P6GMRESH(k, j) + P6GMRESSine(k) * P6GMRESH(k + 1, j)
        P6GMRESH(k + 1, j) = -P6GMRESSine(k) * P6GMRESH(k, j) + P6GMRESCosine(k) * P6GMRESH(k + 1, j)
        P6GMRESH(k, j) = tmp
      Next k
      rot = Sqr(P6GMRESH(j, j) ^ 2 + P6GMRESH(j + 1, j) ^ 2)
      If rot <= 1E-30 Then
        P6SolverEvent "GMRES_BREAKDOWN", "reason=zero_givens_rotation;preconditioner=" & P6CSRPreconditionerMode & ";iterations=" & CStr(total)
        Exit For
      End If
      P6GMRESCosine(j) = P6GMRESH(j, j) / rot: P6GMRESSine(j) = P6GMRESH(j + 1, j) / rot
      P6GMRESH(j, j) = rot: P6GMRESH(j + 1, j) = 0#
      tmp = P6GMRESCosine(j) * P6GMRESG(j)
      P6GMRESG(j + 1) = -P6GMRESSine(j) * P6GMRESG(j): P6GMRESG(j) = tmp
      used = j + 1: total = total + 1: P6IterativeLastIterations = total
      If Abs(P6GMRESG(j + 1)) <= tol Or nextNorm <= 1E-30 Or total >= P6IterativeMaxIterations Then Exit For
    Next j
    If used < 1 Then Exit Function
    For i = used - 1 To 0 Step -1
      tmp = P6GMRESG(i)
      For k = i + 1 To used - 1: tmp = tmp - P6GMRESH(i, k) * P6GMRESY(k): Next k
      If Abs(P6GMRESH(i, i)) <= 1E-30 Then
        P6SolverEvent "GMRES_BREAKDOWN", "reason=zero_upper_triangular_diagonal;preconditioner=" & P6CSRPreconditionerMode & ";iterations=" & CStr(total)
        Exit Function
      End If
      P6GMRESY(i) = tmp / P6GMRESH(i, i)
    Next i
    For k = 0 To used - 1
      For i = 0 To P6KrylovLast: P6GMRESX(i) = P6GMRESX(i) + zBasis(i, k) * P6GMRESY(k): Next i
    Next k
  Loop
  residual = P6CSRMatVecResidualNorm(P6GMRESX, P6GMRESB, P6GMRESR, P6CSRValues)
  If Not P2IsFinite(residual) Then Exit Function
  P6IterativeLastResidual = residual / normB
  If residual > tol Then Exit Function
  For i = 0 To lastDof
    If Not P2IsFinite(P6GMRESX(i)) Then Exit Function
  Next i
  For i = 0 To lastDof: Disp(i) = P6GMRESX(i): Next i
  P6SolveCSRGMRES = True
  Exit Function
GMRESFailed:
  If Err.Number = 7 Then SetAnalysisFailure RESULT_CAPACITY_ERROR, "非対称GMRESのメモリ確保に失敗しました。", Err.Number, -1, -1, CurrentIncrement, CurrentIteration
  Err.Clear
End Function

Private Function P6CSRApplyILU(ByRef inputVector() As Double, ByRef outputVector() As Double) As Boolean
  Dim rowNo As Long, position As Long, columnId As Long
  Dim rowValue As Double, diagValue As Double
  P6CSRApplyILU = False
  If Not P6CSRILUReady Then Exit Function
  For rowNo = 0 To lastDof
    rowValue = inputVector(rowNo)
    For position = P6CSRRowPtr(rowNo) To P6CSRRowPtr(rowNo + 1) - 1
      columnId = P6CSRColumnIndex(position)
      If columnId < rowNo Then rowValue = rowValue - P6CSRILUValues(position) * outputVector(columnId)
    Next position
    outputVector(rowNo) = rowValue
  Next rowNo
  For rowNo = lastDof To 0 Step -1
    rowValue = outputVector(rowNo)
    diagValue = 0#
    For position = P6CSRRowPtr(rowNo) To P6CSRRowPtr(rowNo + 1) - 1
      columnId = P6CSRColumnIndex(position)
      If columnId = rowNo Then diagValue = P6CSRILUValues(position)
      If columnId > rowNo Then rowValue = rowValue - P6CSRILUValues(position) * outputVector(columnId)
    Next position
    If Abs(diagValue) <= 0.000000000001 Then Exit Function
    outputVector(rowNo) = rowValue / diagValue
  Next rowNo
  P6CSRApplyILU = True
End Function

Private Function P6BuildCSRILU() As Boolean
  Dim rowNo As Long, position As Long, kPos As Long, updatePos As Long
  Dim columnId As Long, kCol As Long, updateCol As Long
  Dim diagK As Double, multiplier As Double
  P6BuildCSRILU = False
  P6CSRILUReady = False
  P6CSRILUFailRow = -1: P6CSRILUFailColumn = -1: P6CSRILUMinAbsDiag = 1E+308
  If P6CSRNNZ <= 0 Then Exit Function
  ReDim P6CSRILUValues(0 To P6CSRNNZ - 1)
  For position = 0 To P6CSRNNZ - 1
    P6CSRILUValues(position) = P6CSRValues(position)
  Next position
  For rowNo = 0 To lastDof
    For position = P6CSRRowPtr(rowNo) To P6CSRRowPtr(rowNo + 1) - 1
      columnId = P6CSRColumnIndex(position)
      If columnId < rowNo Then
        diagK = P6CSRILUValues(P6CSRDiagonalPosition(columnId))
        If Not P2IsFinite(diagK) Or Abs(diagK) <= 0.000000000001 Then
          P6CSRILUFailRow = rowNo: P6CSRILUFailColumn = columnId
          If P2IsFinite(diagK) Then P6CSRILUMinAbsDiag = Abs(diagK) Else P6CSRILUMinAbsDiag = 0#
          Exit Function
        End If
        multiplier = P6CSRILUValues(position) / diagK
        If Not P2IsFinite(multiplier) Then
          P6CSRILUFailRow = rowNo: P6CSRILUFailColumn = columnId: P6CSRILUMinAbsDiag = Abs(diagK)
          Exit Function
        End If
        P6CSRILUValues(position) = multiplier
        For kPos = P6CSRRowPtr(columnId) To P6CSRRowPtr(columnId + 1) - 1
          kCol = P6CSRColumnIndex(kPos)
          If kCol > columnId Then
            updatePos = P6CSRFindColumnPosition(rowNo, kCol)
            If updatePos >= 0 Then P6CSRILUValues(updatePos) = P6CSRILUValues(updatePos) - multiplier * P6CSRILUValues(kPos)
          End If
        Next kPos
      End If
    Next position
    diagK = P6CSRILUValues(P6CSRDiagonalPosition(rowNo))
    If Not P2IsFinite(diagK) Or Abs(diagK) <= 0.000000000001 Then
      P6CSRILUFailRow = rowNo: P6CSRILUFailColumn = rowNo
      If P2IsFinite(diagK) Then P6CSRILUMinAbsDiag = Abs(diagK) Else P6CSRILUMinAbsDiag = 0#
      Exit Function
    End If
    If Abs(diagK) < P6CSRILUMinAbsDiag Then P6CSRILUMinAbsDiag = Abs(diagK)
  Next rowNo
  If P6CSRILUMinAbsDiag = 1E+308 Then P6CSRILUMinAbsDiag = 0#
  P6CSRILUReady = True
  P6BuildCSRILU = True
End Function

Private Sub P6MixedApplyQp(ByRef pressure() As Double, ByRef outU() As Double)
  Dim elementId As Long, localDof As Long, cornerId As Long
  Dim pressureIndex As Long, nodeId As Long, uDof As Long
  For elementId = 0 To NumberOfElement - 1
    For cornerId = 0 To 3
      nodeId = Elem(elementId).node(cornerId)
      If nodeId < 0 Or nodeId >= NumberOfFreeNode Then GoTo NextQpCorner
      pressureIndex = P6NodePressure(nodeId)
      If pressureIndex < 0 Then GoTo NextQpCorner
      For localDof = 0 To 15
        uDof = Elem(elementId).ElNode(localDof)
        If NodeCond(uDof) = 0 Then
          outU(uDof) = outU(uDof) + P6MixedQ(elementId, localDof, cornerId) * pressure(pressureIndex)
        End If
      Next localDof
NextQpCorner:
    Next cornerId
  Next elementId
End Sub

Private Sub P6MixedApplyQTu(ByRef uVec() As Double, ByRef outP() As Double)
  Dim elementId As Long, localDof As Long, cornerId As Long
  Dim pressureIndex As Long, nodeId As Long, uDof As Long
  For pressureIndex = 0 To P6PressureCount - 1
    outP(pressureIndex) = 0#
  Next pressureIndex
  For elementId = 0 To NumberOfElement - 1
    For cornerId = 0 To 3
      nodeId = Elem(elementId).node(cornerId)
      If nodeId < 0 Or nodeId >= NumberOfFreeNode Then GoTo NextQtCorner
      pressureIndex = P6NodePressure(nodeId)
      If pressureIndex < 0 Then GoTo NextQtCorner
      For localDof = 0 To 15
        uDof = Elem(elementId).ElNode(localDof)
        outP(pressureIndex) = outP(pressureIndex) + P6MixedQ(elementId, localDof, cornerId) * uVec(uDof)
      Next localDof
NextQtCorner:
    Next cornerId
  Next elementId
End Sub

Private Function P6BuildDevBandFromCSR() As Boolean
  Dim rowNo As Long, position As Long, columnId As Long, offsetCol As Long, dofId As Long
  P6BuildDevBandFromCSR = False
  If lastDof < 0 Or BandWidth < 0 Or P6CSRNNZ <= 0 Then Exit Function
  ReDim TotalMat(lastDof, BandWidth)
  For rowNo = 0 To lastDof
    For position = P6CSRRowPtr(rowNo) To P6CSRRowPtr(rowNo + 1) - 1
      columnId = P6CSRColumnIndex(position)
      If columnId >= rowNo Then
        offsetCol = columnId - rowNo
        If offsetCol <= BandWidth Then TotalMat(rowNo, offsetCol) = P6CSROriginalValues(position)
      End If
    Next position
  Next rowNo
  For dofId = 0 To lastDof
    If NodeCond(dofId) <> 0 Then TotalMat(dofId, 0) = 1E+30
  Next dofId
  ReDim OriginalMat(lastDof, BandWidth)
  OriginalMat = TotalMat
  P6CopyBand TotalMat, P6FactoredBand
  P6BuildDevBandFromCSR = True
End Function

Private Function P6SolveMixedUzawa() As Boolean
  Dim i As Long, uzawaIter As Long, uzawaMax As Long
  Dim useDirectA As Boolean
  Dim normB As Double, residualNorm As Double, pressureNorm As Double, omega As Double
  Dim rhsValue As Double, bandWorkBytes As Double

  P6SolveMixedUzawa = False
  If P6PressureCount <= 0 Then Exit Function
  ReDim P6UzawaRhs(0 To lastDof)
  ReDim P6UzawaQP(0 To lastDof)
  ReDim P6UzawaRP(0 To P6PressureCount - 1)
  If P6PressureCount > 0 And (Not P6MixedUP) Then Exit Function
  For i = 0 To lastDof
    P6UzawaRhs(i) = Force(i)
  Next i
  For i = 0 To P6PressureCount - 1
    P6Pressure(i) = 0#
  Next i

  uzawaMax = CLng(P6ReadSetting("UZAWA_MAX", 40#))
  If uzawaMax < 4 Then uzawaMax = 4
  If uzawaMax > 200 Then uzawaMax = 200
  omega = 1#
  bandWorkBytes = CDbl(nDof) * (CDbl(BandWidth) + 1#) * 8# * 4#
  useDirectA = True
  If P6SolverMemoryLimitBytes > 0# And bandWorkBytes > P6SolverMemoryLimitBytes Then useDirectA = False
  If useDirectA Then
    If P6DevFactorGeneration <> P6TangentGeneration Or Not P6FactorReady Then
      If Not P6BuildDevBandFromCSR() Then
        useDirectA = False
      ElseIf Not BandLUFactor() Then
        useDirectA = False
        P6FactorReady = False
        ResultStatus = RESULT_NOT_RUN
        AnalysisOK = False
        AnalysisMessage = vbNullString
        AnalysisErrorNumber = 0
      Else
        P6DevFactorGeneration = P6TangentGeneration
      End If
    Else
      P6FactorizationReuseCount = P6FactorizationReuseCount + 1
    End If
  End If
  If Not useDirectA Then
    P6BuildCSRILU
    P6KrylovLast = lastDof
  End If
  If useDirectA Then
    P6SolverMode = "UZAWA_BAND"
  Else
    P6SolverMode = "UZAWA_ILU"
  End If

  normB = 0#
  For i = 0 To lastDof
    If NodeCond(i) = 0 Then normB = normB + P6UzawaRhs(i) * P6UzawaRhs(i)
  Next i
  If normB <= 0# Then normB = 1# Else normB = Sqr(normB)

  For uzawaIter = 1 To uzawaMax
    For i = 0 To lastDof
      P6UzawaQP(i) = 0#
    Next i
    P6MixedApplyQp P6Pressure, P6UzawaQP
    For i = 0 To lastDof
      If NodeCond(i) = 0 Then
        Force(i) = P6UzawaRhs(i) - P6UzawaQP(i)
      Else
        Force(i) = P6UzawaRhs(i)
      End If
    Next i

    If useDirectA Then
      If Not BandLUSolve() Then Exit Function
    Else
      If Not P6SolveCSRBICGSTAB() Then Exit Function
    End If

    P6MixedApplyQTu Disp, P6UzawaRP
    pressureNorm = 0#
    For i = 0 To P6PressureCount - 1
      rhsValue = P6UzawaRP(i) - P6MixedCompress(i) * P6Pressure(i)
      P6UzawaRP(i) = omega * P6MixedSchurInv(i) * rhsValue
      P6Pressure(i) = P6Pressure(i) + P6UzawaRP(i)
      pressureNorm = pressureNorm + rhsValue * rhsValue
    Next i
    If pressureNorm <= 0# Then pressureNorm = 0# Else pressureNorm = Sqr(pressureNorm)

    For i = 0 To lastDof
      P6UzawaQP(i) = 0#
    Next i
    P6MixedApplyQp P6UzawaRP, P6UzawaQP
    residualNorm = 0#
    For i = 0 To lastDof
      If NodeCond(i) = 0 Then residualNorm = residualNorm + P6UzawaQP(i) * P6UzawaQP(i)
    Next i
    residualNorm = Sqr(residualNorm + pressureNorm * pressureNorm)
    P6UzawaIterations = uzawaIter
    P6IterativeLastIterations = uzawaIter
    P6IterativeLastResidual = residualNorm / normB
    If residualNorm <= P6EngineeringAcceptLimit(normB) Then
      For i = 0 To lastDof
        P6UzawaQP(i) = 0#
      Next i
      P6MixedApplyQp P6Pressure, P6UzawaQP
      For i = 0 To lastDof
        If NodeCond(i) = 0 Then
          Force(i) = P6UzawaRhs(i) - P6UzawaQP(i)
        Else
          Force(i) = P6UzawaRhs(i)
        End If
      Next i
      If useDirectA Then
        If Not BandLUSolve() Then Exit Function
      Else
        If Not P6SolveCSRBICGSTAB() Then Exit Function
      End If
      P6IterativeFallbackStatus = "uzawa_" & P6SolverMode & "_iter=" & CStr(uzawaIter)
      P6SolveMixedUzawa = True
      Exit Function
    End If
    If uzawaIter = 4 And P6IterativeLastResidual > 0.1 Then omega = 0.5
  Next uzawaIter
  P6IterativeFallbackStatus = "uzawa_stalled_relres=" & Format$(P6IterativeLastResidual, "0.000E+00") & "; iter=" & CStr(uzawaMax)
End Function

Public Function P6SolveResidualRecovery(ByVal damping As Double) As Boolean
  ' Column-scaled Levenberg-Marquardt direction for the current free residual.
  ' A = S^-1 K^T K S^-1 + damping*I, rhs = S^-1 K^T r.
  ' The plastic operator, primary linear residual and committed state are unchanged.
  Dim a() As Double, rhs() As Double, solution() As Double, columnScale() As Double
  Dim row As Long, col As Long, p As Long, q As Long, c1 As Long, c2 As Long
  Dim i As Long, j As Long, k As Long, firstCol As Long, bw As Long, lastRow As Long
  Dim v As Double, normB As Double, normalResidual As Double, tmp As Double, startedAt As Double, factorDone As Boolean
  P6SolveResidualRecovery = False
  If Not P6CSRReady Or damping <= 0# Or Not P2IsFinite(damping) Then GoTo Rejected
  bw = 2 * BandWidth: If bw > lastDof Then bw = lastDof
  If P6SolverMemoryLimitBytes > 0# Then
    If P6CSRStorageBytes + CDbl(nDof) * (CDbl(bw + 1) * 8# + 32#) > P6SolverMemoryLimitBytes Then GoTo Rejected
  End If
  On Error GoTo Failed
  startedAt = Timer
  P6FactorizationCount = P6FactorizationCount + 1
  ReDim a(0 To lastDof, 0 To bw)
  ReDim rhs(0 To lastDof): ReDim solution(0 To lastDof): ReDim columnScale(0 To lastDof)
  For row = 0 To lastDof
    If NodeCond(row) = 0 Then
      If Not P2IsFinite(Force(row)) Then GoTo Rejected
      For p = P6CSRRowPtr(row) To P6CSRRowPtr(row + 1) - 1
        col = P6CSRColumnIndex(p)
        If NodeCond(col) = 0 Then
          v = P6CSROriginalValues(p)
          If Not P2IsFinite(v) Then GoTo Rejected
          columnScale(col) = columnScale(col) + v * v
        End If
      Next p
    End If
  Next row
  For col = 0 To lastDof
    columnScale(col) = Sqr(columnScale(col))
    If columnScale(col) < 1E-30 Then columnScale(col) = 1#
  Next col
  For row = 0 To lastDof
    If NodeCond(row) = 0 Then
      For p = P6CSRRowPtr(row) To P6CSRRowPtr(row + 1) - 1
        c1 = P6CSRColumnIndex(p)
        If NodeCond(c1) = 0 And P6CSROriginalValues(p) <> 0# Then
          v = P6CSROriginalValues(p) / columnScale(c1)
          rhs(c1) = rhs(c1) + v * Force(row)
          For q = p To P6CSRRowPtr(row + 1) - 1
            c2 = P6CSRColumnIndex(q)
            If NodeCond(c2) = 0 And P6CSROriginalValues(q) <> 0# Then
              If c2 < c1 Or c2 - c1 > bw Then GoTo Rejected
              a(c2, bw + c1 - c2) = a(c2, bw + c1 - c2) + v * P6CSROriginalValues(q) / columnScale(c2)
            End If
          Next q
        End If
      Next p
    End If
  Next row
  For i = 0 To lastDof
    a(i, bw) = a(i, bw) + damping
    If NodeCond(i) <> 0 Then a(i, bw) = 1#
    normB = normB + rhs(i) * rhs(i)
  Next i
  If Not P2IsFinite(normB) Or normB <= 1E-60 Then GoTo Rejected
  ' Cholesky in the lower band; positive damping makes the normal operator SPD.
  For i = 0 To lastDof
    firstCol = i - bw: If firstCol < 0 Then firstCol = 0
    For j = firstCol To i
      tmp = a(i, bw + j - i)
      For k = firstCol To j - 1
        If j - k <= bw Then tmp = tmp - a(i, bw + k - i) * a(j, bw + k - j)
      Next k
      If i = j Then
        If tmp <= 0# Or Not P2IsFinite(tmp) Then GoTo Rejected
        a(i, bw) = Sqr(tmp)
      Else
        a(i, bw + j - i) = tmp / a(j, bw)
      End If
    Next j
    tmp = rhs(i)
    For j = firstCol To i - 1: tmp = tmp - a(i, bw + j - i) * solution(j): Next j
    solution(i) = tmp / a(i, bw)
  Next i
  P6ProfFactorMs = P6ProfFactorMs + P6ElapsedMs(startedAt)
  factorDone = True: startedAt = Timer
  For i = lastDof To 0 Step -1
    tmp = solution(i): lastRow = i + bw: If lastRow > lastDof Then lastRow = lastDof
    For j = i + 1 To lastRow: tmp = tmp - a(j, bw + i - j) * solution(j): Next j
    solution(i) = tmp / a(i, bw)
    If Not P2IsFinite(solution(i)) Then GoTo Rejected
  Next i
  ' Independently apply K^T K + damping to verify the regularized solve.
  For row = 0 To lastDof
    If NodeCond(row) = 0 Then
      tmp = 0#
      For p = P6CSRRowPtr(row) To P6CSRRowPtr(row + 1) - 1
        col = P6CSRColumnIndex(p)
        If NodeCond(col) = 0 Then tmp = tmp + P6CSROriginalValues(p) * solution(col) / columnScale(col)
      Next p
      For p = P6CSRRowPtr(row) To P6CSRRowPtr(row + 1) - 1
        col = P6CSRColumnIndex(p)
        If NodeCond(col) = 0 Then rhs(col) = rhs(col) - P6CSROriginalValues(p) * tmp / columnScale(col)
      Next p
    End If
  Next row
  For i = 0 To lastDof
    tmp = rhs(i) - damping * solution(i)
    normalResidual = normalResidual + tmp * tmp
  Next i
  If Not P2IsFinite(normalResidual) Then GoTo Rejected
  If Sqr(normalResidual / normB) > P6IterativeTolerance Then GoTo Rejected
  For i = 0 To lastDof
    Disp(i) = solution(i) / columnScale(i)
    If NodeCond(i) <> 0 Then Disp(i) = 0#
  Next i
  P6SolverEvent "REGULARIZED_SOLVE", "damping=" & Format$(damping, "0.000E+00") & ";normal_relres=" & Format$(Sqr(normalResidual / normB), "0.000E+00")
  P6SolveResidualRecovery = True
  GoTo Rejected
Failed:
  Err.Clear
Rejected:
  If startedAt <> 0# Then
    If factorDone Then
      P6ProfSolveMs = P6ProfSolveMs + P6ElapsedMs(startedAt)
    Else
      P6ProfFactorMs = P6ProfFactorMs + P6ElapsedMs(startedAt)
    End If
  End If
End Function

Public Function P6SolveNonsymmetricBand() As Boolean
  ' Gaussian elimination with partial row pivoting in a full nonsymmetric band.
  ' Lower bandwidth is b; pivoting can extend the upper bandwidth to 2*b.
  ' Store chronological row exchanges and elimination multipliers for residual refinement.
  Dim a() As Double, rhs() As Double, solution() As Double, residualVector() As Double, refineRhs() As Double, correction() As Double
  Dim pivotRows() As Long, lowerMult() As Double, refineIter As Long, reused As Boolean
  Dim b As Long, upper As Long, k As Long, i As Long, j As Long, p As Long
  Dim pivotRow As Long, lastRow As Long, lastColumn As Long, columnId As Long
  Dim pivot As Double, largest As Double, multiplier As Double, tmp As Double
  Dim normB As Double, residual As Double, memoryBytes As Double, startedAt As Double, factorDone As Boolean
  P6SolveNonsymmetricBand = False
  P6BandRejectReason = "INPUT"
  P6IterativeLastResidual = 1E+100
  If nDof < 1 Or lastDof <> nDof - 1 Or BandWidth < 0 Then GoTo Rejected
  b = BandWidth: If b > lastDof Then b = lastDof
  upper = 2 * b
  reused = P6BandCSRCacheMatches(b)
  If Not reused Then P6ClearElasticRecoveryCSR
  ' Both the retained factors/CSR identity and this call's work arrays count.
  memoryBytes = CDbl(nDof) * (CDbl(8 * b + 6) * 8# + 80#) + CDbl(P6CSRNNZ) * 12# + P6CSRStorageBytes
  If P6SolverMemoryLimitBytes > 0# And memoryBytes > P6SolverMemoryLimitBytes Then
    SetAnalysisFailure RESULT_CAPACITY_ERROR, "非対称帯域LUの作業メモリが設定上限を超えます。", vbObjectError + 3511, -1, -1, CurrentIncrement, CurrentIteration
    Exit Function
  End If
  On Error GoTo Failed
  startedAt = Timer
  If reused Then
    a = mBandCacheA: pivotRows = mBandCachePivots: lowerMult = mBandCacheLower
    P6FactorizationReuseCount = P6FactorizationReuseCount + 1
  Else
    P6FactorizationCount = P6FactorizationCount + 1
    ReDim a(0 To lastDof, 0 To 3 * b)
    ReDim pivotRows(0 To lastDof): ReDim lowerMult(0 To lastDof, 0 To b)
  End If
  ReDim rhs(0 To lastDof): ReDim solution(0 To lastDof): ReDim residualVector(0 To lastDof)
  For i = 0 To lastDof
    If Not P2IsFinite(Force(i)) Then P6BandRejectReason = "NONFINITE_RHS": GoTo Rejected
    rhs(i) = Force(i): normB = normB + rhs(i) * rhs(i)
    If Not reused Then
    For p = P6CSRRowPtr(i) To P6CSRRowPtr(i + 1) - 1
      columnId = P6CSRColumnIndex(p)
      If Not P2IsFinite(P6CSRValues(p)) Then GoTo Rejected
      If Abs(columnId - i) > b Then
        If Abs(P6CSRValues(p)) > 1E-30 Then GoTo Rejected
      Else
        a(i, columnId - i + b) = a(i, columnId - i + b) + P6CSRValues(p)
      End If
    Next p
    End If
  Next i
  normB = Sqr(normB)
  If reused Then
    P6BandReplayForward rhs, pivotRows, lowerMult, b
  Else
  For k = 0 To lastDof
    lastRow = k + b: If lastRow > lastDof Then lastRow = lastDof
    lastColumn = k + upper: If lastColumn > lastDof Then lastColumn = lastDof
    pivotRow = k: largest = Abs(a(k, b))
    For i = k + 1 To lastRow
      tmp = Abs(a(i, k - i + b))
      If tmp > largest Then largest = tmp: pivotRow = i
    Next i
    If Not P2IsFinite(largest) Or largest <= 1E-30 Then P6BandRejectReason = "ZERO_PIVOT": GoTo Rejected
    pivotRows(k) = pivotRow
    If pivotRow <> k Then
      For j = k To lastColumn
        tmp = a(k, j - k + b)
        a(k, j - k + b) = a(pivotRow, j - pivotRow + b)
        a(pivotRow, j - pivotRow + b) = tmp
      Next j
      tmp = rhs(k): rhs(k) = rhs(pivotRow): rhs(pivotRow) = tmp
    End If
    pivot = a(k, b)
    For i = k + 1 To lastRow
      multiplier = a(i, k - i + b) / pivot
      If Not P2IsFinite(multiplier) Then P6BandRejectReason = "NONFINITE_MULTIPLIER": GoTo Rejected
      a(i, k - i + b) = 0#
      lowerMult(i, i - k) = multiplier
      If multiplier <> 0# Then
        For j = k + 1 To lastColumn
          a(i, j - i + b) = a(i, j - i + b) - multiplier * a(k, j - k + b)
        Next j
        rhs(i) = rhs(i) - multiplier * rhs(k)
      End If
    Next i
  Next k
    mBandCacheA = a: mBandCachePivots = pivotRows: mBandCacheLower = lowerMult
    mBandCacheRows = P6CSRRowPtr: mBandCacheColumns = P6CSRColumnIndex: mBandCacheValues = P6CSRValues
    mBandCacheN = nDof: mBandCacheBw = b: mBandCacheReady = True
  End If
  P6ProfFactorMs = P6ProfFactorMs + P6ElapsedMs(startedAt)
  factorDone = True: startedAt = Timer
  For i = lastDof To 0 Step -1
    tmp = rhs(i): lastColumn = i + upper: If lastColumn > lastDof Then lastColumn = lastDof
    For j = i + 1 To lastColumn: tmp = tmp - a(i, j - i + b) * solution(j): Next j
    solution(i) = tmp / a(i, b)
    If Not P2IsFinite(solution(i)) Then GoTo Rejected
  Next i
  residual = P6CSRMatVecResidualNorm(solution, Force, residualVector, P6CSRValues)
  If Not P2IsFinite(residual) Or Not P2IsFinite(normB) Then P6BandRejectReason = "NONFINITE_RESIDUAL": GoTo Rejected
  If normB > 1E-30 Then
    P6IterativeLastResidual = residual / normB
    If P6IterativeLastResidual > P6IterativeTolerance Then
      P6SolverEvent "NONSYM_BAND_REFINEMENT_START", "relative_residual=" & Format$(P6IterativeLastResidual, "0.000E+00")
      ReDim refineRhs(0 To lastDof): ReDim correction(0 To lastDof)
      For refineIter = 1 To 5
        For i = 0 To lastDof: refineRhs(i) = residualVector(i): correction(i) = 0#: Next i
        For k = 0 To lastDof
          pivotRow = pivotRows(k)
          If pivotRow <> k Then tmp = refineRhs(k): refineRhs(k) = refineRhs(pivotRow): refineRhs(pivotRow) = tmp
          lastRow = k + b: If lastRow > lastDof Then lastRow = lastDof
          For i = k + 1 To lastRow
            refineRhs(i) = refineRhs(i) - lowerMult(i, i - k) * refineRhs(k)
          Next i
        Next k
        For i = lastDof To 0 Step -1
          tmp = refineRhs(i): lastColumn = i + upper: If lastColumn > lastDof Then lastColumn = lastDof
          For j = i + 1 To lastColumn: tmp = tmp - a(i, j - i + b) * correction(j): Next j
          If Abs(a(i, b)) <= 1E-30 Then GoTo Rejected
          correction(i) = tmp / a(i, b)
        Next i
        For i = 0 To lastDof
          solution(i) = solution(i) + correction(i)
          If Not P2IsFinite(solution(i)) Then GoTo Rejected
        Next i
        residual = P6CSRMatVecResidualNorm(solution, Force, residualVector, P6CSRValues)
        If Not P2IsFinite(residual) Then GoTo Rejected
        P6IterativeLastResidual = residual / normB
        P6SolverEvent "NONSYM_BAND_REFINEMENT", "iteration=" & CStr(refineIter) & ";relative_residual=" & Format$(P6IterativeLastResidual, "0.000E+00")
        If P6IterativeLastResidual <= P6IterativeTolerance Then Exit For
      Next refineIter
      If P6IterativeLastResidual > P6IterativeTolerance Then P6BandRejectReason = "RESIDUAL": GoTo Rejected
    End If
  Else
    P6IterativeLastResidual = residual
    If residual > 1E-30 Then GoTo Rejected
  End If
  For i = 0 To lastDof: Disp(i) = solution(i): Next i
  P6IterativeLastIterations = 0
  P6SolveNonsymmetricBand = True
  GoTo Rejected
Failed:
  If Err.Number = 7 Then SetAnalysisFailure RESULT_CAPACITY_ERROR, "非対称帯域LUのメモリ確保に失敗しました。", Err.Number, -1, -1, CurrentIncrement, CurrentIteration
  Err.Clear
Rejected:
  If Not P6SolveNonsymmetricBand Then
    P6ClearElasticRecoveryCSR
    P6SolverEvent "NONSYM_BAND_REJECT", "pivot_index=" & CStr(k) & ";reason=" & P6BandRejectReason & ";pivot=" & Format$(pivot, "0.000E+00") & ";pivot_scale=" & Format$(largest, "0.000E+00") & ";bandwidth=" & CStr(b) & ";norm_rhs=" & Format$(normB, "0.000E+00") & ";relative_residual=" & Format$(P6IterativeLastResidual, "0.000E+00")
  End If
  If startedAt <> 0# Then
    If factorDone Then
      P6ProfSolveMs = P6ProfSolveMs + P6ElapsedMs(startedAt)
    Else
      P6ProfFactorMs = P6ProfFactorMs + P6ElapsedMs(startedAt)
    End If
  End If
End Function

' Elastic correction API. The caller may use this bounded path without
' changing the nonlinear solver's existing fallback policy.
Public Function P6SolveElasticRecoveryCSR() As Boolean
  P6SolveElasticRecoveryCSR = P6SolveNonsymmetricBand()
End Function

Public Sub P6ClearElasticRecoveryCSR()
  mBandCacheReady = False: mBandCacheN = 0: mBandCacheBw = 0
  Erase mBandCacheA, mBandCacheLower, mBandCachePivots, mBandCacheRows, mBandCacheColumns, mBandCacheValues
End Sub

Private Function P6BandCSRCacheMatches(ByVal b As Long) As Boolean
  Dim i As Long
  P6BandCSRCacheMatches = False
  If Not mBandCacheReady Or mBandCacheN <> nDof Or mBandCacheBw <> b Then Exit Function
  On Error GoTo NotMatching
  If UBound(mBandCacheValues) <> P6CSRNNZ - 1 Then Exit Function
  For i = 0 To nDof
    If mBandCacheRows(i) <> P6CSRRowPtr(i) Then Exit Function
  Next i
  For i = 0 To P6CSRNNZ - 1
    If mBandCacheColumns(i) <> P6CSRColumnIndex(i) Then Exit Function
    If mBandCacheValues(i) <> P6CSRValues(i) Then Exit Function
  Next i
  P6BandCSRCacheMatches = True
NotMatching:
End Function

Private Sub P6BandReplayForward(ByRef rhs() As Double, ByRef pivots() As Long, ByRef multipliers() As Double, ByVal b As Long)
  Dim k As Long, i As Long, p As Long, lastRow As Long, tmp As Double
  For k = 0 To lastDof
    p = pivots(k)
    If p <> k Then tmp = rhs(k): rhs(k) = rhs(p): rhs(p) = tmp
    lastRow = k + b: If lastRow > lastDof Then lastRow = lastDof
    For i = k + 1 To lastRow
      rhs(i) = rhs(i) - multipliers(i, i - k) * rhs(k)
    Next i
  Next k
End Sub

Private Function P6SolveCSR() As Boolean
  Dim bicgIterations As Long, bicgResidual As Double

  P6SolveCSR = False
  If Not P6CSRReady Then
    SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "CSR疎行列が準備されていません。", vbObjectError + 3301, -1, -1, CurrentIncrement, CurrentIteration
    Exit Function
  End If
  If Not P6PrepareMixedUP() Then
    SetAnalysisFailure RESULT_GLOBAL_SINGULAR, "u-p混合の準備に失敗しました。", vbObjectError + 3307, -1, -1, CurrentIncrement, CurrentIteration
    Exit Function
  End If
  If P6MixedUP Then
    P6SolveCSR = P6SolveMixedUzawa()
    Exit Function
  End If
  If P3FlowPolicyIsInconsistent() Then
    If P6SolveNonsymmetricBand() Then
      P6SolveCSR = True
      Exit Function
    End If
    If ResultStatus = RESULT_CAPACITY_ERROR Then Exit Function
    P6IterativeFallbackStatus = "right_gmres_after_nonsymmetric_band"
    P6SolverMode = "CSR_GMRES"
  End If
  If P3FlowPolicyIsInconsistent() And Not P6CSRILUReady Then
    ' ILU(0) is only a preconditioner; convergence uses the true CSR residual.
    If P6SolverMemoryLimitBytes > 0# Then
      If P6CSRStorageBytes + CDbl(P6CSRNNZ) * 8# + CDbl(nDof) * (23# + 2# * P6GMRESRestart) * 8# + CDbl(P6GMRESRestart + 1) ^ 2 * 8# > P6SolverMemoryLimitBytes Then
        SetAnalysisFailure RESULT_CAPACITY_ERROR, "非対称GMRESの作業メモリが設定上限を超えます。", vbObjectError + 3511, -1, -1, CurrentIncrement, CurrentIteration
        Exit Function
      End If
    End If
    If Not P6BuildCSRILU() Then
      P6CSRPreconditionerMode = "JACOBI_AFTER_ILU_FAILURE"
      P6SolverEvent "CSR_ILU", "status=FAIL;row=" & CStr(P6CSRILUFailRow) & ";column=" & CStr(P6CSRILUFailColumn) & ";min_abs_diag=" & Format$(P6CSRILUMinAbsDiag, "0.000E+00") & ";fallback=JACOBI"
    Else
      P6CSRPreconditionerMode = "ILU0"
      P6SolverEvent "CSR_ILU", "status=PASS;min_abs_diag=" & Format$(P6CSRILUMinAbsDiag, "0.000E+00")
    End If
  End If
  If P6SolverMode = "CSR_GMRES" Then
    P6SolveCSR = P6SolveCSRGMRES()
    If Not P6SolveCSR Then
      SetAnalysisFailure RESULT_NONCONVERGED, "CSR-GMRESが収束しませんでした。反復=" & CStr(P6IterativeLastIterations) & "、相対残差=" & Format$(P6IterativeLastResidual, "0.000E+00"), vbObjectError + 3202, -1, -1, CurrentIncrement, CurrentIteration
    End If
    Exit Function
  End If

  P6SolveCSR = P6SolveCSRBICGSTAB()
  If P6SolveCSR Then Exit Function
  bicgIterations = P6IterativeLastIterations
  bicgResidual = P6IterativeLastResidual
  P6IterativeFallbackStatus = "gmres_after_bicgstab"
  P6SolveCSR = P6SolveCSRGMRES()
  If P6SolveCSR Then Exit Function
  SetAnalysisFailure RESULT_NONCONVERGED, "CSR-BiCGSTABが収束しませんでした。Ver=" & FEM_BUILD_STAMP & " 反復=" & CStr(bicgIterations) & "、相対残差=" & Format$(bicgResidual, "0.000E+00") & "。GMRES再試行も失敗 反復=" & CStr(P6IterativeLastIterations) & "、相対残差=" & Format$(P6IterativeLastResidual, "0.000E+00"), vbObjectError + 3202, -1, -1, CurrentIncrement, CurrentIteration
End Function

Private Function P6RebuildBandFromElements() As Boolean
  Dim i As Long, j As Long, k As Long, ii As Long, jj As Long
  P6RebuildBandFromElements = False
  If lastDof < 0 Or BandWidth < 0 Then Exit Function
  ReDim OriginalMat(lastDof, BandWidth)
  For k = 0 To NumberOfElement - 1
    If Not P3IsElementActive(k) Then GoTo NextRebuildElement
    For i = 0 To 15
      If Elem(k).ElNode(i) < 0 Then GoTo NextRebuildI
      For j = 0 To 15
        If Elem(k).ElNode(j) < 0 Then GoTo NextRebuildJ
        ii = Elem(k).ElNode(i)
        jj = Elem(k).ElNode(j) - ii
        If jj >= 0 Then OriginalMat(ii, jj) = OriginalMat(ii, jj) + Elem(k).kmat(i, j)
NextRebuildJ:
      Next j
NextRebuildI:
    Next i
NextRebuildElement:
  Next k
  P6RebuildBandFromElements = True
End Function

Private Function P6TryBandDirectFallback() As Boolean
  Dim i As Long
  Dim savedResidual As Double, savedIterations As Long
  Dim bandWorkBytes As Double

  P6TryBandDirectFallback = False
  savedResidual = P6IterativeLastResidual
  savedIterations = P6IterativeLastIterations
  bandWorkBytes = CDbl(nDof) * (CDbl(BandWidth) + 1#) * 8# * 4#
  If P6SolverMemoryLimitBytes > 0# And bandWorkBytes > P6SolverMemoryLimitBytes Then Exit Function
  If Not P6RebuildBandFromElements() Then Exit Function

  P6UseCSR = False
  P6SolverMode = "BAND_LU"
  P6FactorReady = False
  MatrixFactored = False
  For i = 0 To lastDof
    If P6ForceBeforeBoundaryReady Then
      Force(i) = P6ForceBeforeBoundary(i)
    End If
    If NodeCond(i) <> 0 Then
      Force(i) = 1E+30 * Disp(i)
    End If
  Next i
  P6CopyBand OriginalMat, P6FactoredBand
  For i = 0 To lastDof
    If NodeCond(i) <> 0 Then P6FactoredBand(i, 0) = 1E+30
  Next i

  AnalysisOK = False
  ResultStatus = RESULT_NOT_RUN
  AnalysisMessage = vbNullString
  AnalysisErrorNumber = 0

  If Not BandLUFactor() Then Exit Function
  If Not BandLUSolve() Then Exit Function

  P6BandAfterIterative = True
  P6IterativeFallbackStatus = "band_after_iterative_relres=" & Format$(savedResidual, "0.000E+00") & "; iter=" & CStr(savedIterations)
  P6MatrixStoragePolicy = "帯域格納＋分解済み帯域因子; mode=BAND_LU; AUTO切替（反復残差>" & Format$(P6_ITERATIVE_BAND_FALLBACK, "0.0E+00") & "）"
  P6TryBandDirectFallback = True
End Function


Private Function P6LdltShadowDue() As Boolean
  Dim n As Long
  Dim modeText As String
  ' Production solve stays band LU. LDLT remains an opt-in shadow only.
  P6LdltShadowDue = False
  modeText = UCase$(Trim$(P6ReadTextSetting("SOLVER", "BAND_LU")))
  If modeText <> "BAND_LDLT_EXPERIMENTAL" Then Exit Function
  n = P6FactorizationCount + 1
  P6LdltShadowDue = (n <= 20) Or ((n Mod 100) = 0)
End Function

Private Sub P6EnsureLdltWorkspace()
  If P6LdltDim <> lastDof Or P6LdltBw <> BandWidth Then
    ReDim P6LdltBand(lastDof, BandWidth)
    ReDim P6LdltHold(lastDof, BandWidth)
    ReDim P6LdltRhs(0 To lastDof)
    ReDim P6LdltX(0 To lastDof)
    ReDim P6LdltAx(0 To lastDof)
    ReDim P6LdltMult(0 To BandWidth)
    P6LdltDim = lastDof
    P6LdltBw = BandWidth
  End If
End Sub

Private Sub P6CaptureLdltRhs()
  Dim i As Long
  P6EnsureLdltWorkspace
  For i = 0 To lastDof
    P6LdltRhs(i) = Force(i)
  Next i
  P6LdltRhsReady = True
End Sub

Private Sub P6SnapshotLdltHold()
  Dim i As Long, j As Long
  P6EnsureLdltWorkspace
  For i = 0 To lastDof
    For j = 0 To BandWidth
      P6LdltHold(i, j) = P6FactoredBand(i, j)
    Next j
  Next i
  P6LdltSnapReady = True
End Sub

Private Sub P6CopyHoldToLdltBand()
  Dim i As Long, j As Long
  For i = 0 To lastDof
    For j = 0 To BandWidth
      P6LdltBand(i, j) = P6LdltHold(i, j)
    Next j
  Next i
End Sub

Private Function P6TryBandLDLT(ByRef dMin As Double, ByRef dMax As Double, ByRef failIndex As Long, ByRef failD As Double, ByRef usedTol As Double) As Long
  Dim k As Long, i As Long, j As Long, it As Long, jt As Long, m As Long, n2 As Long
  Dim d As Double, xscale As Double, ad As Double
  ' rcode 0 = OK, 2 = NONPOSITIVE_D. Never test this with If P6TryBandLDLT Then (0 is False).
  P6TryBandLDLT = 2
  failIndex = -1
  failD = 0#
  usedTol = 0#
  n2 = nDof
  dMin = 1E+308
  dMax = 0#
  xscale = 0#
  For i = 0 To n2 - 1
    If NodeCond(i) = 0 Then
      If Abs(P6LdltBand(i, 0)) > xscale Then xscale = Abs(P6LdltBand(i, 0))
    End If
  Next i
  If xscale < 1# Then xscale = 1#
  usedTol = 0.000000000001 * xscale
  For k = 0 To n2 - 1
    d = P6LdltBand(k, 0)
    If d <= usedTol Then
      failIndex = k
      failD = d
      P6TryBandLDLT = 2
      Exit Function
    End If
    If NodeCond(k) = 0 Then
      ad = Abs(d)
      If ad < dMin Then dMin = ad
      If ad > dMax Then dMax = ad
    End If
    m = BandWidth
    If k + m > n2 - 1 Then m = n2 - 1 - k
    For j = 1 To m
      P6LdltMult(j) = P6LdltBand(k, j) / d
      P6LdltBand(k, j) = P6LdltMult(j)
    Next j
    For i = 1 To m
      it = k + i
      For j = i To m
        jt = j - i
        P6LdltBand(it, jt) = P6LdltBand(it, jt) - d * P6LdltMult(i) * P6LdltMult(j)
      Next j
    Next i
  Next k
  failIndex = -1
  failD = 0#
  P6TryBandLDLT = 0
End Function

Private Sub P6SolveBandLDLT()
  Dim k As Long, j As Long, m As Long, n2 As Long
  Dim s As Double
  n2 = nDof
  For k = 0 To n2 - 1
    P6LdltX(k) = P6LdltRhs(k)
  Next k
  For k = 0 To n2 - 2
    m = BandWidth
    If k + m > n2 - 1 Then m = n2 - 1 - k
    For j = 1 To m
      P6LdltX(k + j) = P6LdltX(k + j) - P6LdltBand(k, j) * P6LdltX(k)
    Next j
  Next k
  For k = 0 To n2 - 1
    P6LdltX(k) = P6LdltX(k) / P6LdltBand(k, 0)
  Next k
  For k = n2 - 1 To 0 Step -1
    s = P6LdltX(k)
    m = BandWidth
    If k + m > n2 - 1 Then m = n2 - 1 - k
    For j = 1 To m
      s = s - P6LdltBand(k, j) * P6LdltX(k + j)
    Next j
    P6LdltX(k) = s
  Next k
End Sub

Private Sub P6BandMatVecUnfactored(ByRef x() As Double, ByRef y() As Double)
  Dim i As Long, j As Long, m As Long, n2 As Long
  Dim a As Double, diag As Double
  n2 = nDof
  For i = 0 To n2 - 1
    y(i) = 0#
  Next i
  For i = 0 To n2 - 1
    diag = P6LdltHold(i, 0)
    y(i) = y(i) + diag * x(i)
    m = BandWidth
    If i + m > n2 - 1 Then m = n2 - 1 - i
    For j = 1 To m
      a = P6LdltHold(i, j)
      y(i) = y(i) + a * x(i + j)
      y(i + j) = y(i + j) + a * x(i)
    Next j
  Next i
End Sub

Private Function P6LdltNum(ByVal v As Double) As String
  If v > 1E+307 Or v < -1E+307 Then
    P6LdltNum = "NA"
  Else
    P6LdltNum = Format$(v, "0.000E+00")
  End If
End Function

Private Sub P6ShadowBandLDLT()
  Dim t0 As Double, mS As Double
  Dim rcode As Long, failIndex As Long, i As Long
  Dim dMin As Double, dMax As Double, dRatio As Double, failD As Double, usedTol As Double
  Dim num As Double, den As Double, ex As Double, er As Double, diff As Double
  Dim detailText As String
  Dim counted As Boolean
  On Error GoTo ShadowFail
  counted = False
  If Not P6LdltRhsReady Then Exit Sub
  If Not P6LdltSnapReady Then Exit Sub
  If lastDof < 0 Or BandWidth < 0 Then Exit Sub
  P6EnsureLdltWorkspace
  P6CopyHoldToLdltBand
  t0 = Timer
  rcode = P6TryBandLDLT(dMin, dMax, failIndex, failD, usedTol)
  mS = P6ElapsedMs(t0)
  counted = True
  P6LdltN = P6LdltN + 1
  P6LdltMs = P6LdltMs + mS
  If rcode = 0 Then
    P6SolveBandLDLT
    num = 0#
    den = 0#
    For i = 0 To lastDof
      diff = P6LdltX(i) - Disp(i)
      num = num + diff * diff
      den = den + Disp(i) * Disp(i)
    Next i
    ex = Sqr(num) / (Sqr(den) + 1E-30)
    P6BandMatVecUnfactored P6LdltX, P6LdltAx
    num = 0#
    den = 0#
    For i = 0 To lastDof
      diff = P6LdltAx(i) - P6LdltRhs(i)
      num = num + diff * diff
      den = den + P6LdltRhs(i) * P6LdltRhs(i)
    Next i
    er = Sqr(num) / (Sqr(den) + 1E-30)
    If dMax > 0# Then
      dRatio = dMin / dMax
    Else
      dRatio = 0#
    End If
    If ex > P6LdltExMax Then P6LdltExMax = ex
    If er > P6LdltErMax Then P6LdltErMax = er
    If ex > P6LdltExMaxTrial Then P6LdltExMaxTrial = ex
    If er > P6LdltErMaxTrial Then P6LdltErMaxTrial = er
    If dRatio < P6LdltDRatioMin Then P6LdltDRatioMin = dRatio
    If dRatio < P6LdltDRatioMinTrial Then P6LdltDRatioMinTrial = dRatio
    P6LdltOk = P6LdltOk + 1
    detailText = "status=OK;rcode=0;fail_index=-1;d=NA;tol=" & P6LdltNum(usedTol)
    detailText = detailText & ";dmin=" & P6LdltNum(dMin) & ";dmax=" & P6LdltNum(dMax) & ";dratio=" & P6LdltNum(dRatio)
    detailText = detailText & ";ex=" & P6LdltNum(ex) & ";er=" & P6LdltNum(er)
    detailText = detailText & ";ms=" & Format$(mS, "0.000") & ";factor=" & CStr(P6FactorizationCount)
    P6SolverEvent "LDLT_SHADOW", detailText
    Exit Sub
  End If
  P6LdltFail = P6LdltFail + 1
  If dMax > 0# And dMin < 1E+307 Then
    dRatio = dMin / dMax
  Else
    dRatio = 1E+308
  End If
  detailText = "status=NONPOSITIVE_D;rcode=" & CStr(rcode) & ";fail_index=" & CStr(failIndex)
  detailText = detailText & ";d=" & P6LdltNum(failD) & ";tol=" & P6LdltNum(usedTol)
  detailText = detailText & ";dmin=" & P6LdltNum(dMin) & ";dmax=" & P6LdltNum(dMax) & ";dratio=" & P6LdltNum(dRatio)
  detailText = detailText & ";ex=NA;er=NA;ms=" & Format$(mS, "0.000") & ";factor=" & CStr(P6FactorizationCount)
  P6SolverEvent "LDLT_SHADOW", detailText
  Exit Sub
ShadowFail:
  If Not counted Then
    P6LdltN = P6LdltN + 1
    P6LdltFail = P6LdltFail + 1
  End If
  P6SolverEvent "LDLT_SHADOW", "status=ERR;rcode=3;fail_index=-1;d=NA;tol=NA;dmin=NA;dmax=NA;dratio=NA;ex=NA;er=NA;factor=" & CStr(P6FactorizationCount)
  Err.Clear
End Sub

Function BandSolver() As Boolean
  Dim csrOk As Boolean
  Dim ldltShadow As Boolean
  P6BandSolveCallCount = P6BandSolveCallCount + 1
  BandSolver = False
  MatrixFactored = False
  ldltShadow = False
  If P6UseCSR Then
    csrOk = P6SolveCSR()
    If P3FlowPolicyIsInconsistent() Then
      If csrOk Then MatrixFactored = True: BandSolver = True
      Exit Function
    End If
    If csrOk Then
      If P6IterativeLastResidual <= P6_ITERATIVE_BAND_FALLBACK Then
        MatrixFactored = True
        BandSolver = True
        Exit Function
      End If
    End If
    If P3FlowPolicyIsInconsistent() Then Exit Function
    If P6TryBandDirectFallback() Then
      MatrixFactored = True
      BandSolver = True
      Exit Function
    End If
    If Not csrOk Then Exit Function
    SetAnalysisFailure RESULT_NONCONVERGED, "反復法の相対残差=" & Format$(P6IterativeLastResidual, "0.000E+00") & " が " & Format$(P6_ITERATIVE_BAND_FALLBACK, "0.0E+00") & " を超え、BAND切替も失敗しました。Ver=" & FEM_BUILD_STAMP, vbObjectError + 3202, -1, -1, CurrentIncrement, CurrentIteration
    Exit Function
  End If
  If Not P6FactorReady Then
    If P6AccelTryGMRES() Then
      MatrixFactored = False ' No exact fresh LU exists; never publish false factor validity.
      BandSolver = True
      Exit Function
    End If
    If Not BandLUFactor() Then Exit Function
    ldltShadow = P6LdltShadowDue()
  Else
    P6FactorizationReuseCount = P6FactorizationReuseCount + 1
  End If
  If ldltShadow Then P6CaptureLdltRhs
  If Not BandLUSolve() Then Exit Function
  If ldltShadow Then P6ShadowBandLDLT
  If P6MixedUP Then P6RecoverPressureIncrement
  MatrixFactored = True
  BandSolver = True
End Function
Private Sub P6RememberFactorKeys()
  Dim k As Long, gp As Long, nJoint As Long, cursor As Long, maskValue As Long, bitValue As Long
  AccelCaptureFactorTangent
  P6AccelCapturePreconditioner
  P6FactorActivePlastic = P3ActivePlasticPointCount
  P6FactorFs = P3CurrentStrengthFactor
  P6FactorActiveGen = P3ActiveSetGen
  P6FactorConstraintGen = P3ConstraintGeneration
  P6FactorJointContactN = 0
  If NumberOfElement > 0 Then
    If P6FactorYieldMaskN <> NumberOfElement Then ReDim P6FactorYieldMask(0 To NumberOfElement - 1)
    For k = 0 To NumberOfElement - 1
      maskValue = 0
      bitValue = 1
      For gp = 0 To 3
        If Elem(k).P2Yielded(gp) Then maskValue = maskValue + bitValue
        bitValue = bitValue * 2
      Next gp
      P6FactorYieldMask(k) = CByte(maskValue)
    Next k
    P6FactorYieldMaskN = NumberOfElement
  Else
    Erase P6FactorYieldMask
    P6FactorYieldMaskN = 0
  End If
  If (Not P3HasJointElements) Or NumberOfElement < 1 Then
    Erase P6FactorJointContact
    Erase P6FactorJointElem
  Else
    nJoint = 0
    For k = 0 To NumberOfElement - 1
      If Elem(k).IsJoint Then nJoint = nJoint + 1
    Next k
    If nJoint < 1 Then
      Erase P6FactorJointContact
      Erase P6FactorJointElem
    Else
      ReDim P6FactorJointElem(0 To nJoint - 1)
      ReDim P6FactorJointContact(0 To 2, 0 To nJoint - 1)
      cursor = 0
      For k = 0 To NumberOfElement - 1
        If Elem(k).IsJoint Then
          P6FactorJointElem(cursor) = k
          For gp = 0 To 2
            P6FactorJointContact(gp, cursor) = Elem(k).JointContact(gp)
          Next gp
          cursor = cursor + 1
        End If
      Next k
      P6FactorJointContactN = nJoint
    End If
  End If
  If lastDof < 0 Then
    P6FactorNodeCondN = -1
    Exit Sub
  End If
  P6FactorNodeCondN = lastDof
End Sub

Public Function P6CanReuseAssembledFactor() As Boolean
  P6CanReuseAssembledFactor = False
  If Not P6FactorReady Then Exit Function
  If P6MixedUP Then Exit Function
  If Abs(P3CurrentStrengthFactor - P6FactorFs) > 0.000000000001 Then Exit Function
  If P3ActiveSetGen <> P6FactorActiveGen Then Exit Function
  If P3ConstraintGeneration <> P6FactorConstraintGen Then Exit Function
  If P6FactorNodeCondN <> lastDof Then Exit Function
  If lastDof < 0 Then Exit Function
  P6CanReuseAssembledFactor = True
End Function



Public Sub P6AccelResetPreconditioner()
  mPreReady = False: mPreUses = 0
End Sub

Private Sub P6AccelCapturePreconditioner()
  On Error GoTo CapturePreSkip
  mPreReady = False
  If Not AccelV5 Or AccelBaselineRetry Or P6MixedUP Or P3HasJointElements Or P6UseCSR Then Exit Sub
  If Not P6BandKernelIsNew() Then Exit Sub
  If P6BandStorageBytes * 5# + 16# * nDof * (P6GMRESRestart + 1#) > P6SolverMemoryLimitBytes Then Exit Sub
  mPrePacked = P6Packed: mPreOffsets = P6PackRowOff
  mPreN = nDof: mPreBw = BandWidth: mPreFs = P3CurrentStrengthFactor
  mPreActive = P3ActiveSetGen: mPreConstraint = P3ConstraintGeneration
  mPreMaterial = P6MaterialGeneration: mPreStrength = P6StrengthGeneration: mPreGeometry = P6GeometryRevision
  mPreTol = P6FactorPivotTolerance: mPreUses = 0: mPreReady = True
  Exit Sub
CapturePreSkip:
  mPreReady = False
  If Err.Number = 18 Then Err.Raise 18
  Err.Clear
End Sub

Private Function P6AccelGMRESCore(ByRef mat() As Double, ByRef packed() As Double, ByRef rowOff() As Long, ByRef rhs() As Double, ByRef x() As Double, ByRef fixed() As Boolean, ByVal n As Long, ByVal bw As Long, ByVal tol As Double, ByVal pivotTol As Double, ByVal restart As Long, ByVal maxIters As Long, ByRef relres As Double, ByRef usedIters As Long, Optional ByVal maxCostMs As Double = 0#) As Boolean
  Dim costStart As Double
  Dim v() As Double, zBasis() As Double, h() As Double, cs() As Double, sn() As Double, g() As Double, y() As Double
  Dim r() As Double, w() As Double, vin() As Double, z() As Double, preRhs() As Double, ax() As Double
  Dim i As Long, j As Long, k As Long, pass As Long, used As Long
  Dim normB As Double, beta As Double, inner As Double, nextNorm As Double, rot As Double, tmp As Double, residual As Double
  P6AccelGMRESCore = False: relres = 1E+30: usedIters = 0
  costStart = Timer: mAdaptiveGMRESBudgetStop = False
  On Error GoTo GMRESFail
  If n < 1 Or restart < 1 Or maxIters < 1 Or tol <= 0# Then Exit Function
  If restart > n Then restart = n
  ReDim x(n - 1): ReDim r(n - 1): ReDim w(n - 1): ReDim vin(n - 1): ReDim z(n - 1): ReDim ax(n - 1)
  ReDim v(n - 1, restart): ReDim zBasis(n - 1, restart - 1)
  ReDim h(restart, restart - 1): ReDim cs(restart - 1): ReDim sn(restart - 1): ReDim g(restart): ReDim y(restart - 1)
  For i = 0 To n - 1: normB = normB + rhs(i) * rhs(i): r(i) = rhs(i): Next i
  normB = Sqr(normB)
  If normB <= 1E-30 Then relres = 0#: P6AccelGMRESCore = True: Exit Function
  Do While usedIters < maxIters
    P6BandApply mat, x, ax, n, bw
    beta = 0#
    For i = 0 To n - 1: r(i) = rhs(i) - ax(i): beta = beta + r(i) * r(i): Next i
    beta = Sqr(beta): relres = beta / normB
    If relres <= tol Then P6AccelGMRESCore = True: Exit Function
    For i = 0 To n - 1: v(i, 0) = r(i) / beta: Next i
    For i = 0 To restart: g(i) = 0#: Next i
    g(0) = beta: used = 0
    For j = 0 To restart - 1
      If maxCostMs > 0# Then
        If P6ElapsedMs(costStart) >= maxCostMs Then mAdaptiveGMRESBudgetStop = True: Exit Function
      End If
      For i = 0 To n - 1: vin(i) = v(i, j): Next i
      preRhs = vin
      If Not P6BandSolve1D(packed, rowOff, preRhs, z, n, bw, pivotTol) Then Exit Function
      For i = 0 To n - 1
        If fixed(i) Then z(i) = 0#
        zBasis(i, j) = z(i)
      Next i
      P6BandApply mat, z, w, n, bw
      For k = 0 To restart: h(k, j) = 0#: Next k
      ' Modified Gram-Schmidt with a second pass near loss of orthogonality.
      For pass = 1 To 2
        For k = 0 To j
          inner = 0#
          For i = 0 To n - 1: inner = inner + w(i) * v(i, k): Next i
          h(k, j) = h(k, j) + inner
          For i = 0 To n - 1: w(i) = w(i) - inner * v(i, k): Next i
        Next k
      Next pass
      nextNorm = 0#
      For i = 0 To n - 1: nextNorm = nextNorm + w(i) * w(i): Next i
      nextNorm = Sqr(nextNorm): h(j + 1, j) = nextNorm
      If nextNorm > 1E-30 Then
        For i = 0 To n - 1: v(i, j + 1) = w(i) / nextNorm: Next i
      End If
      For k = 0 To j - 1
        tmp = cs(k) * h(k, j) + sn(k) * h(k + 1, j)
        h(k + 1, j) = -sn(k) * h(k, j) + cs(k) * h(k + 1, j): h(k, j) = tmp
      Next k
      rot = Sqr(h(j, j) ^ 2 + h(j + 1, j) ^ 2)
      If rot <= 1E-30 Then Exit For
      cs(j) = h(j, j) / rot: sn(j) = h(j + 1, j) / rot
      h(j, j) = rot: h(j + 1, j) = 0#
      tmp = cs(j) * g(j): g(j + 1) = -sn(j) * g(j): g(j) = tmp
      used = j + 1: usedIters = usedIters + 1
      If maxCostMs > 0# Then
        If P6ElapsedMs(costStart) >= maxCostMs Then mAdaptiveGMRESBudgetStop = True: Exit Function
      End If
      If Abs(g(j + 1)) <= tol * normB Or nextNorm <= 1E-30 Or usedIters >= maxIters Then Exit For
    Next j
    If used < 1 Then Exit Function
    For i = used - 1 To 0 Step -1
      tmp = g(i)
      For k = i + 1 To used - 1: tmp = tmp - h(i, k) * y(k): Next k
      If Abs(h(i, i)) <= 1E-30 Then Exit Function
      y(i) = tmp / h(i, i)
    Next i
    For i = 0 To n - 1
      For k = 0 To used - 1: x(i) = x(i) + zBasis(i, k) * y(k): Next k
      If Not P2IsFinite(x(i)) Then Exit Function
    Next i
  Loop
  P6BandApply mat, x, ax, n, bw
  residual = 0#
  For i = 0 To n - 1: residual = residual + (rhs(i) - ax(i)) ^ 2: Next i
  relres = Sqr(residual) / normB
  P6AccelGMRESCore = (relres <= tol)
  Exit Function
GMRESFail:
  If Err.Number = 18 Then Err.Raise 18
  Err.Clear
End Function

Private Function P6AccelTryGMRES() As Boolean
  Dim mat() As Double, rhs() As Double, x() As Double, fixed() As Boolean, boundary() As Double
  Dim i As Long, j As Long, col As Long, iters As Long, maxIters As Long
  Dim relres As Double, t0 As Double, tol As Double, budgetMs As Double, remainingMs As Double, gmresOk As Boolean
  P6AccelTryGMRES = False
  If Not AccelV5 Or AccelBaselineRetry Or Not mPreReady Then Exit Function
  If P6MixedUP Or P3HasJointElements Or P6UseCSR Then Exit Function
  If mPreN <> nDof Or mPreBw <> BandWidth Or mPreFs <> P3CurrentStrengthFactor Then Exit Function
  If mPreActive <> P3ActiveSetGen Or mPreConstraint <> P3ConstraintGeneration Then Exit Function
  If mPreMaterial <> P6MaterialGeneration Or mPreStrength <> P6StrengthGeneration Or mPreGeometry <> P6GeometryRevision Then Exit Function
  If mPreUses >= 4 Then mPreReady = False: Exit Function
  budgetMs = AdaptiveGMRESBudgetMs()
  If budgetMs < 0# Then Exit Function
  AdaptiveUseMethod 5
  mAdaptiveGMRESBudgetStop = False
  t0 = Timer
  On Error GoTo GMRESOptionalFail
  mat = OriginalMat: ReDim rhs(lastDof): ReDim fixed(lastDof): boundary = Disp
  For i = 0 To lastDof
    fixed(i) = (NodeCond(i) <> 0)
    If Not fixed(i) Then rhs(i) = P6ForceBeforeBoundary(i)
  Next i
  ' Exact Dirichlet elimination for the fresh operator. The old LU is only a preconditioner.
  For i = 0 To lastDof
    If fixed(i) Then mat(i, 0) = 1#
    For j = 1 To BandWidth
      col = i + j
      If col > lastDof Then Exit For
      If fixed(i) Or fixed(col) Then
        If Not fixed(i) Then rhs(i) = rhs(i) - mat(i, j) * boundary(col)
        If Not fixed(col) Then rhs(col) = rhs(col) - mat(i, j) * boundary(i)
        mat(i, j) = 0#
      End If
    Next j
  Next i
  maxIters = 60
  If P6IterativeMaxIterations < maxIters Then maxIters = P6IterativeMaxIterations
  tol = P6IterativeTolerance
  If tol > 0.000000000001 Then tol = 0.000000000001
  remainingMs = 0#
  If budgetMs > 0# Then remainingMs = budgetMs - P6ElapsedMs(t0)
  If budgetMs > 0# And remainingMs <= 0# Then
    mAdaptiveGMRESBudgetStop = True
  Else
    gmresOk = P6AccelGMRESCore(mat, mPrePacked, mPreOffsets, rhs, x, fixed, nDof, BandWidth, tol, mPreTol, P6GMRESRestart, maxIters, relres, iters, remainingMs)
  End If
  If gmresOk Then
    For i = 0 To lastDof
      If fixed(i) Then Disp(i) = boundary(i) Else Disp(i) = x(i)
    Next i
    mPreUses = mPreUses + 1
    AccelGMRESCount = AccelGMRESCount + 1
    P6IterativeLastResidual = relres: P6IterativeLastIterations = iters
    P6AccelTryGMRES = True
  Else
    AccelGMRESFallback = AccelGMRESFallback + 1
    mPreReady = False
  End If
  P6ProfSolveMs = P6ProfSolveMs + P6ElapsedMs(t0)
  AdaptiveNoteGMRES P6ElapsedMs(t0), P6AccelTryGMRES, mAdaptiveGMRESBudgetStop
  If AccelTrace Then P6SolverEvent "GMRES_LU", "ok=" & CStr(P6AccelTryGMRES) & ";iterations=" & CStr(iters) & ";true_relres=" & Format$(relres, "0.000E+00") & ";tol=" & Format$(tol, "0.000E+00")
  Exit Function
GMRESOptionalFail:
  If Err.Number = 18 Then Err.Raise 18
  Err.Clear
  mPreReady = False: P6AccelTryGMRES = False
  AccelGMRESFallback = AccelGMRESFallback + 1
  P6ProfSolveMs = P6ProfSolveMs + P6ElapsedMs(t0)
  AdaptiveNoteGMRES P6ElapsedMs(t0), False, mAdaptiveGMRESBudgetStop
End Function

Public Function P6AccelLinearSelfTest() As String
  Dim mat() As Double, old() As Double, rhs() As Double, x() As Double, fixed() As Boolean
  Dim packed() As Double, offsets() As Long, akj() As Double, aik() As Double
  Dim rel As Double, iters As Long, ok As Boolean, i As Long
  ReDim mat(2, 1): ReDim old(2, 1): ReDim rhs(2): ReDim fixed(2): ReDim akj(1): ReDim aik(1)
  mat(0, 0) = 4#: mat(1, 0) = 5#: mat(2, 0) = 6#: mat(0, 1) = 1#: mat(1, 1) = -1#
  old = mat: old(0, 0) = 4.5: old(1, 0) = 4.5: old(2, 0) = 5.5
  ReDim offsets(2)
  P6BandBuildRowOff offsets, 3, 2
  ReDim packed(5)
  P6PackBand old, packed, offsets, 3, 1
  ok = P6BandFactor1D(packed, offsets, akj, aik, 3, 1, 0.000000000001)
  rhs(0) = 6#: rhs(1) = 8#: rhs(2) = 16# ' exact x=(1,2,3)
  If ok Then ok = P6AccelGMRESCore(mat, packed, offsets, rhs, x, fixed, 3, 1, 0.0000000001, 0.000000000001, 3, 12, rel, iters)
  If ok Then
    For i = 0 To 2: If Abs(x(i) - (i + 1)) > 0.00000001 Then ok = False
    Next i
  End If
  If Not ok Then Err.Raise vbObjectError + 3601, "P6AccelLinearSelfTest", "Old-LU GMRES regression failed"
  ' Exhausted iteration budget must return False, never a false convergence.
  ok = P6AccelGMRESCore(mat, packed, offsets, rhs, x, fixed, 3, 1, 0.00000000000001, 0.000000000001, 1, 1, rel, iters)
  If ok Then Err.Raise vbObjectError + 3602, "P6AccelLinearSelfTest", "GMRES budget guard failed"
  P6AccelLinearSelfTest = "PASS: old-LU GMRES, true residual, iteration-budget rejection"
End Function




