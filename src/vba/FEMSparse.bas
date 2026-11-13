Option Explicit
' Optional Eigen sparse LDLT for an assembled symmetric band operator.
' Force/material evaluation, nonlinear convergence and fallback remain in VBA.
#If Win64 Then
Private Declare PtrSafe Function LoadLibraryExW Lib "kernel32" (ByVal path As LongPtr, ByVal file As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal name As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleFileNameW Lib "kernel32" (ByVal module As LongPtr, ByVal buffer As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal module As LongPtr, ByVal name As String) As LongPtr
Private Declare PtrSafe Function SparseAbi Lib "soilfem_sparse_x64.dll" Alias "sfs_abi" () As Long
Private Declare PtrSafe Function SparseFactor Lib "soilfem_sparse_x64.dll" Alias "sfs_factor_band" (ByVal n As Long, ByVal bw As Long, ByVal ld As Long, ByVal band As LongPtr, ByVal locked As LongPtr, ByVal limitBytes As Double, ByRef token As LongLong, ByVal metrics As LongPtr) As Long
Private Declare PtrSafe Function SparseSolve Lib "soilfem_sparse_x64.dll" Alias "sfs_solve" (ByVal token As LongLong, ByVal n As Long, ByVal rhs As LongPtr, ByVal prescribed As LongPtr, ByVal tolerance As Double, ByVal result As LongPtr, ByVal metrics As LongPtr) As Long
Private Declare PtrSafe Function SparseRelease Lib "soilfem_sparse_x64.dll" Alias "sfs_release" (ByVal token As LongLong) As Long
Private mHandle As LongPtr, mToken As LongLong
#End If
Private mChecked As Boolean, mAvailable As Boolean, mUsed As Boolean, mStarted As Boolean
Private mFactorGeneration As Long, mConstraintGeneration As Long, mRejectedGeneration As Long
Private mMetrics(0 To 7) As Double
Public FEMSparseFactorCalls As Long, FEMSparseSolveCalls As Long
Public FEMSparseReason As String

Public Sub FEMSparseReset()
  FEMSparseReleaseFactor
  mChecked = False: mAvailable = False: mStarted = True: mRejectedGeneration = -1
  FEMSparseFactorCalls = 0: FEMSparseSolveCalls = 0: FEMSparseReason = "NOT_SELECTED"
End Sub

Public Sub FEMSparseReleaseFactor()
#If Win64 Then
  On Error Resume Next
  If mToken <> 0 Then SparseRelease mToken
  mToken = 0
  On Error GoTo 0
#End If
  mUsed = False
End Sub

Public Function FEMSparseHasFactor() As Boolean
  FEMSparseHasFactor = mUsed
End Function

Private Function SparseAvailable() As Boolean
#If Win64 Then
  Dim path As String, dllName As String, loaded As String, size As Long, entry As Variant
  If FEMNativeModeText() <> "AUTO" And FEMNativeModeText() <> "BAND" Then FEMSparseReason = "NATIVE_OFF": Exit Function
  If mChecked Then SparseAvailable = mAvailable: Exit Function
  mChecked = True
  On Error GoTo Failed
  dllName = "soilfem_sparse_x64.dll": path = ThisWorkbook.path & Application.PathSeparator & dllName
  If Len(Dir$(path)) = 0 Then FEMSparseReason = "DLL_NOT_FOUND": Exit Function
  If mHandle = 0 Then mHandle = GetModuleHandleW(StrPtr(dllName))
  If mHandle <> 0 Then
    loaded = String$(32768, vbNullChar): size = GetModuleFileNameW(mHandle, StrPtr(loaded), Len(loaded))
    If size = 0 Or size >= Len(loaded) Then FEMSparseReason = "DLL_PATH_UNVERIFIED": Exit Function
    If StrComp(Left$(loaded, size), path, vbTextCompare) <> 0 Then FEMSparseReason = "DLL_LOADED_OTHER_PATH": Exit Function
  Else
    mHandle = LoadLibraryExW(StrPtr(path), 0, &H1100)
  End If
  If mHandle = 0 Then FEMSparseReason = "DLL_LOAD_FAILED": Exit Function
  For Each entry In Array("sfs_abi", "sfs_factor_band", "sfs_solve", "sfs_release")
    If GetProcAddress(mHandle, CStr(entry)) = 0 Then FEMSparseReason = "DLL_ENTRY_MISSING": Exit Function
  Next entry
  If SparseAbi() <> &H10000 Then FEMSparseReason = "DLL_ABI_MISMATCH": Exit Function
  mAvailable = True: FEMSparseReason = "READY": SparseAvailable = True
  Exit Function
Failed:
  If Err.number = 18 Then Err.Raise 18
  FEMSparseReason = "DLL_ERROR_" & CStr(Err.number): Err.Clear
#Else
  FEMSparseReason = "WIN64_REQUIRED"
#End If
End Function

' 0=not selected/available, 1=solved, 2=discard this factor and use band fallback.
Public Function FEMSparseTryBandSolve(ByRef newFactor As Boolean) As Long
#If Win64 Then
  Dim eligible As Boolean, rc As Long, tolerance As Double
  On Error GoTo Failed
  If Not mStarted Then FEMSparseReset
  newFactor = False
  eligible = (P6ReadSetting("NATIVE_SPARSE_DIRECT", 0#) = 1#)
  If P6UseCSR Or P6NeedsFullNonsymmetric() Or P3HasJointElements Or FSActive Or P6MixedUP Or P6ConsolActive Then eligible = False
  If Not eligible Then
    If mUsed Then FEMSparseReleaseFactor: FEMSparseTryBandSolve = 2
    Exit Function
  End If
  If Not SparseAvailable() Then
    If mUsed Then FEMSparseReleaseFactor: FEMSparseTryBandSolve = 2
    Exit Function
  End If
  If mRejectedGeneration = P6TangentGeneration Then Exit Function
  If nDof < 1 Or lastDof <> nDof - 1 Or BandWidth < 0 Then GoTo Failed
  If LBound(OriginalMat, 1) <> 0 Or UBound(OriginalMat, 1) < lastDof Or LBound(OriginalMat, 2) <> 0 Or UBound(OriginalMat, 2) < BandWidth Then GoTo Failed
  If Not P6ForceBeforeBoundaryReady Or LBound(P6ForceBeforeBoundary) <> 0 Or UBound(P6ForceBeforeBoundary) < lastDof Then GoTo Failed
  If LBound(NodeCond) <> 0 Or UBound(NodeCond) < lastDof Or LBound(Disp) <> 0 Or UBound(Disp) < lastDof Then GoTo Failed
  If Not P6FactorReady Or mToken = 0 Or mFactorGeneration <> P6TangentGeneration Or mConstraintGeneration <> P3ConstraintGeneration Then
    FEMSparseReleaseFactor
    rc = SparseFactor(nDof, BandWidth, UBound(OriginalMat, 1) + 1, VarPtr(OriginalMat(0, 0)), VarPtr(NodeCond(0)), P6ReadSetting("SOLVER_MEMORY_LIMIT_MB", 512#) * 1048576#, mToken, VarPtr(mMetrics(0)))
    If rc <> 0 Then FEMSparseReason = "FACTOR_RC_" & CStr(rc): GoTo Failed
    newFactor = True: mFactorGeneration = P6TangentGeneration: mConstraintGeneration = P3ConstraintGeneration
    FEMSparseFactorCalls = FEMSparseFactorCalls + 1
    P6ProfFactorMs = P6ProfFactorMs + 1000# * mMetrics(4)
    P6SolverEvent "SPARSE_DIRECT_FACTOR", "scaled_min_D=" & CStr(mMetrics(0)) & ";nnz_L=" & CStr(mMetrics(2)) & ";seconds=" & CStr(mMetrics(4)) & ";estimated_bytes=" & CStr(mMetrics(5))
  End If
  tolerance = P6ReadSetting("SOLVER_TOLERANCE", 0.00000001)
  If tolerance > 0.0000000001 Then tolerance = 0.0000000001
  If tolerance < 0.00000000000001 Then tolerance = 0.00000000000001
  rc = SparseSolve(mToken, nDof, VarPtr(P6ForceBeforeBoundary(0)), VarPtr(Disp(0)), tolerance, VarPtr(Disp(0)), VarPtr(mMetrics(0)))
  If rc <> 0 Then FEMSparseReason = "SOLVE_RC_" & CStr(rc): GoTo Failed
  FEMSparseSolveCalls = FEMSparseSolveCalls + 1: mUsed = True: FEMSparseReason = "OK"
  P6ProfSolveMs = P6ProfSolveMs + 1000# * mMetrics(7)
  P6IterativeLastResidual = mMetrics(6): P6IterativeLastIterations = 0
  P6SolverMode = "SPARSE_LDLT": FEMSparseTryBandSolve = 1
  Exit Function
Failed:
  If Err.number = 18 Then Err.Raise 18
  If Err.number <> 0 Then FEMSparseReason = "VBA_ERROR_" & CStr(Err.number)
  Err.Clear
  P6SolverEvent "SPARSE_DIRECT_FALLBACK", FEMSparseReason
  FEMSparseReleaseFactor: mRejectedGeneration = P6TangentGeneration
  FEMSparseTryBandSolve = 2
#End If
End Function
