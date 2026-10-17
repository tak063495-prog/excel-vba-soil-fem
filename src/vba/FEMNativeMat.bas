Option Explicit
' Optional x64 native material kernels (soilfem_mat_x64.dll next to the workbook).
' Coarse-grained: one call evaluates every element it can (material-point
' update, stresses, algorithmic tangent, internal force) directly on the Elem()
' array; one call builds every element tangent. The DLL stops at the first
' element it does not handle (joint, substepping, vertex return, finite-
' difference tangent, arithmetic error); VBA evaluates that element with the
' unchanged code and calls the DLL again. Logging, failures and fallback stay
' in VBA. NATIVE_DLL: AUTO = band DLL + material DLL, BAND = band DLL only, OFF.
#If Win64 Then
Private Declare PtrSafe Function LoadLibraryExW Lib "kernel32" (ByVal path As LongPtr, ByVal file As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal module As LongPtr, ByVal name As String) As LongPtr
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal name As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleFileNameW Lib "kernel32" (ByVal module As LongPtr, ByVal buffer As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Function SfmInit Lib "soilfem_mat_x64.dll" Alias "sfm_init" () As Long
Private Declare PtrSafe Function SfmEval Lib "soilfem_mat_x64.dll" Alias "sfm_eval" (ByVal params As LongPtr, ByVal kStart As Long) As Long
Private Declare PtrSafe Function SfmTangent Lib "soilfem_mat_x64.dll" Alias "sfm_tangent" (ByVal params As LongPtr, ByVal doneFlags As LongPtr) As Long
Private mHandle As LongPtr
Private mOff(0 To 25) As LongLong
Private mPar(0 To 15) As LongLong
Private mOut(0 To 15) As LongLong
#End If
Private Const SFM_DLL As String = "soilfem_mat_x64.dll"
Private Const SFM_API As Long = 4
Private Const SFM_MAT_STRIDE As Long = 16
Private mChecked As Boolean, mAvailable As Boolean, mReason As String
Private mMat() As Double
Private mDone() As Integer
Private mTangentOK As Boolean, mDoneReady As Boolean
Public FEMMatEvalCalls As Long, FEMMatEvalElements As Long, FEMMatDeclines As Long
Public FEMMatTangentCalls As Long, FEMMatTangentElements As Long, FEMMatFallbacks As Long
Public FEMMatDeclineReason(0 To 7) As Long

Public Sub FEMMatResetCounts()
  Dim i As Long
  FEMMatEvalCalls = 0: FEMMatEvalElements = 0: FEMMatDeclines = 0
  FEMMatTangentCalls = 0: FEMMatTangentElements = 0: FEMMatFallbacks = 0
  FEMKernResetCounts
  For i = 0 To 7: FEMMatDeclineReason(i) = 0: Next i
End Sub

Public Sub FEMMatRefresh()
  mChecked = False: mAvailable = False: mTangentOK = False
  FEMMatResetCounts
End Sub

Public Function FEMMatStatusText() As String
  If FEMMatAvailable() Then
    FEMMatStatusText = "NATIVE(" & SFM_DLL & ";eval_calls=" & CStr(FEMMatEvalCalls) & ";eval_elements=" & CStr(FEMMatEvalElements) & _
      ";declined=" & CStr(FEMMatDeclines) & "[joint=" & CStr(FEMMatDeclineReason(2)) & ",material=" & CStr(FEMMatDeclineReason(3)) & _
      ",substep=" & CStr(FEMMatDeclineReason(4)) & ",point=" & CStr(FEMMatDeclineReason(5)) & ",fp=" & CStr(FEMMatDeclineReason(6)) & _
      ",dof=" & CStr(FEMMatDeclineReason(7)) & "];tangent_calls=" & CStr(FEMMatTangentCalls) & ";tangent_elements=" & CStr(FEMMatTangentElements) & _
      ";fallback=" & CStr(FEMMatFallbacks) & ")"
  Else
    FEMMatStatusText = "VBA(" & mReason & ")"
  End If
End Function

Public Function FEMMatAvailable() As Boolean
  Dim path As String, loadedPath As String, size As Long, dllName As String
  Dim exportName As Variant
  If mChecked Then FEMMatAvailable = mAvailable: Exit Function
  mChecked = True: mAvailable = False
  On Error GoTo LoadFailed
  If FEMNativeModeText() <> "AUTO" And FEMNativeModeText() <> "MAT" Then mReason = FEMNativeModeText() & "_SELECTED": Exit Function
#If Win64 Then
  path = ThisWorkbook.path & Application.PathSeparator & SFM_DLL
  If Len(Dir$(path)) = 0 Then mReason = "DLL_NOT_FOUND": Exit Function
  dllName = SFM_DLL
  If mHandle = 0 Then mHandle = GetModuleHandleW(StrPtr(dllName))
  If mHandle <> 0 Then
    loadedPath = String$(32768, vbNullChar)
    size = GetModuleFileNameW(mHandle, StrPtr(loadedPath), Len(loadedPath))
    If size = 0 Or size >= Len(loadedPath) Then mReason = "DLL_PATH_UNVERIFIED": Exit Function
    If StrComp(Left$(loadedPath, size), path, vbTextCompare) <> 0 Then mReason = "DLL_ALREADY_LOADED_OTHER_PATH": Exit Function
  Else
    mHandle = LoadLibraryExW(StrPtr(path), 0, &H1100)
  End If
  If mHandle = 0 Then mReason = "DLL_LOAD_FAILED_" & CStr(GetLastError()): Exit Function
  For Each exportName In Array("sfm_init", "sfm_eval", "sfm_tangent")
    If GetProcAddress(mHandle, CStr(exportName)) = 0 Then mReason = "DLL_ENTRY_MISSING_" & CStr(exportName): Exit Function
  Next exportName
  size = SfmInit()
  If size = -1 Then mReason = "VBA_RUNTIME_MSVCR100_NOT_LOADED": Exit Function
  If size <> SFM_API Then mReason = "DLL_API_MISMATCH_" & CStr(size): Exit Function
  If Not FEMMatBuildOffsets() Then Exit Function
  mReason = "OK"
  mAvailable = True
  FEMMatAvailable = True
#Else
  mReason = "NOT_WIN64"
#End If
  Exit Function
LoadFailed:
  mReason = "DLL_ERROR_" & CStr(Err.Number)
  mAvailable = False
  FEMMatAvailable = False
  Err.Clear
End Function

Private Sub FEMMatDisable(ByVal why As String)
  mAvailable = False: mChecked = True: mReason = why
  FEMMatFallbacks = FEMMatFallbacks + 1
End Sub

#If Win64 Then
' Member offsets of Element_Data measured with VarPtr (no layout assumption).
Private Function FEMMatBuildOffsets() As Boolean
  Dim b As LongPtr, i As Long
  FEMMatBuildOffsets = False
  ' Not cached: the mesh may not be loaded yet on the first query.
  If NumberOfElement < 1 Then mReason = "NO_ELEMENTS": mChecked = False: Exit Function
  b = VarPtr(Elem(0))
  If NumberOfElement >= 2 Then mOff(0) = VarPtr(Elem(1)) - b Else mOff(0) = 0
  mOff(1) = VarPtr(Elem(0).u(0)) - b
  mOff(2) = VarPtr(Elem(0).MatNo) - b
  mOff(3) = VarPtr(Elem(0).ElNode(0)) - b
  mOff(4) = VarPtr(Elem(0).Bmat(0, 0, 0)) - b
  mOff(5) = VarPtr(Elem(0).Smat(0, 0, 0)) - b
  mOff(6) = VarPtr(Elem(0).Spmat(0, 0, 0)) - b
  mOff(7) = VarPtr(Elem(0).Stmat(0, 0)) - b
  mOff(8) = VarPtr(Elem(0).mStmat(0, 0)) - b
  mOff(9) = VarPtr(Elem(0).P2PlasticStrain(0, 0)) - b
  mOff(10) = VarPtr(Elem(0).P2PlasticMultiplier(0)) - b
  mOff(11) = VarPtr(Elem(0).P2YieldFunction(0)) - b
  mOff(12) = VarPtr(Elem(0).P2Yielded(0)) - b
  mOff(13) = VarPtr(Elem(0).P2PrincipalStress(0, 0)) - b
  mOff(14) = VarPtr(Elem(0).P2LastSuccessfulSubsteps(0)) - b
  mOff(15) = VarPtr(Elem(0).P2EasySubstepStreak(0)) - b
  mOff(16) = VarPtr(Elem(0).Dpmat(0, 0, 0)) - b
  mOff(17) = VarPtr(Elem(0).kmat(0, 0)) - b
  mOff(18) = VarPtr(Elem(0).dj(0)) - b
  mOff(19) = VarPtr(Elem(0).ElmWeight) - b
  mOff(20) = VarPtr(Elem(0).HourglassModeCount) - b
  mOff(21) = VarPtr(Elem(0).HourglassScale) - b
  mOff(22) = VarPtr(Elem(0).HourglassMode(0, 0)) - b
  mOff(23) = VarPtr(Elem(0).IsJoint) - b
  mOff(24) = VarPtr(Elem(0).SpmatValid(0)) - b
  mOff(25) = VarPtr(Elem(0).TangentDirty) - b
  ' Layout checks: column-major inline arrays, 2-byte Boolean, 4-byte Long.
  If VarPtr(Elem(0).Bmat(1, 0, 0)) - VarPtr(Elem(0).Bmat(0, 0, 0)) <> 8 Then mReason = "LAYOUT_BMAT": Exit Function
  If VarPtr(Elem(0).Bmat(0, 1, 0)) - VarPtr(Elem(0).Bmat(0, 0, 0)) <> 24 Then mReason = "LAYOUT_BMAT_COL": Exit Function
  If VarPtr(Elem(0).Stmat(0, 1)) - VarPtr(Elem(0).Stmat(0, 0)) <> 136 Then mReason = "LAYOUT_STMAT": Exit Function
  If VarPtr(Elem(0).P2Yielded(1)) - VarPtr(Elem(0).P2Yielded(0)) <> 2 Then mReason = "LAYOUT_BOOLEAN": Exit Function
  If VarPtr(Elem(0).ElNode(1)) - VarPtr(Elem(0).ElNode(0)) <> 4 Then mReason = "LAYOUT_LONG": Exit Function
  If VarPtr(Elem(0).HourglassMode(0, 1)) - VarPtr(Elem(0).HourglassMode(0, 0)) <> 32 Then mReason = "LAYOUT_HGMODE": Exit Function
  If VarPtr(Elem(0).kmat(0, 1)) - VarPtr(Elem(0).kmat(0, 0)) <> 128 Then mReason = "LAYOUT_KMAT": Exit Function
  For i = 1 To 25
    If mOff(i) < 0 Or (NumberOfElement >= 2 And mOff(i) >= mOff(0)) Then mReason = "LAYOUT_OFFSET_" & CStr(i): Exit Function
  Next i
  FEMMatBuildOffsets = True
End Function

Private Sub FEMMatBuildMaterialTable()
  Dim i As Long, n As Long, o As Long
  n = UBound(Material) + 1
  ReDim mMat(0 To n * SFM_MAT_STRIDE - 1)
  For i = 0 To n - 1
    o = i * SFM_MAT_STRIDE
    If i < NumberOfMaterial Then
      If Material(i).MaterialCacheReady Then mMat(o) = 1#
    End If
    mMat(o + 1) = Material(i).ElasticD00
    mMat(o + 2) = Material(i).ElasticD01
    mMat(o + 3) = Material(i).ElasticD22
    mMat(o + 4) = Material(i).SinFriction
    mMat(o + 5) = Material(i).CosFriction
    mMat(o + 6) = Material(i).SinDilation
    mMat(o + 7) = Material(i).young
    mMat(o + 8) = Material(i).Poisson
    mMat(o + 9) = Material(i).cohesion
    mMat(o + 10) = Material(i).thickness
    mMat(o + 11) = Material(i).weight
  Next i
  mPar(9) = VarPtr(mMat(0))
  mPar(10) = n
End Sub

Private Sub FEMMatCommonParams(ByVal flags As Long)
  Dim i As Long
  mPar(0) = VarPtr(Elem(0))
  mPar(1) = VarPtr(mOff(0))
  mPar(2) = NumberOfElement
  mPar(6) = 0: mPar(7) = 0: mPar(8) = -1
  If P3ElementActiveReady Then
    mPar(6) = VarPtr(P3ElementActive(LBound(P3ElementActive)))
    mPar(7) = LBound(P3ElementActive): mPar(8) = UBound(P3ElementActive)
  End If
  FEMMatBuildMaterialTable
  mPar(12) = flags
  mPar(13) = VarPtr(mOut(0))
  mPar(14) = VarPtr(P6GaussN(0, 0))
  For i = 0 To 15: mOut(i) = 0: Next i
  mOut(8) = -1
End Sub
#End If

' Evaluates elements kStart.. natively. Returns the first element index that
' VBA must evaluate (NumberOfElement when all done). Never raises.
Public Function FEMMatEval(ByVal kStart As Long, ByRef incrementBase() As Double, ByVal flags As Long) As Long
  Dim r As Long, reason As Long
  FEMMatEval = kStart
  If Not FEMMatAvailable() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  If LBound(TDisp) <> 0 Or LBound(incrementBase) <> 0 Or LBound(iNForce) <> 0 Then Exit Function
  If (flags And 2) <> 0 Then
    If LBound(P3ActivePlasticPoint, 1) <> 0 Or UBound(P3ActivePlasticPoint, 1) <> 3 Or LBound(P3ActivePlasticPoint, 2) <> 0 Or UBound(P3ActivePlasticPoint, 2) < NumberOfElement - 1 Then Exit Function
  End If
  FEMMatCommonParams flags
  mPar(3) = VarPtr(TDisp(0))
  mPar(4) = VarPtr(incrementBase(0))
  mPar(5) = VarPtr(iNForce(0))
  mPar(11) = 0
  If (flags And 2) <> 0 Then mPar(11) = VarPtr(P3ActivePlasticPoint(0, 0))
  mPar(15) = UBound(TDisp)
  If UBound(incrementBase) < mPar(15) Then mPar(15) = UBound(incrementBase)
  If UBound(iNForce) < mPar(15) Then mPar(15) = UBound(iNForce)
  r = SfmEval(VarPtr(mPar(0)), kStart)
  If r < kStart Or r > NumberOfElement Then FEMMatDisable "EVAL_RC_" & CStr(r): Exit Function
  FEMMatEvalCalls = FEMMatEvalCalls + 1
  If mOut(7) > 0 Then
    FEMMatEvalElements = FEMMatEvalElements + CLng(mOut(7))
    AccelMaterialCalls = AccelMaterialCalls + CLng(mOut(0))
    AccelSubstepCalls(0) = AccelSubstepCalls(0) + CLng(mOut(1))
    AccelMaterialSuccess = AccelMaterialSuccess + CLng(mOut(2))
    P3ElasticFastCount = P3ElasticFastCount + CLng(mOut(3))
    P3PlasticStateChangeCount = P3PlasticStateChangeCount + CLng(mOut(4))
    P3PlasticPointUpdateCount = P3PlasticPointUpdateCount + CLng(mOut(5))
    P3ActivePlasticPointCount = P3ActivePlasticPointCount + CLng(mOut(6))
  End If
  If mOut(8) >= 0 Then FailureElement = CLng(mOut(8))
  If r < NumberOfElement Then
    FEMMatDeclines = FEMMatDeclines + 1
    reason = CLng(mOut(9))
    If reason >= 0 And reason <= 7 Then FEMMatDeclineReason(reason) = FEMMatDeclineReason(reason) + 1
  End If
  FEMMatEval = r
  Exit Function
Failed:
  FEMMatDisable "EVAL_ERR_" & CStr(Err.Number)
  Err.Clear
#End If
End Function

' Builds kmat/ElmWeight for every eligible element. False = VBA builds all.
Public Function FEMMatTangent(ByVal elasticCorrection As Boolean, ByVal inconsistent As Boolean) As Boolean
  Dim rc As Long, flags As Long
  mTangentOK = False
  FEMMatTangent = False
  If Not FEMMatAvailable() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  If elasticCorrection Then flags = flags Or 8
  If inconsistent Then flags = flags Or 16
  FEMMatCommonParams flags
  mPar(3) = 0: mPar(4) = 0: mPar(5) = 0: mPar(11) = 0: mPar(15) = -1
  If mDoneReady Then
    If UBound(mDone) <> NumberOfElement - 1 Then mDoneReady = False
  End If
  If Not mDoneReady Then ReDim mDone(0 To NumberOfElement - 1): mDoneReady = True
  rc = SfmTangent(VarPtr(mPar(0)), VarPtr(mDone(0)))
  FEMMatTangentCalls = FEMMatTangentCalls + 1
  If rc <> 0 Then FEMMatFallbacks = FEMMatFallbacks + 1: Exit Function
  FEMMatTangentElements = FEMMatTangentElements + CLng(mOut(10))
  mTangentOK = True
  FEMMatTangent = True
  Exit Function
Failed:
  FEMMatDisable "TANGENT_ERR_" & CStr(Err.Number)
  Err.Clear
#End If
End Function

#If Win64 Then
Public Function FEMMatOffset(ByVal i As Long) As LongLong
  FEMMatOffset = mOff(i)
End Function
Public Function FEMMatDonePtr() As LongPtr
  FEMMatDonePtr = 0
  If mTangentOK And mDoneReady Then FEMMatDonePtr = VarPtr(mDone(0))
End Function
Public Function FEMMatParamsPtr() As LongPtr
  FEMMatCommonParams 0
  FEMMatParamsPtr = VarPtr(mPar(0))
End Function
#End If

Public Function FEMMatTangentDone(ByVal k As Long) As Boolean
  If Not mTangentOK Then Exit Function
  FEMMatTangentDone = (mDone(k) <> 0)
End Function
