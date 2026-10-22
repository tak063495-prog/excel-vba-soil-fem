Option Explicit
' Optional x64 native kernels (soilfem_native_x64.dll next to the workbook).
' Coarse-grained: one call factors the whole band / solves the whole RHS.
' VBA owns every array. Any load/ABI/argument problem falls back to the
' unchanged VBA kernels. Setting NATIVE_DLL: AUTO (default, band + material DLL)
' / BAND (band DLL only, 20261016 behaviour) / OFF.
#If Win64 Then
Private Declare PtrSafe Function LoadLibraryExW Lib "kernel32" (ByVal path As LongPtr, ByVal file As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal module As LongPtr, ByVal name As String) As LongPtr
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal name As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleFileNameW Lib "kernel32" (ByVal module As LongPtr, ByVal buffer As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Function SfnAbi Lib "soilfem_native_x64.dll" Alias "sfn_abi" () As Long
Private Declare PtrSafe Function SfnSetPrecision Lib "soilfem_native_x64.dll" Alias "sfn_set_precision" (ByVal mode As Long) As Long
Private Declare PtrSafe Function SfnGetPrecision Lib "soilfem_native_x64.dll" Alias "sfn_get_precision" () As Long
Private Declare PtrSafe Function SfnCSRScatterList Lib "soilfem_native_x64.dll" Alias "sfn_csr_scatter_list" (ByVal params As LongPtr) As Long
Private Declare PtrSafe Function SfnGmresTuned Lib "soilfem_native_x64.dll" Alias "sfn_gmres_tuned" (ByVal params As LongPtr, ByVal tol As Double) As Long
Private Declare PtrSafe Function SfnGmresPrepared Lib "soilfem_native_x64.dll" Alias "sfn_gmres_prepared" (ByVal params As LongPtr, ByVal tol As Double) As Long
Private Declare PtrSafe Function SfnGmres Lib "soilfem_native_x64.dll" Alias "sfn_gmres" (ByVal params As LongPtr, ByVal tol As Double) As Long
Private Declare PtrSafe Function SfnRestoreCommitted Lib "soilfem_native_x64.dll" Alias "sfn_restore_committed" (ByVal n As Long, ByVal stride As LongLong, ByVal dst As LongPtr, ByVal src As LongPtr, ByVal active As LongPtr, ByVal activeLo As Long, ByVal activeHi As Long, ByVal plasticActive As LongPtr, ByRef activeCount As Long) As Long
Private Declare PtrSafe Function SfnBandFactor Lib "soilfem_native_x64.dll" Alias "sfn_band_factor" (ByVal n As Long, ByVal bw As Long, ByVal ld As Long, ByVal band As LongPtr, ByVal packed As LongPtr, ByVal akj As LongPtr, ByVal tol As Double, ByRef failK As Long) As Long
Private Declare PtrSafe Function SfnBandSolve Lib "soilfem_native_x64.dll" Alias "sfn_band_solve" (ByVal n As Long, ByVal bw As Long, ByVal packed As LongPtr, ByVal rhs As LongPtr, ByVal x As LongPtr, ByVal tol As Double, ByRef failK As Long) As Long
Private Declare PtrSafe Function SfnNsFactor Lib "soilfem_native_x64.dll" Alias "sfn_nsband_factor" (ByVal n As Long, ByVal b As Long, ByVal ld As Long, ByVal a As LongPtr, ByVal rhs As LongPtr, ByVal pivotRows As LongPtr, ByVal lower As LongPtr, ByRef failK As Long, ByRef reason As Long, ByRef largest As Double) As Long
Private Declare PtrSafe Function SfnNsForward Lib "soilfem_native_x64.dll" Alias "sfn_nsband_forward" (ByVal n As Long, ByVal b As Long, ByVal ld As Long, ByVal pivotRows As LongPtr, ByVal lower As LongPtr, ByVal rhs As LongPtr) As Long
Private Declare PtrSafe Function SfnNsBack Lib "soilfem_native_x64.dll" Alias "sfn_nsband_back" (ByVal n As Long, ByVal b As Long, ByVal ld As Long, ByVal a As LongPtr, ByVal rhs As LongPtr, ByVal sol As LongPtr, ByVal checkPivot As Long, ByRef failK As Long) As Long
Private mHandle As LongPtr
#End If
Private Const SFN_DLL As String = "soilfem_native_x64.dll"
Private Const SFN_ABI As Long = &H10001
Private mChecked As Boolean, mAvailable As Boolean, mForcedOff As Boolean, mForcedSet As Boolean
Private mForcedText As String
Private mReason As String
Private mPrecisionText As String
Private mPrecisionAPI As Boolean, mPrecisionMode As Long
Public FEMNativeFactorCalls As Long, FEMNativeSolveCalls As Long, FEMNativeNsCalls As Long, FEMNativeFallbacks As Long
Public FEMNativeStateCopies As Long
Public FEMNativeGmresCalls As Long
Public FEMNativeGmresTunedCalls As Long
Public FEMNativeGmresPreparedCalls As Long
Public FEMNativeAssemblyCalls As Long

' Test/diagnostic switch: "AUTO", "BAND", "OFF" or "" (= follow NATIVE_DLL setting).
Public Sub FEMNativeSetMode(ByVal modeText As String)
  modeText = UCase$(Trim$(modeText))
  mForcedSet = (modeText = "AUTO" Or modeText = "OFF" Or modeText = "BAND" Or modeText = "MAT")
  mForcedOff = (modeText = "OFF")
  mForcedText = modeText
  mChecked = False: mAvailable = False
  FEMNativeResetCounts
  FEMMatRefresh
End Sub

Public Sub FEMNativeResetCounts()
  FEMNativeFactorCalls = 0: FEMNativeSolveCalls = 0: FEMNativeNsCalls = 0: FEMNativeFallbacks = 0
  FEMNativeStateCopies = 0
  FEMNativeGmresCalls = 0
  FEMNativeGmresTunedCalls = 0
  FEMNativeGmresPreparedCalls = 0
  FEMNativeAssemblyCalls = 0
End Sub

#If Win64 Then
Public Function FEMNativeGmres(ByVal params As LongPtr, ByVal tol As Double) As Long
  FEMNativeGmres = -1
  On Error GoTo FailedGmres
  If P6ReadSetting("NATIVE_GMRES_SIMD", 0#) <> 1# Then Exit Function
  If FEMNativeModeText() <> "AUTO" Then Exit Function
  If Not FEMNativeAvailable() Then Exit Function
  If mPrecisionMode = 0 Or Not mPrecisionAPI Then Exit Function
  If GetProcAddress(mHandle, "sfn_gmres") = 0 Then Exit Function
  If P6ReadSetting("NATIVE_GMRES_TUNED", 0#) = 1# And GetProcAddress(mHandle, "sfn_gmres_tuned") <> 0 Then
    If P6ReadSetting("NATIVE_GMRES_PREPARED", 0#) = 1# And GetProcAddress(mHandle, "sfn_gmres_prepared") <> 0 Then
      FEMNativeGmres = SfnGmresPrepared(params, tol)
      If FEMNativeGmres >= 0 Then FEMNativeGmresPreparedCalls = FEMNativeGmresPreparedCalls + 1
    Else
      FEMNativeGmres = SfnGmresTuned(params, tol)
    End If
    If FEMNativeGmres >= 0 Then FEMNativeGmresTunedCalls = FEMNativeGmresTunedCalls + 1
  Else
    FEMNativeGmres = SfnGmres(params, tol)
  End If
  If FEMNativeGmres >= 0 Then FEMNativeGmresCalls = FEMNativeGmresCalls + 1
  Exit Function
FailedGmres:
  If Err.number = 18 Then Err.Raise 18
  Err.Clear
End Function
#End If

' Optional acceleration: packaged builds enable it explicitly; absent key stays OFF.
' Exact byte copies of existing committed arrays; no floating-point operations.
Public Function FEMNativeRestoreState() As Boolean
#If Win64 Then
  Dim dst(0 To 8) As LongLong, src(0 To 8) As LongLong, stride As LongLong
  Dim activePtr As LongPtr, activeLo As Long, activeHi As Long, n As Long
  Dim i As Long, rc As Long
  On Error GoTo FailedState
  If P6ReadSetting("NATIVE_STATE_COPY", 0#) <> 1# Then Exit Function
  If FEMNativeModeText() <> "AUTO" Then Exit Function
  If Not FEMNativeAvailable() Then Exit Function
  If GetProcAddress(mHandle, "sfn_restore_committed") = 0 Then Exit Function
  n = NumberOfElement
  If n < 1 Or LBound(Elem) <> 0 Or UBound(Elem) < n - 1 Then Exit Function
  If n > 1 Then stride = VarPtr(Elem(1)) - VarPtr(Elem(0))
  If VarPtr(Elem(0).Stmat(0, 1)) - VarPtr(Elem(0).Stmat(0, 0)) <> 136 Then Exit Function
  If VarPtr(Elem(0).Spmat(0, 0, 1)) - VarPtr(Elem(0).Spmat(0, 0, 0)) <> 384 Then Exit Function
  If VarPtr(Elem(0).Dpmat(0, 0, 1)) - VarPtr(Elem(0).Dpmat(0, 0, 0)) <> 72 Then Exit Function
  If VarPtr(Elem(0).P2Yielded(1)) - VarPtr(Elem(0).P2Yielded(0)) <> 2 Then Exit Function
  ' Check all source dimensions before exposing their storage to the DLL.
  If UBound(P3CommittedPlasticStrain, 1) <> 3 Or UBound(P3CommittedPlasticStrain, 2) <> 3 Or UBound(P3CommittedPlasticStrain, 3) < n - 1 Then Exit Function
  If UBound(P3CommittedPlasticMultiplier, 1) <> 3 Or UBound(P3CommittedPlasticMultiplier, 2) < n - 1 Then Exit Function
  If UBound(P3CommittedYieldFunction, 1) <> 3 Or UBound(P3CommittedYieldFunction, 2) < n - 1 Then Exit Function
  If UBound(P3CommittedYielded, 1) <> 3 Or UBound(P3CommittedYielded, 2) < n - 1 Then Exit Function
  If UBound(P3CommittedPrincipalStress, 1) <> 2 Or UBound(P3CommittedPrincipalStress, 2) <> 3 Or UBound(P3CommittedPrincipalStress, 3) < n - 1 Then Exit Function
  If UBound(P3CommittedSpmatValid, 1) <> 3 Or UBound(P3CommittedSpmatValid, 2) < n - 1 Then Exit Function
  If UBound(P3CommittedSpmat, 1) <> 2 Or UBound(P3CommittedSpmat, 2) <> 15 Or UBound(P3CommittedSpmat, 3) <> 3 Or UBound(P3CommittedSpmat, 4) < n - 1 Then Exit Function
  If UBound(P3CommittedDpmat, 1) <> 2 Or UBound(P3CommittedDpmat, 2) <> 2 Or UBound(P3CommittedDpmat, 3) <> 3 Or UBound(P3CommittedDpmat, 4) < n - 1 Then Exit Function
  If UBound(P3ActivePlasticPoint, 1) <> 3 Or UBound(P3ActivePlasticPoint, 2) < n - 1 Then Exit Function
  For i = 1 To 2
    If LBound(P3CommittedPlasticMultiplier, i) <> 0 Or LBound(P3CommittedYieldFunction, i) <> 0 Or LBound(P3CommittedYielded, i) <> 0 Then Exit Function
    If LBound(P3CommittedSpmatValid, i) <> 0 Or LBound(P3ActivePlasticPoint, i) <> 0 Then Exit Function
  Next i
  For i = 1 To 3
    If LBound(P3CommittedPlasticStrain, i) <> 0 Or LBound(P3CommittedPrincipalStress, i) <> 0 Then Exit Function
  Next i
  For i = 1 To 4
    If LBound(P3CommittedSpmat, i) <> 0 Or LBound(P3CommittedDpmat, i) <> 0 Then Exit Function
  Next i
  dst(0) = VarPtr(Elem(0).Stmat(0, 0)): src(0) = VarPtr(Elem(0).mStmat(0, 0))
  dst(1) = VarPtr(Elem(0).P2PlasticStrain(0, 0)): src(1) = VarPtr(P3CommittedPlasticStrain(0, 0, 0))
  dst(2) = VarPtr(Elem(0).P2PlasticMultiplier(0)): src(2) = VarPtr(P3CommittedPlasticMultiplier(0, 0))
  dst(3) = VarPtr(Elem(0).P2YieldFunction(0)): src(3) = VarPtr(P3CommittedYieldFunction(0, 0))
  dst(4) = VarPtr(Elem(0).P2Yielded(0)): src(4) = VarPtr(P3CommittedYielded(0, 0))
  dst(5) = VarPtr(Elem(0).P2PrincipalStress(0, 0)): src(5) = VarPtr(P3CommittedPrincipalStress(0, 0, 0))
  dst(6) = VarPtr(Elem(0).SpmatValid(0)): src(6) = VarPtr(P3CommittedSpmatValid(0, 0))
  dst(7) = VarPtr(Elem(0).Spmat(0, 0, 0)): src(7) = VarPtr(P3CommittedSpmat(0, 0, 0, 0))
  dst(8) = VarPtr(Elem(0).Dpmat(0, 0, 0)): src(8) = VarPtr(P3CommittedDpmat(0, 0, 0, 0))
  If P3ElementActiveReady Then
    activeLo = LBound(P3ElementActive): activeHi = UBound(P3ElementActive)
    activePtr = VarPtr(P3ElementActive(activeLo))
  End If
  rc = SfnRestoreCommitted(n, stride, VarPtr(dst(0)), VarPtr(src(0)), activePtr, activeLo, activeHi, VarPtr(P3ActivePlasticPoint(0, 0)), P3ActivePlasticPointCount)
  If rc = 0 Then FEMNativeStateCopies = FEMNativeStateCopies + 1: FEMNativeRestoreState = True
  Exit Function
FailedState:
  If Err.number = 18 Then Err.Raise 18
  Err.Clear
#End If
End Function


' Build the current nonzero scatter list without changing matrix arithmetic.
' Missing setting/export or a rejected layout keeps the original VBA path.
Public Function FEMNativeCSRScatterList() As Boolean
#If Win64 Then
  Dim p(0 To 12) As LongLong, n As Long, count As Long, stride As LongLong
  On Error GoTo FailedAssembly
  If P6ReadSetting("NATIVE_ASSEMBLY", 0#) <> 1# Then Exit Function
  If FEMNativeModeText() <> "AUTO" Then Exit Function
  If Not FEMNativeAvailable() Then Exit Function
  If GetProcAddress(mHandle, "sfn_csr_scatter_list") = 0 Then Exit Function
  n = NumberOfElement
  If n < 1 Or n > 8388607 Or P6CSRNNZ < 1 Then Exit Function
  If LBound(Elem) <> 0 Or UBound(Elem) < n - 1 Then Exit Function
  If LBound(P6CSRElementPosition, 1) <> 0 Or UBound(P6CSRElementPosition, 1) <> n - 1 Then Exit Function
  If LBound(P6CSRElementPosition, 2) <> 0 Or UBound(P6CSRElementPosition, 2) <> 15 Then Exit Function
  If LBound(P6CSRElementPosition, 3) <> 0 Or UBound(P6CSRElementPosition, 3) <> 15 Then Exit Function
  stride = LenB(Elem(0))
  If n > 1 Then
    If VarPtr(Elem(1)) - VarPtr(Elem(0)) <> stride Then Exit Function
  End If
  If VarPtr(Elem(0).kmat(1, 0)) - VarPtr(Elem(0).kmat(0, 0)) <> 8 Then Exit Function
  If VarPtr(Elem(0).kmat(0, 1)) - VarPtr(Elem(0).kmat(0, 0)) <> 128 Then Exit Function
  ReDim P6CSRScatterStart(0 To n)
  p(0) = VarPtr(Elem(0)): p(1) = stride: p(2) = n
  p(3) = VarPtr(Elem(0).kmat(0, 0)) - VarPtr(Elem(0))
  p(4) = VarPtr(P6CSRElementPosition(0, 0, 0)): p(5) = P6CSRNNZ
  p(6) = VarPtr(P6CSRScatterStart(0)): p(7) = -1: p(12) = VarPtr(count)
  If SfnCSRScatterList(VarPtr(p(0))) <> 0 Then Exit Function
  If count < 0 Or count > n * 256 Then Exit Function
  If count > 0 Then
    ReDim P6CSRScatterLocalRow(0 To count - 1)
    ReDim P6CSRScatterLocalColumn(0 To count - 1)
    ReDim P6CSRScatterPosition(0 To count - 1)
    ReDim P6CSRScatterValueIndex(0 To count - 1)
    p(7) = count
    p(8) = VarPtr(P6CSRScatterLocalRow(0)): p(9) = VarPtr(P6CSRScatterLocalColumn(0))
    p(10) = VarPtr(P6CSRScatterPosition(0)): p(11) = VarPtr(P6CSRScatterValueIndex(0))
    If SfnCSRScatterList(VarPtr(p(0))) <> 0 Then Exit Function
  Else
    Erase P6CSRScatterLocalRow: Erase P6CSRScatterLocalColumn
    Erase P6CSRScatterPosition: Erase P6CSRScatterValueIndex
  End If
  P6CSRScatterCount = count
  FEMNativeAssemblyCalls = FEMNativeAssemblyCalls + 1
  FEMNativeCSRScatterList = True
  Exit Function
FailedAssembly:
  If Err.number = 18 Then Err.Raise 18
  Err.Clear
#End If
End Function

' Called at analysis start: re-read the setting, keep the loaded module.
Public Sub FEMNativeRefresh()
  mChecked = False: mAvailable = False
  FEMNativeResetCounts
  FEMMatRefresh
End Sub

' Effective mode: "OFF", "BAND" (band DLL only) or "AUTO" (band + material DLL).
Public Function FEMNativeModeText() As String
  Dim modeText As String
  If mForcedSet Then
    modeText = mForcedText
  Else
    modeText = UCase$(Trim$(P6ReadTextSetting("NATIVE_DLL", "AUTO")))
  End If
  If modeText = "OFF" Or modeText = "0" Or modeText = "VBA" Then
    FEMNativeModeText = "OFF"
  ElseIf modeText = "BAND" Then
    FEMNativeModeText = "BAND"
  ElseIf modeText = "MAT" Then
    FEMNativeModeText = "MAT"
  Else
    FEMNativeModeText = "AUTO"
  End If
End Function

Public Function FEMNativeReason() As String
  FEMNativeAvailable
  FEMNativeReason = mReason
End Function

Public Function FEMNativeStatusText() As String
  If FEMNativeAvailable() Then
    FEMNativeStatusText = "NATIVE(" & SFN_DLL & ";precision=" & mPrecisionText & ";factor=" & CStr(FEMNativeFactorCalls) & ";solve=" & CStr(FEMNativeSolveCalls) & ";nonsym=" & CStr(FEMNativeNsCalls) & ";statecopy=" & CStr(FEMNativeStateCopies) & ";gmres_simd=" & CStr(FEMNativeGmresCalls) & ";fallback=" & CStr(FEMNativeFallbacks) & ")"
  Else
    FEMNativeStatusText = "VBA(" & mReason & ")"
  End If
End Function

Public Function FEMNativeAvailable() As Boolean
  Dim path As String, loadedPath As String, size As Long, dllName As String, modeText As String
  Dim exportName As Variant
  Dim precisionText As String, precisionMode As Long
  On Error GoTo LoadFailed
  If mChecked Then
#If Win64 Then
    ' Other workbooks may share this DLL on the Excel thread.
    If mAvailable And mPrecisionAPI Then
      If SfnGetPrecision() <> mPrecisionMode Then
        If SfnSetPrecision(mPrecisionMode) <> 0 Then mAvailable = False: mReason = "DLL_PRECISION_REJECTED"
      End If
    End If
#End If
    FEMNativeAvailable = mAvailable: Exit Function
  End If
  mChecked = True: mAvailable = False
  On Error GoTo LoadFailed
  If mForcedSet Then
    If mForcedOff Then mReason = "OFF_SELECTED": Exit Function
  Else
    modeText = UCase$(Trim$(P6ReadTextSetting("NATIVE_DLL", "AUTO")))
    If modeText = "OFF" Or modeText = "0" Or modeText = "VBA" Then mReason = "OFF_SELECTED": Exit Function
  End If
#If Win64 Then
  path = ThisWorkbook.path & Application.PathSeparator & SFN_DLL
  If Len(Dir$(path)) = 0 Then mReason = "DLL_NOT_FOUND": Exit Function
  dllName = SFN_DLL
  If mHandle = 0 Then mHandle = GetModuleHandleW(StrPtr(dllName))
  If mHandle <> 0 Then
    loadedPath = String$(32768, vbNullChar)
    size = GetModuleFileNameW(mHandle, StrPtr(loadedPath), Len(loadedPath))
    If size = 0 Or size >= Len(loadedPath) Then mReason = "DLL_PATH_UNVERIFIED": Exit Function
    If StrComp(left$(loadedPath, size), path, vbTextCompare) <> 0 Then mReason = "DLL_ALREADY_LOADED_OTHER_PATH": Exit Function
  Else
    ' LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS
    mHandle = LoadLibraryExW(StrPtr(path), 0, &H1100)
  End If
  If mHandle = 0 Then mReason = "DLL_LOAD_FAILED_" & CStr(GetLastError()): Exit Function
  For Each exportName In Array("sfn_abi", "sfn_band_factor", "sfn_band_solve", "sfn_nsband_factor", "sfn_nsband_forward", "sfn_nsband_back")
    If GetProcAddress(mHandle, CStr(exportName)) = 0 Then mReason = "DLL_ENTRY_MISSING_" & CStr(exportName): Exit Function
  Next exportName
  If SfnAbi() <> SFN_ABI Then mReason = "DLL_ABI_MISMATCH": Exit Function
  ' Missing setting preserves the original VBA-compatible arithmetic.
  precisionText = UCase$(Trim$(P6ReadTextSetting("NATIVE_PRECISION", "STRICT")))
  Select Case precisionText
    Case "STRICT": precisionMode = 0
    Case "SSE2": precisionMode = 1
    Case "AUTO", "AVX2": precisionMode = 2
    Case Else: mReason = "INVALID_NATIVE_PRECISION": Exit Function
  End Select
  mPrecisionText = "STRICT_LEGACY"
  mPrecisionAPI = False
  If GetProcAddress(mHandle, "sfn_set_precision") <> 0 And GetProcAddress(mHandle, "sfn_get_precision") <> 0 Then
    If SfnSetPrecision(precisionMode) <> 0 Then mReason = "DLL_PRECISION_REJECTED": Exit Function
    Select Case SfnGetPrecision()
      Case 0: mPrecisionText = "STRICT"
      Case 1: mPrecisionText = "SSE2"
      Case 2: mPrecisionText = "AVX2_FMA"
      Case Else: mReason = "DLL_PRECISION_INVALID": Exit Function
    End Select
    mPrecisionMode = SfnGetPrecision(): mPrecisionAPI = True
  ElseIf precisionMode <> 0 Then
    mReason = "DLL_PRECISION_UNSUPPORTED": Exit Function
  End If
  mReason = "OK"
  mAvailable = True
  FEMNativeAvailable = True
#Else
  mReason = "NOT_WIN64"
#End If
  Exit Function
LoadFailed:
  If Err.number = 18 Then mChecked = False: Err.Raise 18 ' ユーザー中断は利用可否を確定せず呼出側へ送る
  mReason = "DLL_ERROR_" & CStr(Err.number)
  mAvailable = False
  FEMNativeAvailable = False
  Err.Clear
End Function

Private Sub FEMNativeDisable(ByVal why As String)
  mAvailable = False: mChecked = True: mReason = why
  FEMNativeFallbacks = FEMNativeFallbacks + 1
End Sub

' Return: 0 OK, 1 pivot failure (failK), -1 not handled (caller runs VBA).
Public Function FEMNativeBandFactor(ByRef band() As Double, ByRef packed() As Double, ByRef work() As Double, _
    ByVal n As Long, ByVal bw As Long, ByVal tol As Double, ByRef failK As Long) As Long
  Dim rc As Long, ld As Long
  FEMNativeBandFactor = -1
  If Not FEMNativeAvailable() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  If LBound(band, 1) <> 0 Or LBound(band, 2) <> 0 Or UBound(band, 2) < bw Then Exit Function
  ld = UBound(band, 1) + 1
  If ld < n Or LBound(packed) <> 0 Or UBound(packed) < CDbl(n) * CDbl(bw + 1) - 1 Then Exit Function
  If LBound(work) <> 0 Or UBound(work) < bw Then Exit Function
  rc = SfnBandFactor(n, bw, ld, VarPtr(band(0, 0)), VarPtr(packed(0)), VarPtr(work(0)), tol, failK)
  If rc = 0 Or rc = 1 Then
    FEMNativeFactorCalls = FEMNativeFactorCalls + 1
    FEMNativeBandFactor = rc
  Else
    FEMNativeDisable "BAND_FACTOR_RC_" & CStr(rc)
  End If
  Exit Function
Failed:
  If Err.number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  FEMNativeDisable "BAND_FACTOR_ERR_" & CStr(Err.number)
  Err.Clear
#End If
End Function

Public Function FEMNativeBandSolve(ByRef packed() As Double, ByRef rhs() As Double, ByRef x() As Double, _
    ByVal n As Long, ByVal bw As Long, ByVal tol As Double, ByRef failK As Long) As Long
  Dim rc As Long
  FEMNativeBandSolve = -1
  If Not FEMNativeAvailable() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  If LBound(packed) <> 0 Or UBound(packed) < CDbl(n) * CDbl(bw + 1) - 1 Then Exit Function
  If LBound(rhs) <> 0 Or UBound(rhs) < n - 1 Or LBound(x) <> 0 Or UBound(x) < n - 1 Then Exit Function
  rc = SfnBandSolve(n, bw, VarPtr(packed(0)), VarPtr(rhs(0)), VarPtr(x(0)), tol, failK)
  If rc = 0 Or rc = 1 Then
    FEMNativeSolveCalls = FEMNativeSolveCalls + 1
    FEMNativeBandSolve = rc
  Else
    FEMNativeDisable "BAND_SOLVE_RC_" & CStr(rc)
  End If
  Exit Function
Failed:
  If Err.number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  FEMNativeDisable "BAND_SOLVE_ERR_" & CStr(Err.number)
  Err.Clear
#End If
End Function

' Nonsymmetric pivoted band LU. a(0..n-1,0..3b), lower(0..n-1,0..b).
' Return: 0 OK, 1 rejected (failK, reason 1=ZERO_PIVOT 2=NONFINITE_MULTIPLIER),
' -1 not handled (arrays untouched; caller runs VBA).
Public Function FEMNativeNsFactor(ByRef a() As Double, ByRef rhs() As Double, ByRef pivotRows() As Long, ByRef lower() As Double, _
    ByVal n As Long, ByVal b As Long, ByRef failK As Long, ByRef reason As Long, ByRef largest As Double) As Long
  Dim rc As Long, ld As Long
  FEMNativeNsFactor = -1
  If Not FEMNativeAvailable() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  ld = UBound(a, 1) + 1
  If LBound(a, 1) <> 0 Or LBound(a, 2) <> 0 Or UBound(a, 2) < 3 * b Or ld <> n Then Exit Function
  If LBound(lower, 1) <> 0 Or UBound(lower, 1) + 1 <> ld Or UBound(lower, 2) < b Then Exit Function
  If LBound(pivotRows) <> 0 Or UBound(pivotRows) < n - 1 Or LBound(rhs) <> 0 Or UBound(rhs) < n - 1 Then Exit Function
  rc = SfnNsFactor(n, b, ld, VarPtr(a(0, 0)), VarPtr(rhs(0)), VarPtr(pivotRows(0)), VarPtr(lower(0, 0)), failK, reason, largest)
  If rc = 0 Or rc = 1 Then
    FEMNativeNsCalls = FEMNativeNsCalls + 1
    FEMNativeNsFactor = rc
  Else
    FEMNativeDisable "NS_FACTOR_RC_" & CStr(rc)
  End If
  Exit Function
Failed:
  If Err.number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  FEMNativeDisable "NS_FACTOR_ERR_" & CStr(Err.number)
  Err.Clear
#End If
End Function

Public Function FEMNativeNsForward(ByRef pivotRows() As Long, ByRef lower() As Double, ByRef rhs() As Double, ByVal n As Long, ByVal b As Long) As Long
  Dim rc As Long, ld As Long
  FEMNativeNsForward = -1
  If Not FEMNativeAvailable() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  ld = UBound(lower, 1) + 1
  If LBound(lower, 1) <> 0 Or LBound(lower, 2) <> 0 Or UBound(lower, 2) < b Or ld <> n Then Exit Function
  If LBound(pivotRows) <> 0 Or UBound(pivotRows) < n - 1 Or LBound(rhs) <> 0 Or UBound(rhs) < n - 1 Then Exit Function
  rc = SfnNsForward(n, b, ld, VarPtr(pivotRows(0)), VarPtr(lower(0, 0)), VarPtr(rhs(0)))
  If rc = 0 Then FEMNativeNsForward = 0 Else FEMNativeDisable "NS_FORWARD_RC_" & CStr(rc)
  Exit Function
Failed:
  If Err.number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  FEMNativeDisable "NS_FORWARD_ERR_" & CStr(Err.number)
  Err.Clear
#End If
End Function

' Return: 0 OK, 1 pivot (checkPivot), 2 nonfinite solution, -1 not handled.
Public Function FEMNativeNsBack(ByRef a() As Double, ByRef rhs() As Double, ByRef sol() As Double, ByVal n As Long, ByVal b As Long, _
    ByVal checkPivot As Boolean, ByRef failK As Long) As Long
  Dim rc As Long, ld As Long
  FEMNativeNsBack = -1
  If Not FEMNativeAvailable() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  ld = UBound(a, 1) + 1
  If LBound(a, 1) <> 0 Or LBound(a, 2) <> 0 Or UBound(a, 2) < 3 * b Or ld <> n Then Exit Function
  If LBound(rhs) <> 0 Or UBound(rhs) < n - 1 Or LBound(sol) <> 0 Or UBound(sol) < n - 1 Then Exit Function
  rc = SfnNsBack(n, b, ld, VarPtr(a(0, 0)), VarPtr(rhs(0)), VarPtr(sol(0)), IIf(checkPivot, 1, 0), failK)
  If rc >= 0 And rc <= 2 Then FEMNativeNsBack = rc Else FEMNativeDisable "NS_BACK_RC_" & CStr(rc)
  Exit Function
Failed:
  If Err.number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  FEMNativeDisable "NS_BACK_ERR_" & CStr(Err.number)
  Err.Clear
#End If
End Function
