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
Public FEMNativeFactorCalls As Long, FEMNativeSolveCalls As Long, FEMNativeNsCalls As Long, FEMNativeFallbacks As Long

' Test/diagnostic switch: "AUTO", "BAND", "OFF" or "" (= follow NATIVE_DLL setting).
Public Sub FEMNativeSetMode(ByVal modeText As String)
  modeText = UCase$(Trim$(modeText))
  mForcedSet = (modeText = "AUTO" Or modeText = "OFF" Or modeText = "BAND")
  mForcedOff = (modeText = "OFF")
  mForcedText = modeText
  mChecked = False: mAvailable = False
  FEMNativeResetCounts
  FEMMatRefresh
End Sub

Public Sub FEMNativeResetCounts()
  FEMNativeFactorCalls = 0: FEMNativeSolveCalls = 0: FEMNativeNsCalls = 0: FEMNativeFallbacks = 0
End Sub

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
    FEMNativeStatusText = "NATIVE(" & SFN_DLL & ";factor=" & CStr(FEMNativeFactorCalls) & ";solve=" & CStr(FEMNativeSolveCalls) & ";nonsym=" & CStr(FEMNativeNsCalls) & ";fallback=" & CStr(FEMNativeFallbacks) & ")"
  Else
    FEMNativeStatusText = "VBA(" & mReason & ")"
  End If
End Function

Public Function FEMNativeAvailable() As Boolean
  Dim path As String, loadedPath As String, size As Long, dllName As String, modeText As String
  Dim exportName As Variant
  If mChecked Then FEMNativeAvailable = mAvailable: Exit Function
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
    If StrComp(Left$(loadedPath, size), path, vbTextCompare) <> 0 Then mReason = "DLL_ALREADY_LOADED_OTHER_PATH": Exit Function
  Else
    ' LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS
    mHandle = LoadLibraryExW(StrPtr(path), 0, &H1100)
  End If
  If mHandle = 0 Then mReason = "DLL_LOAD_FAILED_" & CStr(GetLastError()): Exit Function
  For Each exportName In Array("sfn_abi", "sfn_band_factor", "sfn_band_solve", "sfn_nsband_factor", "sfn_nsband_forward", "sfn_nsband_back")
    If GetProcAddress(mHandle, CStr(exportName)) = 0 Then mReason = "DLL_ENTRY_MISSING_" & CStr(exportName): Exit Function
  Next exportName
  If SfnAbi() <> SFN_ABI Then mReason = "DLL_ABI_MISMATCH": Exit Function
  mReason = "OK"
  mAvailable = True
  FEMNativeAvailable = True
#Else
  mReason = "NOT_WIN64"
#End If
  Exit Function
LoadFailed:
  mReason = "DLL_ERROR_" & CStr(Err.Number)
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
  FEMNativeDisable "BAND_FACTOR_ERR_" & CStr(Err.Number)
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
  FEMNativeDisable "BAND_SOLVE_ERR_" & CStr(Err.Number)
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
  FEMNativeDisable "NS_FACTOR_ERR_" & CStr(Err.Number)
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
  FEMNativeDisable "NS_FORWARD_ERR_" & CStr(Err.Number)
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
  FEMNativeDisable "NS_BACK_ERR_" & CStr(Err.Number)
  Err.Clear
#End If
End Function
