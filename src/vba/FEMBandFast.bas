Option Explicit
' Optional VBA-compatible band kernel; analysis settings and failure rules stay in VBA.
#If Win64 Then
Private Declare PtrSafe Function LoadLibraryExW Lib "kernel32" (ByVal path As LongPtr, ByVal file As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal h As LongPtr, ByVal name As String) As LongPtr
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal name As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleFileNameW Lib "kernel32" (ByVal h As LongPtr, ByVal buffer As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Function FastAbi Lib "soilfem_band_fast_x64.dll" Alias "sfbs_abi" () As Long
Private Declare PtrSafe Function FastProbe Lib "soilfem_band_fast_x64.dll" Alias "sfbs_probe" (ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal op As Long) As Double
Private Declare PtrSafe Function FastFactor Lib "soilfem_band_fast_x64.dll" Alias "sfbs_factor" (ByVal n As Long, ByVal bw As Long, ByVal ld As Long, ByVal band As LongPtr, ByVal packed As LongPtr, ByVal work As LongPtr, ByVal tol As Double, ByRef failK As Long) As Long
Private Declare PtrSafe Function FastSolve Lib "soilfem_band_fast_x64.dll" Alias "sfbs_solve" (ByVal n As Long, ByVal bw As Long, ByVal packed As LongPtr, ByVal rhs As LongPtr, ByVal x As LongPtr, ByVal tol As Double, ByRef failK As Long) As Long
Private mHandle As LongPtr
#End If
Private Const FastDllFile As String = "soilfem_band_fast_x64.dll"
Private mChecked As Boolean, mAvailable As Boolean, mReason As String
Public FEMBandFastFactorCalls As Long, FEMBandFastSolveCalls As Long

Public Sub FEMBandFastReset()
    mChecked = False: mAvailable = False: mReason = ""
    FEMBandFastFactorCalls = 0: FEMBandFastSolveCalls = 0
End Sub

Public Function FEMBandFastAvailable() As Boolean
    Dim mode As String
    mode = FEMNativeModeText()
    If mode = "OFF" Or mode = "MAT" Then mReason = "OFF_SELECTED": Exit Function
    If mChecked Then FEMBandFastAvailable = mAvailable: Exit Function
    mChecked = True: mAvailable = False
#If Win64 Then
    Dim path As String, loaded As String, size As Long, name As Variant
    On Error GoTo Failed
    path = ThisWorkbook.Path & Application.PathSeparator & FastDllFile
    If Len(Dir$(path)) = 0 Then mReason = "DLL_NOT_FOUND": Exit Function
    If mHandle = 0 Then mHandle = GetModuleHandleW(StrPtr(FastDllFile))
    If mHandle <> 0 Then
        loaded = String$(32768, vbNullChar)
        size = GetModuleFileNameW(mHandle, StrPtr(loaded), Len(loaded))
        If size = 0 Or size >= Len(loaded) Then mReason = "DLL_PATH_UNVERIFIED": Exit Function
        If StrComp(Left$(loaded, size), path, vbTextCompare) <> 0 Then mReason = "DLL_ALREADY_LOADED_OTHER_PATH": Exit Function
    Else
        mHandle = LoadLibraryExW(StrPtr(path), 0, &H1100)
    End If
    If mHandle = 0 Then mReason = "DLL_LOAD_FAILED": Exit Function
    For Each name In Array("sfbs_abi", "sfbs_probe", "sfbs_factor", "sfbs_solve")
        If GetProcAddress(mHandle, CStr(name)) = 0 Then mReason = "DLL_ENTRY_MISSING": Exit Function
    Next name
    If FastAbi() <> &H10002 Then mReason = "DLL_ABI_MISMATCH": Exit Function
    If Not RoundingCompatible() Then mReason = "ROUNDING_MISMATCH": Exit Function
    mAvailable = True: mReason = "STRICT_X87"
#Else
    mReason = "NOT_WIN64"
#End If
    FEMBandFastAvailable = mAvailable
    Exit Function
Failed:
    If Err.Number = 18 Then mChecked = False: Err.Raise 18
    mReason = "DLL_ERROR_" & CStr(Err.Number): mAvailable = False: Err.Clear
End Function

Private Function RoundingCompatible() As Boolean
#If Win64 Then
    Dim v As Variant, a As Double, b As Double, c As Double, expected As Double, op As Long
    For Each v In Array(Array(6.51592972722763, 7.88723351135513, 0.938595867742349), Array(2.55069025739422, 4.95435087091941, 4.49491064788738), Array(7.21540032340783, 2.28762221270453, 9.45270695553922))
        a = v(0): b = v(1): c = v(2)
        For op = 0 To 3
            Select Case op
                Case 0: expected = a - b * c
                Case 1: expected = a + b * c
                Case 2: expected = a - b / c
                Case 3: expected = b / c
            End Select
            If expected <> FastProbe(a, b, c, op) Then Exit Function
        Next op
    Next v
    RoundingCompatible = True
#End If
End Function

Public Function FEMBandFastStatus() As String
    If FEMBandFastAvailable() Then
        FEMBandFastStatus = "NATIVE(" & FastDllFile & ";STRICT_X87;factor=" & CStr(FEMBandFastFactorCalls) & ";solve=" & CStr(FEMBandFastSolveCalls) & ")"
    Else
        FEMBandFastStatus = "VBA(" & mReason & ")"
    End If
End Function

' Return -1 only when the unchanged VBA/legacy-DLL path must run.
' Return 1 preserves the original pivot rejection; it must never be hidden by fallback.
Public Function FEMBandFastFactor(ByRef band() As Double, ByRef packed() As Double, ByRef work() As Double, ByVal n As Long, ByVal bw As Long, ByVal tol As Double, ByRef failK As Long) As Long
    FEMBandFastFactor = -1
    If Not FEMBandFastAvailable() Then Exit Function
#If Win64 Then
    Dim rc As Long, ld As Long
    On Error GoTo Failed
    If n < 1 Or bw < 0 Then Exit Function
    If LBound(band, 1) <> 0 Or LBound(band, 2) <> 0 Or UBound(band, 2) < bw Then Exit Function
    ld = UBound(band, 1) + 1
    If ld < n Or LBound(packed) <> 0 Or UBound(packed) < CDbl(n) * CDbl(bw + 1) - 1 Then Exit Function
    If LBound(work) <> 0 Or UBound(work) < bw Then Exit Function
    rc = FastFactor(n, bw, ld, VarPtr(band(0, 0)), VarPtr(packed(0)), VarPtr(work(0)), tol, failK)
    If rc = 0 Or rc = 1 Then
        FEMBandFastFactorCalls = FEMBandFastFactorCalls + 1
        FEMBandFastFactor = rc
    Else
        mAvailable = False: mReason = "FACTOR_ARGUMENTS_REJECTED"
    End If
#End If
    Exit Function
Failed:
    If Err.Number = 18 Then Err.Raise 18
    mAvailable = False: mReason = "FACTOR_ERROR_" & CStr(Err.Number): Err.Clear
End Function

Public Function FEMBandFastSolve(ByRef packed() As Double, ByRef rhs() As Double, ByRef x() As Double, ByVal n As Long, ByVal bw As Long, ByVal tol As Double, ByRef failK As Long) As Long
    FEMBandFastSolve = -1
    If Not FEMBandFastAvailable() Then Exit Function
#If Win64 Then
    Dim rc As Long
    On Error GoTo Failed
    If n < 1 Or bw < 0 Then Exit Function
    If LBound(packed) <> 0 Or UBound(packed) < CDbl(n) * CDbl(bw + 1) - 1 Then Exit Function
    If LBound(rhs) <> 0 Or UBound(rhs) < n - 1 Or LBound(x) <> 0 Or UBound(x) < n - 1 Then Exit Function
    rc = FastSolve(n, bw, VarPtr(packed(0)), VarPtr(rhs(0)), VarPtr(x(0)), tol, failK)
    If rc = 0 Or rc = 1 Then
        FEMBandFastSolveCalls = FEMBandFastSolveCalls + 1
        FEMBandFastSolve = rc
    Else
        mAvailable = False: mReason = "SOLVE_ARGUMENTS_REJECTED"
    End If
#End If
    Exit Function
Failed:
    If Err.Number = 18 Then Err.Raise 18
    mAvailable = False: mReason = "SOLVE_ERROR_" & CStr(Err.Number): Err.Clear
End Function
