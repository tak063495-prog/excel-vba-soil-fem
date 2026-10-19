Option Explicit
' Phase C native kernels in soilfem_mat_x64.dll (API 4), active only in
' NATIVE_DLL=AUTO (MAT = Phase B material kernels only, BAND, OFF).
' Each wrapper returns False (or -1) when it did not handle the request; the
' caller then runs the unchanged VBA code, so logs/failures stay identical.
#If Win64 Then
Private Declare PtrSafe Function SfmBandScatter Lib "soilfem_mat_x64.dll" Alias "sfm_band_scatter" (ByVal n As Long, ByVal r As LongPtr, ByVal c As LongPtr, ByVal e As LongPtr, ByVal li As LongPtr, ByVal lj As LongPtr, ByVal band As LongPtr, ByVal ld As LongLong, ByVal elemBase As LongPtr, ByVal stride As LongLong, ByVal kmatOff As LongLong) As Long
Private Declare PtrSafe Function SfmCsrScatter Lib "soilfem_mat_x64.dll" Alias "sfm_csr_scatter" (ByVal params As LongPtr, ByVal nnz As Long, ByVal st As LongPtr, ByVal pos As LongPtr, ByVal lr As LongPtr, ByVal lc As LongPtr, ByVal vals As LongPtr, ByVal kmatOff As LongLong) As Long
Private Declare PtrSafe Function SfmCsrResNorm Lib "soilfem_mat_x64.dll" Alias "sfm_csr_resnorm" (ByVal n As Long, ByVal rp As LongPtr, ByVal ci As LongPtr, ByVal v As LongPtr, ByVal x As LongPtr, ByVal b As LongPtr, ByVal r As LongPtr, ByRef norm As Double) As Long
Private Declare PtrSafe Function SfmIlu0 Lib "soilfem_mat_x64.dll" Alias "sfm_ilu0" (ByVal n As Long, ByVal nnz As Long, ByVal rp As LongPtr, ByVal ci As LongPtr, ByVal dp As LongPtr, ByVal vals As LongPtr, ByVal ilu As LongPtr, ByVal fail As LongPtr, ByRef minAbs As Double) As Long
Private Declare PtrSafe Function SfmGmres Lib "soilfem_mat_x64.dll" Alias "sfm_gmres" (ByVal params As LongPtr, ByVal tol As Double) As Long
Private Declare PtrSafe Function SfmRegularized Lib "soilfem_mat_x64.dll" Alias "sfm_regularized" (ByVal params As LongPtr, ByVal damping As Double, ByVal itTol As Double) As Long
Private Declare PtrSafe Function SfmResidual Lib "soilfem_mat_x64.dll" Alias "sfm_residual" (ByVal params As LongPtr, ByVal forceRef As Double) As Long
Private Declare PtrSafe Function SfmRcm Lib "soilfem_mat_x64.dll" Alias "sfm_rcm" (ByVal nNode As Long, ByVal nElem As Long, ByVal nodes As LongPtr, ByVal elnode As LongPtr, ByVal o2n As LongPtr, ByVal n2o As LongPtr, ByRef candBw As Long) As Long
Private Declare PtrSafe Function SfmHgApply Lib "soilfem_mat_x64.dll" Alias "sfm_hg_apply" (ByVal params As LongPtr, ByVal tdone As LongPtr, ByVal hgdone As LongPtr, ByVal unsym As LongPtr, ByVal hoff As LongPtr, ByVal shapeGen As Long, ByVal hgFactor As Double, ByVal applyHg As Long) As Long
Private mG(0 To 15) As LongLong
Private mHOff(0 To 7) As LongLong
Private mOutD(0 To 15) As Double
#End If
Public FEMKernScatterCalls As Long, FEMKernGmresCalls As Long, FEMKernIluCalls As Long, FEMKernRegCalls As Long
Public FEMKernResidualCalls As Long, FEMKernRcmCalls As Long, FEMKernResNormCalls As Long, FEMKernDeclines As Long
Private mHgDone() As Integer, mUnsym() As Integer, mHgOK As Boolean
Public FEMKernHgCalls As Long, FEMKernHgApplied As Long
Private mWork() As Double, mEv() As Double, mVec() As Double, mBand() As Double

Public Sub FEMKernResetCounts()
  FEMKernScatterCalls = 0: FEMKernGmresCalls = 0: FEMKernIluCalls = 0: FEMKernRegCalls = 0
  FEMKernResidualCalls = 0: FEMKernRcmCalls = 0: FEMKernResNormCalls = 0: FEMKernDeclines = 0
  FEMKernHgCalls = 0: FEMKernHgApplied = 0
End Sub

Public Function FEMKernOn() As Boolean
  FEMKernOn = False
  If FEMNativeModeText() <> "AUTO" Then Exit Function
  FEMKernOn = FEMMatAvailable()
End Function

Public Function FEMKernStatusText() As String
  If FEMKernOn() Then
    FEMKernStatusText = "KERN(scatter=" & CStr(FEMKernScatterCalls) & ";residual=" & CStr(FEMKernResidualCalls) & ";resnorm=" & CStr(FEMKernResNormCalls) & _
      ";ilu=" & CStr(FEMKernIluCalls) & ";gmres=" & CStr(FEMKernGmresCalls) & ";regularized=" & CStr(FEMKernRegCalls) & ";rcm=" & CStr(FEMKernRcmCalls) & ";hg=" & CStr(FEMKernHgCalls) & "/" & CStr(FEMKernHgApplied) & ";declined=" & CStr(FEMKernDeclines) & ")"
  Else
    FEMKernStatusText = "KERN(OFF)"
  End If
End Function

' SetTotalMat: OriginalMat(row,col) += kmat over the scatter list (same order).
Public Function FEMKernBandScatter(ByRef band() As Double, ByVal ld As Long, ByVal n As Long, ByRef sr() As Long, ByRef sc() As Long, _
    ByRef se() As Long, ByRef si() As Long, ByRef sj() As Long) As Boolean
  FEMKernBandScatter = False
  If Not FEMKernOn() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  If n < 1 Then Exit Function
  If SfmBandScatter(n, VarPtr(sr(0)), VarPtr(sc(0)), VarPtr(se(0)), VarPtr(si(0)), VarPtr(sj(0)), VarPtr(band(0, 0)), ld, VarPtr(Elem(0)), FEMMatOffset(0), FEMMatOffset(17)) <> 0 Then Exit Function
  FEMKernScatterCalls = FEMKernScatterCalls + 1
  FEMKernBandScatter = True
  Exit Function
Failed:
  If Err.Number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  Err.Clear: FEMKernDeclines = FEMKernDeclines + 1
#End If
End Function

Public Function FEMKernCsrScatter(ByVal nnz As Long, ByRef st() As Long, ByRef pos() As Long, ByRef lr() As Long, ByRef lc() As Long, ByRef vals() As Double) As Boolean
  FEMKernCsrScatter = False
  If Not FEMKernOn() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  If nnz < 1 Or UBound(st) < NumberOfElement Then Exit Function
  If st(NumberOfElement) < 1 Then Exit Function
  If SfmCsrScatter(FEMMatParamsPtr(), nnz, VarPtr(st(0)), VarPtr(pos(0)), VarPtr(lr(0)), VarPtr(lc(0)), VarPtr(vals(0)), FEMMatOffset(17)) <> 0 Then FEMKernDeclines = FEMKernDeclines + 1: Exit Function
  FEMKernScatterCalls = FEMKernScatterCalls + 1
  FEMKernCsrScatter = True
  Exit Function
Failed:
  If Err.Number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  Err.Clear: FEMKernDeclines = FEMKernDeclines + 1
#End If
End Function

Public Function FEMKernCsrResNorm(ByVal n As Long, ByRef rp() As Long, ByRef ci() As Long, ByRef v() As Double, ByRef x() As Double, ByRef b() As Double, ByRef r() As Double, ByRef normOut As Double) As Boolean
  FEMKernCsrResNorm = False
  If Not FEMKernOn() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  If SfmCsrResNorm(n, VarPtr(rp(0)), VarPtr(ci(0)), VarPtr(v(0)), VarPtr(x(0)), VarPtr(b(0)), VarPtr(r(0)), normOut) <> 0 Then Exit Function
  FEMKernResNormCalls = FEMKernResNormCalls + 1
  FEMKernCsrResNorm = True
  Exit Function
Failed:
  If Err.Number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  Err.Clear: FEMKernDeclines = FEMKernDeclines + 1
#End If
End Function

' Returns 1 built, 0 failed (failRow/failCol/minAbs set), -1 not handled.
Public Function FEMKernIlu0(ByVal n As Long, ByVal nnz As Long, ByRef rp() As Long, ByRef ci() As Long, ByRef dp() As Long, ByRef vals() As Double, ByRef ilu() As Double, _
    ByRef failRow As Long, ByRef failCol As Long, ByRef minAbs As Double) As Long
  Dim f(0 To 1) As Long
  FEMKernIlu0 = -1
  If Not FEMKernOn() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  FEMKernIlu0 = SfmIlu0(n, nnz, VarPtr(rp(0)), VarPtr(ci(0)), VarPtr(dp(0)), VarPtr(vals(0)), VarPtr(ilu(0)), VarPtr(f(0)), minAbs)
  failRow = f(0): failCol = f(1)
  FEMKernIluCalls = FEMKernIluCalls + 1
  Exit Function
Failed:
  If Err.Number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  Err.Clear: FEMKernDeclines = FEMKernDeclines + 1: FEMKernIlu0 = -1
#End If
End Function

' GMRES(m)+ILU0. Returns 1/0 like the VBA routine, -1 not handled.
' Events: (type, iterations, relres) triples replayed by the caller.
Public Function FEMKernGmres(ByVal n As Long, ByRef rp() As Long, ByRef ci() As Long, ByRef vals() As Double, ByRef ilu() As Double, ByRef rhs() As Double, _
    ByRef x() As Double, ByVal restart As Long, ByVal maxIt As Long, ByVal tol As Double, ByRef lastRes As Double, ByRef iters As Long, ByRef ev() As Double, ByRef nEv As Long) As Long
  Dim need As Double, cap As Long
  FEMKernGmres = -1
  If Not FEMKernOn() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  need = CDbl(n) * (2# * restart + 5#) + CDbl(restart + 1) * restart + 4# * restart + 2#
  If n < 1 Or restart < 1 Or maxIt < 1 Then Exit Function
  If LBound(rp) <> 0 Or UBound(rp) < n Or LBound(ci) <> 0 Or LBound(vals) <> 0 Or LBound(ilu) <> 0 Then Exit Function
  If LBound(rhs) <> 0 Or UBound(rhs) < n - 1 Or LBound(x) <> 0 Or UBound(x) < n - 1 Then Exit Function
  If LBound(P6CSRDiagonalPosition) <> 0 Or UBound(P6CSRDiagonalPosition) < n - 1 Then Exit Function
  If rp(n) < 1 Or UBound(ci) < rp(n) - 1 Or UBound(vals) < rp(n) - 1 Or UBound(ilu) < rp(n) - 1 Then Exit Function
  If need > 200000000# Then Exit Function
  ReDim mWork(0 To CLng(need))
  cap = maxIt + 8
  ReDim ev(0 To 3 * cap - 1)
  mG(0) = n: mG(1) = VarPtr(rp(0)): mG(2) = VarPtr(ci(0)): mG(3) = VarPtr(vals(0)): mG(4) = VarPtr(ilu(0))
  mG(5) = VarPtr(rhs(0)): mG(6) = VarPtr(x(0)): mG(7) = VarPtr(mWork(0)): mG(8) = restart: mG(9) = maxIt
  mG(10) = VarPtr(ev(0)): mG(11) = cap: mG(12) = VarPtr(mOutD(0))
  mG(13) = VarPtr(P6CSRDiagonalPosition(0)): mG(14) = UBound(mWork) + 1: mG(15) = rp(n)
  FEMKernGmres = FEMNativeGmres(VarPtr(mG(0)), tol)
  If FEMKernGmres < 0 Then FEMKernGmres = SfmGmres(VarPtr(mG(0)), tol)
  If FEMKernGmres < 0 Or mOutD(3) <> 0# Then FEMKernGmres = -1: FEMKernDeclines = FEMKernDeclines + 1: Exit Function
  lastRes = mOutD(0): iters = CLng(mOutD(1)): nEv = CLng(mOutD(2))
  FEMKernGmresCalls = FEMKernGmresCalls + 1
  Exit Function
Failed:
  If Err.Number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  Err.Clear: FEMKernDeclines = FEMKernDeclines + 1: FEMKernGmres = -1
#End If
End Function

' Regularized LM correction numeric core. Returns 1 accepted, 0 rejected, -1 not handled.
Public Function FEMKernRegularized(ByVal n As Long, ByVal bw As Long, ByRef rp() As Long, ByRef ci() As Long, ByRef orig() As Double, ByRef nc() As Long, _
    ByRef f() As Double, ByRef d() As Double, ByVal damping As Double, ByVal itTol As Double, ByRef normalRes As Double, ByRef factorDone As Boolean) As Long
  FEMKernRegularized = -1
  If Not FEMKernOn() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  ReDim mBand(0 To n - 1, 0 To bw)
  ReDim mVec(0 To 3 * n - 1)
  mG(0) = n: mG(1) = VarPtr(rp(0)): mG(2) = VarPtr(ci(0)): mG(3) = VarPtr(orig(0)): mG(4) = VarPtr(nc(0))
  mG(5) = VarPtr(f(0)): mG(6) = VarPtr(d(0)): mG(7) = VarPtr(mBand(0, 0)): mG(8) = bw: mG(9) = VarPtr(mVec(0)): mG(10) = VarPtr(mOutD(0))
  FEMKernRegularized = SfmRegularized(VarPtr(mG(0)), damping, itTol)
  normalRes = mOutD(0): factorDone = (mOutD(1) <> 0#)
  Erase mBand
  If FEMKernRegularized >= 0 Then FEMKernRegCalls = FEMKernRegCalls + 1
  Exit Function
Failed:
  If Err.Number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  Err.Clear: FEMKernDeclines = FEMKernDeclines + 1: FEMKernRegularized = -1
#End If
End Function

' Residual metrics (P3AccumulateResidualMetrics, non-mixed). out(0..9) as in C.
Public Function FEMKernResidual(ByRef internalForce() As Double, ByRef targetForce() As Double, ByVal lite As Boolean, ByVal forceRef As Double, ByRef outV() As Double) As Boolean
  Dim i As Long
  FEMKernResidual = False
  If Not FEMKernOn() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  If LBound(internalForce) <> 0 Or LBound(targetForce) <> 0 Or UBound(internalForce) < lastDof Or UBound(targetForce) < lastDof Then Exit Function
  If UBound(NodeCond) < lastDof Or UBound(TDisp) < lastDof Or UBound(P3BoundaryDisp) < lastDof Then Exit Function
  mG(0) = lastDof + 1: mG(1) = VarPtr(internalForce(0)): mG(2) = VarPtr(targetForce(0)): mG(3) = VarPtr(NodeCond(0)): mG(4) = VarPtr(TDisp(0))
  mG(5) = 0
  If Not lite Then
    If UBound(Reaction) < lastDof Then Exit Function
    mG(5) = VarPtr(Reaction(0))
  End If
  mG(6) = VarPtr(P3BoundaryDisp(0)): mG(7) = 0
  If P3StageStartDispReady Then
    If UBound(P3StageStartDisp) < lastDof Then Exit Function
    mG(7) = VarPtr(P3StageStartDisp(0))
  End If
  mG(8) = VarPtr(mOutD(0))
  If SfmResidual(VarPtr(mG(0)), forceRef) <> 0 Then Exit Function
  For i = 0 To 9: outV(i) = mOutD(i): Next i
  FEMKernResidualCalls = FEMKernResidualCalls + 1
  FEMKernResidual = True
  Exit Function
Failed:
  If Err.Number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  Err.Clear: FEMKernDeclines = FEMKernDeclines + 1
#End If
End Function

' RCM ordering + candidate bandwidth (P6BuildRCMOrdering/P6ComputeRCMCandidateBandwidth).
Public Function FEMKernRcm(ByRef oldToNew() As Long, ByRef newToOld() As Long, ByRef candBw As Long) As Boolean
  Dim nodes() As Long, eln() As Long, k As Long, i As Long
  FEMKernRcm = False
  If Not FEMKernOn() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  If NumberOfFreeNode <= 0 Or NumberOfElement <= 0 Then Exit Function
  ReDim nodes(0 To NumberOfElement * 8 - 1): ReDim eln(0 To NumberOfElement * 16 - 1)
  For k = 0 To NumberOfElement - 1
    For i = 0 To 7: nodes(k * 8 + i) = Elem(k).node(i): Next i
    For i = 0 To 15: eln(k * 16 + i) = Elem(k).ElNode(i): Next i
  Next k
  ReDim oldToNew(0 To NumberOfFreeNode - 1): ReDim newToOld(0 To NumberOfFreeNode - 1)
  If SfmRcm(NumberOfFreeNode, NumberOfElement, VarPtr(nodes(0)), VarPtr(eln(0)), VarPtr(oldToNew(0)), VarPtr(newToOld(0)), candBw) <> 0 Then FEMKernDeclines = FEMKernDeclines + 1: Exit Function
  FEMKernRcmCalls = FEMKernRcmCalls + 1
  FEMKernRcm = True
  Exit Function
Failed:
  If Err.Number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  Err.Clear: FEMKernDeclines = FEMKernDeclines + 1
#End If
End Function

' Hourglass (cached-mode path) + unsymmetry check for every element whose
' tangent the material DLL just built. Returns elements with hourglass applied
' (for P6HourglassReuseCount), -1 when not handled.
Public Function FEMKernHgApply(ByVal applyHg As Boolean) As Long
  Dim b As LongPtr, dp As LongPtr
  FEMKernHgApply = -1: mHgOK = False
  If Not FEMKernOn() Then Exit Function
#If Win64 Then
  On Error GoTo Failed
  dp = FEMMatDonePtr()
  If dp = 0 Or NumberOfElement < 1 Then Exit Function
  b = VarPtr(Elem(0))
  mHOff(0) = VarPtr(Elem(0).HourglassShapeGen) - b
  mHOff(1) = VarPtr(Elem(0).Dmat(2, 2)) - b
  mHOff(2) = VarPtr(Elem(0).HourglassKReady) - b
  mHOff(3) = VarPtr(Elem(0).HourglassKScaledReady) - b
  mHOff(4) = VarPtr(Elem(0).HourglassCachedScale) - b
  mHOff(5) = VarPtr(Elem(0).HourglassKScaled(0, 0)) - b
  mHOff(6) = VarPtr(Elem(0).HourglassScale) - b
  mHOff(7) = VarPtr(Elem(0).HourglassModeCount) - b
  If VarPtr(Elem(0).HourglassKScaled(1, 0)) - VarPtr(Elem(0).HourglassKScaled(0, 0)) <> 8 Then Exit Function
  ReDim mHgDone(0 To NumberOfElement - 1): ReDim mUnsym(0 To NumberOfElement - 1)
  FEMKernHgApply = SfmHgApply(FEMMatParamsPtr(), dp, VarPtr(mHgDone(0)), VarPtr(mUnsym(0)), VarPtr(mHOff(0)), P6HgShapeGeneration, P6HourglassFactor, IIf(applyHg, 1, 0))
  If FEMKernHgApply < 0 Then Exit Function
  FEMKernHgCalls = FEMKernHgCalls + 1: FEMKernHgApplied = FEMKernHgApplied + FEMKernHgApply
  mHgOK = True
  Exit Function
Failed:
  If Err.Number = 18 Then Err.Raise 18 ' ユーザー中断はDLL無効化・VBA代替へ進めず呼出側へ送る
  Err.Clear: FEMKernDeclines = FEMKernDeclines + 1: FEMKernHgApply = -1
#End If
End Function

Public Function FEMKernHgDone(ByVal k As Long) As Boolean
  If Not mHgOK Then Exit Function
  FEMKernHgDone = (mHgDone(k) <> 0)
End Function

Public Function FEMKernUnsym(ByVal k As Long) As Boolean
  If Not mHgOK Then Exit Function
  FEMKernUnsym = (mUnsym(k) <> 0)
End Function
