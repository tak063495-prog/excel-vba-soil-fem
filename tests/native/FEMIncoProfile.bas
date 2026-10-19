Option Explicit

Private Type ProfileNode
    Id As Long
    Calls As Long
    Inclusive As Double
    Child As Double
End Type
Private nodes(1 To 64) As ProfileNode
Private stackId(1 To 512) As Long
Private stackStart(1 To 512) As Currency
Private stackChild(1 To 512) As Currency
Private depth As Long
Private freq As Currency
Private active As Boolean

Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long

Public Sub FEMProfileStart()
    Dim i As Long
    For i = 1 To 64: nodes(i).Id = 0: nodes(i).Calls = 0: nodes(i).Inclusive = 0: nodes(i).Child = 0: Next
    depth = 0: active = False: freq = 0
    QueryPerformanceFrequency freq
    If freq <= 0 Then Err.Raise vbObjectError + 6101, "FEMProfileStart", "QPC frequency unavailable"
    active = True
End Sub

Public Sub FEMProfileEnter(ByVal id As Long)
    Dim q As Currency
    If Not active Then Exit Sub
    If id < 1 Or id > 64 Then Err.Raise vbObjectError + 6102, "FEMProfileEnter", "invalid id"
    If depth >= 512 Then Err.Raise vbObjectError + 6103, "FEMProfileEnter", "profile stack overflow"
    QueryPerformanceCounter q: depth = depth + 1
    stackId(depth) = id: stackStart(depth) = q: stackChild(depth) = 0
    nodes(id).Id = id: nodes(id).Calls = nodes(id).Calls + 1
End Sub

Public Sub FEMProfileLeave(ByVal id As Long)
    Dim q As Currency, elapsed As Currency, exclusive As Currency
    If Not active Then Exit Sub
    If depth < 1 Then Err.Raise vbObjectError + 6104, "FEMProfileLeave", "empty profile stack"
    If stackId(depth) <> id Then Err.Raise vbObjectError + 6104, "FEMProfileLeave", "unbalanced profile stack"
    QueryPerformanceCounter q: elapsed = q - stackStart(depth): exclusive = elapsed - stackChild(depth)
    nodes(id).Inclusive = nodes(id).Inclusive + elapsed / freq: nodes(id).Child = nodes(id).Child + exclusive / freq
    If depth > 1 Then stackChild(depth - 1) = stackChild(depth - 1) + elapsed
    depth = depth - 1
End Sub

Public Function FEMProfileCsv() As String
    Dim i As Long, s As String
    If depth <> 0 Then Err.Raise vbObjectError + 6105, "FEMProfileCsv", "unbalanced profile stack"
    s = "id,procedure,calls,inclusive_seconds,exclusive_seconds" & vbCrLf
    For i = 1 To 64
        If nodes(i).Calls > 0 Then s = s & CStr(i) & "," & FEMProfileName(i) & "," & CStr(nodes(i).Calls) & "," & Replace$(Format$(nodes(i).Inclusive, "0.000000000"), Application.DecimalSeparator, ".") & "," & Replace$(Format$(nodes(i).Child, "0.000000000"), Application.DecimalSeparator, ".") & vbCrLf
    Next
    FEMProfileCsv = s
End Function

Private Function FEMProfileName(ByVal id As Long) As String
    Select Case id
      ' PROFILE_NAMES
      Case Else: FEMProfileName = "P" & CStr(id)
    End Select
End Function
