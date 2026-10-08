Option Explicit
Private enabled As Boolean, starts As Object, times As Object, counts As Object
Public Sub MeshPerfReset(Optional ByVal active As Boolean = True)
    enabled = active
    Set starts = CreateObject("Scripting.Dictionary")
    Set times = CreateObject("Scripting.Dictionary")
    Set counts = CreateObject("Scripting.Dictionary")
End Sub
Public Sub MeshPerfEnter(ByVal key As String)
    If Not enabled Then Exit Sub
    starts(key) = CDbl(Timer)
End Sub
Public Sub MeshPerfLeave(ByVal key As String)
    Dim elapsed As Double
    If Not enabled Then Exit Sub
    If Not starts.Exists(key) Then Exit Sub
    elapsed = CDbl(Timer) - CDbl(starts(key)): If elapsed < 0# Then elapsed = elapsed + 86400#
    If Not times.Exists(key) Then times.Add key, 0#: counts.Add key, 0&
    times(key) = CDbl(times(key)) + elapsed: counts(key) = CLng(counts(key)) + 1
    starts.REMOVE key
End Sub
Public Function MeshPerfReport() As String
    Dim key As Variant
    If times Is Nothing Then Exit Function
    MeshPerfReport = "工程,回数,秒" & vbCrLf
    For Each key In times.Keys
        MeshPerfReport = MeshPerfReport & key & "," & counts(key) & "," & Format$(times(key), "0.000000") & vbCrLf
    Next key
End Function
