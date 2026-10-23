Option Explicit
Private mBorn() As Boolean
Private mActive() As Boolean, mKnown() As Boolean, mContent() As Double, mVolume() As Double, mReady As Boolean
Public FSBirthWater As Double, FSDeathWater As Double, FSStorageChange As Double, FSSkeletonChange As Double
Public FSBoundaryVolume As Double, FSSourceVolume As Double, FSVolumeClosure As Double, FSVolumeKnown As Boolean

Public Sub FSLedgerReset()
  mReady = False
  FSBirthWater = 0#: FSDeathWater = 0#: FSStorageChange = 0#: FSSkeletonChange = 0#
  FSBoundaryVolume = 0#: FSSourceVolume = 0#: FSVolumeClosure = 0#
End Sub

Public Sub FSLedgerTopology(ByVal e As Long, ByVal active As Boolean, ByVal volume As Double, ByVal porosity As Double, ByVal initial As Boolean)
  If Not mReady Then
    ReDim mBorn(NumberOfElement - 1)
    ReDim mActive(NumberOfElement - 1): ReDim mKnown(NumberOfElement - 1)
    ReDim mContent(NumberOfElement - 1): ReDim mVolume(NumberOfElement - 1)
    mReady = True
  End If
  mBorn(e) = False
  If active And Not mActive(e) Then
    mBorn(e) = Not initial
    If Not initial And porosity <= 0# Then FSCoupledInputError "湿潤BIRTHの持込水量には間隙率nが必要です。材料=" & CStr(Elem(e).MatNo + 1)
    mKnown(e) = (porosity > 0#): mContent(e) = porosity * volume: mVolume(e) = volume
    If Not initial Then FSBirthWater = FSBirthWater + mContent(e)
  ElseIf Not active And mActive(e) Then
    If Not mKnown(e) Then FSCoupledInputError "湿潤DEATHの持出水量には初期状態から間隙率nが必要です。"
    FSDeathWater = FSDeathWater + mContent(e): mContent(e) = 0#
  End If
  mActive(e) = active
End Sub

Public Sub FSLedgerIncrement(ByVal e As Long, ByVal storage As Double, ByVal skeleton As Double)
  mContent(e) = mContent(e) + storage + skeleton
  FSStorageChange = FSStorageChange + storage: FSSkeletonChange = FSSkeletonChange + skeleton
End Sub

Public Sub FSLedgerStepStart()
  FSStorageChange = 0#: FSSkeletonChange = 0#: FSBoundaryVolume = 0#: FSSourceVolume = 0#: FSVolumeClosure = 0#
End Sub

Public Sub FSLedgerClose(ByVal boundary As Double, ByVal source As Double)
  Dim e As Long
  FSBoundaryVolume = boundary: FSSourceVolume = source
  FSVolumeClosure = FSStorageChange + FSSkeletonChange - boundary - source
  FSVolumeKnown = True
  For e = 0 To NumberOfElement - 1
    If mActive(e) And Not mKnown(e) Then FSVolumeKnown = False
  Next e
End Sub

Public Sub FSLedgerRecorded()
  FSBirthWater = 0#: FSDeathWater = 0#
End Sub

Public Function FSLedgerWrite(ByVal folder As String) As Boolean
  Dim f As Integer, e As Long
  On Error GoTo Failed
  f = FreeFile: Open folder & Application.PathSeparator & "fixed_stress_volume.csv" For Output As #f
  Print #f, "element,active,content_m3,known,reference_volume_m3"
  For e = 0 To NumberOfElement - 1
    Print #f, CStr(e + 1) & "," & IIf(mActive(e), "1", "0") & "," & FSInvariant(mContent(e)) & "," & IIf(mKnown(e), "1", "0") & "," & FSInvariant(mVolume(e))
  Next e
  Close #f: FSLedgerWrite = True: Exit Function
Failed:
  On Error Resume Next: Close #f
End Function

Public Sub FSLedgerRead(ByVal folder As String)
  Dim f As Integer, line As String, fields() As String, e As Long
  f = FreeFile: Open folder & Application.PathSeparator & "fixed_stress_volume.csv" For Input As #f
  Line Input #f, line
  If line <> "element,active,content_m3,known,reference_volume_m3" Then Close #f: FSCoupledInputError "流体量履歴の列が不正です。"
  For e = 0 To NumberOfElement - 1
    If EOF(f) Then Close #f: FSCoupledInputError "流体量履歴が不足しています。"
    Line Input #f, line: fields = Split(line, ",")
    If UBound(fields) <> 4 Then Close #f: FSCoupledInputError "流体量履歴が不完全です。"
    If fields(0) <> CStr(e + 1) Or fields(1) <> IIf(mActive(e), "1", "0") Or (fields(3) <> "0" And fields(3) <> "1") Then Close #f: FSCoupledInputError "流体量履歴の接続が一致しません。"
    mContent(e) = FSCheckpointNumber(fields(2)): mVolume(e) = FSCheckpointNumber(fields(4)): mKnown(e) = (fields(3) = "1")
    If mVolume(e) < 0# Then Close #f: FSCoupledInputError "流体量履歴の体積が不正です。"
  Next e
  If Not EOF(f) Then Close #f: FSCoupledInputError "流体量履歴に余分な行があります。"
  Close #f
  FSLedgerClose 0#, 0#
End Sub

Private Function FSInvariant(ByVal value As Double) As String
  FSInvariant = Replace$(Format$(value, "0.0000000000000000E+00"), Application.International(xlDecimalSeparator), ".")
End Function

Public Function FSLedgerBorn(ByVal e As Long) As Boolean
  If mReady Then FSLedgerBorn = mBorn(e)
End Function

Public Function FSLedgerMemory() As Variant
  FSLedgerMemory = Array(mActive, mKnown, mContent, mVolume, mBorn, FSBirthWater, FSDeathWater, FSStorageChange, FSSkeletonChange, FSBoundaryVolume, FSSourceVolume, FSVolumeClosure, FSVolumeKnown)
End Function
Public Sub FSLedgerRestore(ByVal saved As Variant)
  mActive = saved(0): mKnown = saved(1): mContent = saved(2): mVolume = saved(3): mBorn = saved(4)
  FSBirthWater = saved(5): FSDeathWater = saved(6): FSStorageChange = saved(7): FSSkeletonChange = saved(8)
  FSBoundaryVolume = saved(9): FSSourceVolume = saved(10): FSVolumeClosure = saved(11): FSVolumeKnown = saved(12): mReady = True
End Sub
