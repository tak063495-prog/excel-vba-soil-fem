Option Explicit
Private accepted As Collection
Private restored As Variant

Public Function WF50State() As Variant
  Dim selector As Variant, oldFixed() As Boolean, fixed() As Boolean, pressure() As Double, force() As Double
  Dim heads() As Double, qx() As Double, qy() As Double, reaction() As Double
  Dim flux() As Double, e As Long, g As Long, j As Long
  selector = GWTransientState(): oldFixed = selector(6)
  ' GWTransientBoundary only mutates mHeadFixed; restore it immediately so
  ' observing a dedicated drain set never changes the analysis state.
  GWTransientBoundary fixed, pressure
  GWTransientFixed oldFixed
  GWTransientForce force
  FSViewHydraulics heads, qx, qy, reaction
  ReDim flux(0 To NumberOfElement * 8 - 1)
  For e = 0 To NumberOfElement - 1
    For g = 0 To 3
      flux(j) = qx(g, e): flux(j + 1) = qy(g, e): j = j + 2
    Next g
  Next e
  WF50State = Array(FSRunningStage, FSStageElapsed, FSStageDuration, FSTime, FSAccepted, GWViewDefaultProfile(), _
      selector(1), selector(2), selector(3), force, fixed, pressure, heads, flux, reaction, WFPressure())
End Function

Public Sub WF50Accepted()
  If accepted Is Nothing Then Set accepted = New Collection
  accepted.Add WF50State()
End Sub
Public Sub WF50Restored()
  restored = WF50State()
End Sub
Public Function WF50Frames() As Variant
  Dim values() As Variant, i As Long
  If accepted Is Nothing Then Err.Raise 5, , "No accepted water observations"
  ReDim values(0 To accepted.Count - 1)
  For i = 1 To accepted.Count: values(i - 1) = accepted(i): Next i
  WF50Frames = values
End Function
Public Function WF50Loaded() As Variant
  If IsEmpty(restored) Then Err.Raise 5, , "No restored water observation"
  WF50Loaded = restored
End Function
