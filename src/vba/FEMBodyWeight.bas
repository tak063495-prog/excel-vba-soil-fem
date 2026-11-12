Option Explicit

' Optional N-column saturated TOTAL unit weight. No submerged-weight shortcut:
' pore pressure remains a separate effective-stress/external-water load.
Private mSaturated() As Double, mProfiles() As String, mVariable As Boolean

Public Sub FEMWeightPrepare()
  Dim row As Long, mat As Long, ws As Worksheet, values As Variant, mode As String
  mVariable = False
  If NumberOfMaterial < 1 Then Exit Sub
  ReDim mSaturated(0 To NumberOfMaterial - 1)
  ReDim mProfiles(0 To NumberOfMaterial - 1)
  For mat = 0 To NumberOfMaterial - 1
    mSaturated(mat) = Material(mat).weight: mProfiles(mat) = "*"
  Next mat
  For row = 1 To P6InputMaterialCount
    mat = P6InputMaterialId(row) - 1
    mSaturated(mat) = P6InputMaterialWeightSat(row)
    If mSaturated(mat) <> Material(mat).weight Then mVariable = True
  Next row
  If Not mVariable Then Exit Sub
  If Not GWEnabled() Then mVariable = False: Exit Sub
  Set ws = ThisWorkbook.Worksheets("水条件")
  mode = UCase$(Trim$(CStr(ws.range("B3").value2)))
  If mode <> "HYDROSTATIC" Or FSActive Or Application.CountIf(ws.range("M12:M211"), "FIXED_STRESS") > 0 Then
    Err.Raise vbObjectError + 3482, "FEMBodyWeight", "飽和単位体積重量N列は静的HYDROSTATIC・片方向連携のみです。定常浸透/時間連成ではN列を空欄にし、適切な全重量を指定してください。"
  End If
  GWReadProfileCache
  values = ws.range("F12:I211").value2
  For row = 1 To UBound(values, 1)
    If Len(Trim$(CStr(values(row, 1)))) > 0 Then
      If IsNumeric(values(row, 1)) Then
        mat = CLng(values(row, 1)) - 1
        If mat >= 0 And mat < NumberOfMaterial Then
          If Len(Trim$(CStr(values(row, 4)))) > 0 Then mProfiles(mat) = Trim$(CStr(values(row, 4)))
        End If
      End If
    End If
  Next row
  P6WeightNote = "全重量γt/γsatを積分点の水位で選択。水圧は別途適用（水中重量への置換なし）"
End Sub

Public Function FEMWeightAt(ByVal elementId As Long, ByRef shapeN() As Double) As Double
  Dim mat As Long, j As Long, x As Double, y As Double
  mat = Elem(elementId).MatNo
  FEMWeightAt = Material(mat).weight
  If Not mVariable Then Exit Function
  If mSaturated(mat) = FEMWeightAt Then Exit Function
  For j = 0 To 7
    x = x + shapeN(j) * Elem(elementId).x(j)
    y = y + shapeN(j) * Elem(elementId).y(j)
  Next j
  If GWWeightIsWet(mProfiles(mat), x, y) Then FEMWeightAt = mSaturated(mat)
End Function
