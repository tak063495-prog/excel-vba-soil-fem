Option Explicit

' Input validation for the five user-facing settings pages.
' Unknown keys are deliberately left untouched so older/newer workbooks remain compatible.

Private Function FEMUiValidationPages() As Variant
  FEMUiValidationPages = Array("メッシュ設定", "解析設定", "高速化設定", "出力・表示", "詳細設定")
End Function

Private Function FEMUiValidationList(ByVal keyName As String) As String
  Select Case UCase$(Trim$(keyName))
    Case "MESH_BC_BOTTOM", "MESH_BC_LEFT", "MESH_BC_RIGHT", "MESH_BC_TOP": FEMUiValidationList = "NONE,ROLLER,PINNED"
    Case "MESH_BC_PIN_CORNER": FEMUiValidationList = "AUTO,NONE"
    Case "WEIGHT_MODE": FEMUiValidationList = "TOTAL"
    Case "RCM_POLICY": FEMUiValidationList = "AUTO,ON,OFF"
    Case "FLOW_POLICY": FEMUiValidationList = "INCONSISTENT,DAVIS"
    Case "SRM_PSI_POLICY": FEMUiValidationList = "CAP,REDUCE,KEEP"
    Case "SRM_MODE": FEMUiValidationList = "REAPPLY"
    Case "OUTPUT_STAGE_MODE": FEMUiValidationList = "FINAL,ALL"
    Case "OUTPUT_INACTIVE_ELEMENTS": FEMUiValidationList = "SKIP,INCLUDE"
    Case "DEBUG_MODE": FEMUiValidationList = "OFF,STAGE,ITER"
    Case "SOLVER": FEMUiValidationList = "BAND_LU,BAND_LDLT_EXPERIMENTAL"
    Case "BAND_LU_KERNEL": FEMUiValidationList = "NEW,OLD"
    Case "PHYSICAL_FAILURE_MODE": FEMUiValidationList = "SHADOW"
    Case "EXPORT_MODE": FEMUiValidationList = "OFF,STAGE,ON_FAIL"
    Case "VIEW_SCALE_MODE": FEMUiValidationList = "AUTO,SYMMETRIC,USER"
    Case "VIEW_DETAIL_MODE": FEMUiValidationList = "SIMPLE,DETAIL"
    Case "VIEW_RESULT_SOURCE": FEMUiValidationList = "SMOOTHED,ELEMENT,RAW_GAUSS"
    Case "MESH_GEN_COMPARE": FEMUiValidationList = "0,1,2"
    Case "NATIVE_DLL": FEMUiValidationList = "AUTO,BAND,OFF"
    Case "NATIVE_PRECISION": FEMUiValidationList = "STRICT,SSE2,AUTO,AVX2"
    Case "ACCEL_ADAPTIVE", "ACCEL_V1_PREDICTOR", "ACCEL_V2A_REUSE", "ACCEL_V2B_COST", "ACCEL_V3_COST_SEARCH", "ACCEL_V4_ANDERSON", "ACCEL_V5_GMRES_LU", "ACCEL_TRACE", "ACCEL_STEP_RECOVERY", "ACCEL_STEP_FRESH_LU", "NATIVE_STATE_COPY", "NATIVE_ASSEMBLY", "NATIVE_GMRES_SIMD", "NATIVE_GMRES_TUNED", "NATIVE_GMRES_PREPARED", "IO_FAST_PATH", "EVENT_LOG_FAST", "MESH_REPAIR_TERRAIN_COMPARE": FEMUiValidationList = "0,1"
  End Select
End Function

Private Function FEMUiValidationRule(ByVal keyName As String, ByVal rowNo As Long) As String
  Dim c As String: c = "C" & CStr(rowNo)
  Select Case UCase$(Trim$(keyName))
    Case "MESH_NX", "MESH_NY", "MESH_MATERIAL", "SOLVER_MAX_ITERATIONS", "SOLVER_GMRES_RESTART": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & "=INT(" & c & ")," & c & ">=1," & c & "<=2147483647)"
    Case "MESH_REPAIR_MAX_ITERATIONS": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & "=INT(" & c & ")," & c & ">=1," & c & "<=50)"
    Case "MESH_REPAIR_ANGLE_ITERATIONS": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & "=INT(" & c & ")," & c & ">=1," & c & "<=5000)"
    Case "MESH_POINT1_X", "MESH_POINT1_Y", "MESH_POINT2_X", "MESH_POINT2_Y", "MESH_POINT3_X", "MESH_POINT3_Y", "MESH_POINT4_X", "MESH_POINT4_Y", "VIEW_MIN_X", "VIEW_MAX_Y": FEMUiValidationRule = "=ISNUMBER(" & c & ")"
    Case "Q8_HOURGLASS_FACTOR", "ACCEL_ADAPTIVE", "ACCEL_V1_PREDICTOR", "ACCEL_V2A_REUSE", "ACCEL_V2B_COST", "ACCEL_V3_COST_SEARCH", "ACCEL_V4_ANDERSON", "ACCEL_V5_GMRES_LU", "ACCEL_TRACE", "ACCEL_STEP_RECOVERY", "ACCEL_STEP_FRESH_LU", "NATIVE_STATE_COPY", "NATIVE_ASSEMBLY", "NATIVE_GMRES_SIMD", "NATIVE_GMRES_TUNED", "NATIVE_GMRES_PREPARED", "IO_FAST_PATH", "EVENT_LOG_FAST": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">=0," & c & "<=1)"
    Case "SRM_FMAX", "SRM_TOL", "VIEW_SCALE", "MESH_QUALITY_MIN_AREA_RATIO", "MESH_QUALITY_MIN_JACOBIAN_RATIO", "MESH_QUALITY_MAX_ASPECT_RATIO", "ADAPT_MAX_ELEMENT_RATIO": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">0)"
    Case "DEBUG_FLUSH", "VIEW_DISP_SCALE": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">=0)"
    Case "ACCEL_PREDICT_BETA": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">=0," & c & "<=1)"
    Case "ACCEL_PREDICT_MAX_RATIO": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">0," & c & "<=2)"
    Case "ACCEL_TANGENT_CHANGE": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">0," & c & "<=0.1)"
    Case "SRM_FIXED_FS": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">=0)"
    Case "MESH_QUALITY_MIN_ANGLE_DEG": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">0," & c & "<180)"
    Case "MESH_QUALITY_MAX_ANGLE_DEG": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">0," & c & "<180)"
    Case "MESH_REPAIR_DAMPING", "ADAPT_MARK_FRACTION": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">=0," & c & "<=1)"
    Case "MESH_REPAIR_IMPROVEMENT_TOLERANCE", "MESH_TERRAIN_COUNT_MIN", "MESH_TERRAIN_COUNT_MAX": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">=0)"
    Case "SOLVER_TOLERANCE", "SOLVER_REEVAL_DOF_THRESHOLD", "SOLVER_MEMORY_LIMIT_MB": FEMUiValidationRule = "=AND(ISNUMBER(" & c & ")," & c & ">0)"
  End Select
End Function

Private Sub FEMUiValidationApply(ByVal cell As range, ByVal keyName As String)
  Dim listText As String, ruleText As String
  listText = FEMUiValidationList(keyName)
  ruleText = FEMUiValidationRule(keyName, cell.row)
  If Len(listText) = 0 And Len(ruleText) = 0 Then Exit Sub
  On Error Resume Next
  cell.Validation.Delete
  On Error GoTo 0
  If Len(listText) > 0 Then
    With cell.Validation
      .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listText
      .IgnoreBlank = (UCase$(keyName) = "EXPORT_LOAD")
      .InCellDropdown = True: .ShowError = True: .ErrorTitle = "入力値"
      .errorMessage = "選択肢から入力してください。"
    End With
  ElseIf Len(ruleText) > 0 Then
    With cell.Validation
      .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:=ruleText
      .IgnoreBlank = False: .ShowError = True: .ErrorTitle = "入力値"
      .errorMessage = "指定された範囲・形式の数値を入力してください。"
    End With
  End If
End Sub

Public Sub FEMUiApplyInputValidation()
  Dim page As Variant, ws As Worksheet, lastRow As Long, r As Long, keyName As String
  For Each page In FEMUiValidationPages()
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(CStr(page)): On Error GoTo 0
    If Not ws Is Nothing Then
      lastRow = ws.cells(ws.rows.count, 6).End(xlUp).row
      For r = 1 To lastRow
        keyName = UCase$(Trim$(CStr(ws.cells(r, 6).value2)))
        If Len(keyName) > 0 Then FEMUiValidationApply ws.cells(r, 3), keyName
      Next r
    End If
  Next page
End Sub

Private Function FEMUiValidationListHas(ByVal listText As String, ByVal valueText As String) As Boolean
  Dim item As Variant
  For Each item In Split(listText, ",")
    If StrComp(Trim$(CStr(item)), Trim$(valueText), vbTextCompare) = 0 Then FEMUiValidationListHas = True: Exit Function
  Next item
End Function

Private Function FEMUiValidationValueBlank(ByVal value As Variant) As Boolean
  If IsError(value) Then Exit Function
  If IsEmpty(value) Then FEMUiValidationValueBlank = True: Exit Function
  If VarType(value) = vbString Then FEMUiValidationValueBlank = (Len(Trim$(CStr(value))) = 0)
End Function

Private Function FEMUiValidationAccepts(ByVal keyName As String, ByVal value As Variant) As Boolean
  Dim x As Double, k As String
  If IsError(value) Then Exit Function
  If FEMUiValidationValueBlank(value) Then Exit Function
  Select Case VarType(value)
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
    Case Else: Exit Function
  End Select
  x = CDbl(value): k = UCase$(Trim$(keyName))
  Select Case k
    Case "MESH_NX", "MESH_NY", "MESH_MATERIAL", "SOLVER_MAX_ITERATIONS", "SOLVER_GMRES_RESTART": FEMUiValidationAccepts = (x >= 1# And x <= 2147483647# And x = Fix(x))
    Case "MESH_REPAIR_MAX_ITERATIONS": FEMUiValidationAccepts = (x >= 1# And x <= 50# And x = Fix(x))
    Case "MESH_REPAIR_ANGLE_ITERATIONS": FEMUiValidationAccepts = (x >= 1# And x <= 5000# And x = Fix(x))
    Case "Q8_HOURGLASS_FACTOR", "ACCEL_ADAPTIVE", "ACCEL_V1_PREDICTOR", "ACCEL_V2A_REUSE", "ACCEL_V2B_COST", "ACCEL_V3_COST_SEARCH", "ACCEL_V4_ANDERSON", "ACCEL_V5_GMRES_LU", "ACCEL_TRACE", "ACCEL_STEP_RECOVERY", "ACCEL_STEP_FRESH_LU", "NATIVE_STATE_COPY", "NATIVE_ASSEMBLY", "NATIVE_GMRES_SIMD", "NATIVE_GMRES_TUNED", "NATIVE_GMRES_PREPARED", "IO_FAST_PATH", "EVENT_LOG_FAST": FEMUiValidationAccepts = (x >= 0# And x <= 1#)
    Case "MESH_POINT1_X", "MESH_POINT1_Y", "MESH_POINT2_X", "MESH_POINT2_Y", "MESH_POINT3_X", "MESH_POINT3_Y", "MESH_POINT4_X", "MESH_POINT4_Y", "VIEW_MIN_X", "VIEW_MAX_Y": FEMUiValidationAccepts = True
    Case "SRM_FMAX", "SRM_TOL", "VIEW_SCALE", "MESH_QUALITY_MIN_AREA_RATIO", "MESH_QUALITY_MIN_JACOBIAN_RATIO", "MESH_QUALITY_MAX_ASPECT_RATIO", "ADAPT_MAX_ELEMENT_RATIO": FEMUiValidationAccepts = (x > 0#)
    Case "DEBUG_FLUSH": FEMUiValidationAccepts = (x >= 0#)
    Case "VIEW_DISP_SCALE": FEMUiValidationAccepts = (x >= 0#)
    Case "ACCEL_PREDICT_BETA": FEMUiValidationAccepts = (x >= 0# And x <= 1#)
    Case "ACCEL_PREDICT_MAX_RATIO": FEMUiValidationAccepts = (x > 0# And x <= 2#)
    Case "ACCEL_TANGENT_CHANGE": FEMUiValidationAccepts = (x > 0# And x <= 0.1)
    Case "SRM_FIXED_FS": FEMUiValidationAccepts = (x >= 0#)
    Case "MESH_QUALITY_MIN_ANGLE_DEG", "MESH_QUALITY_MAX_ANGLE_DEG": FEMUiValidationAccepts = (x > 0# And x < 180#)
    Case "MESH_REPAIR_DAMPING", "ADAPT_MARK_FRACTION": FEMUiValidationAccepts = (x >= 0# And x <= 1#)
    Case "MESH_REPAIR_IMPROVEMENT_TOLERANCE", "MESH_TERRAIN_COUNT_MIN", "MESH_TERRAIN_COUNT_MAX": FEMUiValidationAccepts = (x >= 0#)
    Case "SOLVER_TOLERANCE", "SOLVER_REEVAL_DOF_THRESHOLD", "SOLVER_MEMORY_LIMIT_MB": FEMUiValidationAccepts = (x > 0#)
  End Select
End Function

Public Sub FEMUiValidateSettings()
  Dim page As Variant, ws As Worksheet, lastRow As Long, r As Long, keyName As String, value As Variant, listText As String, ruleText As String, values As Object
  Set values = CreateObject("Scripting.Dictionary"): values.CompareMode = vbTextCompare
  For Each page In FEMUiValidationPages()
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(CStr(page)): On Error GoTo 0
    If Not ws Is Nothing Then
      ' Evaluate pasted/edited formulas even when Excel is in manual calculation.
      ws.Calculate
      lastRow = ws.cells(ws.rows.count, 6).End(xlUp).row
      For r = 1 To lastRow
        keyName = UCase$(Trim$(CStr(ws.cells(r, 6).value2)))
        If Len(keyName) > 0 Then
          If FEMUiRetiredMeshKey(keyName) Then GoTo NextCell
          value = ws.cells(r, 3).value2: listText = FEMUiValidationList(keyName): ruleText = FEMUiValidationRule(keyName, r)
          If IsError(value) Then Err.Raise vbObjectError + 3820, "FEMUiValidateSettings", CStr(page) & "!C" & CStr(r) & " (" & keyName & ") の値がエラーです。"
          If keyName = "VIEW_COLOR_RESULT" Then
            If Not FEMUiValidResultId(value) Then Err.Raise vbObjectError + 3821, "FEMUiValidateSettings", CStr(page) & "!C" & CStr(r) & " (" & keyName & ") は表示量一覧の内容名を選択してください。"
          End If
          If Not IsError(value) Then values(keyName) = value
          If keyName = "EXPORT_LOAD" And FEMUiValidationValueBlank(value) Then GoTo NextCell
          If Len(listText) > 0 Then
            If IsError(value) Then Err.Raise vbObjectError + 3820, "FEMUiValidateSettings", CStr(page) & "!C" & CStr(r) & " (" & keyName & ") の値がエラーです。"
            If Not FEMUiValidationListHas(listText, CStr(value)) Then Err.Raise vbObjectError + 3820, "FEMUiValidateSettings", CStr(page) & "!C" & CStr(r) & " (" & keyName & ") の値が不正です。"
          ElseIf Len(ruleText) > 0 Then
            If Not FEMUiValidationAccepts(keyName, value) Then Err.Raise vbObjectError + 3821, "FEMUiValidateSettings", CStr(page) & "!C" & CStr(r) & " (" & keyName & ") の値が範囲外または形式不正です。"
          End If
        End If
NextCell:
      Next r
    End If
  Next page
  If values.Exists("MESH_QUALITY_MIN_ANGLE_DEG") And values.Exists("MESH_QUALITY_MAX_ANGLE_DEG") Then
    If FEMUiValidationAccepts("MESH_QUALITY_MIN_ANGLE_DEG", values("MESH_QUALITY_MIN_ANGLE_DEG")) And FEMUiValidationAccepts("MESH_QUALITY_MAX_ANGLE_DEG", values("MESH_QUALITY_MAX_ANGLE_DEG")) Then
      If CDbl(values("MESH_QUALITY_MIN_ANGLE_DEG")) >= CDbl(values("MESH_QUALITY_MAX_ANGLE_DEG")) Then Err.Raise vbObjectError + 3822, "FEMUiValidateSettings", "詳細設定!C（MESH_QUALITY_MIN_ANGLE_DEG/MAX_ANGLE_DEG）の大小関係が不正です。"
    End If
  End If
  If values.Exists("MESH_TERRAIN_COUNT_MIN") And values.Exists("MESH_TERRAIN_COUNT_MAX") Then
    If CDbl(values("MESH_TERRAIN_COUNT_MIN")) > CDbl(values("MESH_TERRAIN_COUNT_MAX")) Then Err.Raise vbObjectError + 3823, "FEMUiValidateSettings", "詳細設定!C（MESH_TERRAIN_COUNT_MIN/MAX）の大小関係が不正です。"
  End If
End Sub

Private Function FEMUiValidResultId(ByVal value As Variant) As Boolean
  FEMUiValidResultId = (FEMResultChoiceId(value) > 0)
End Function
