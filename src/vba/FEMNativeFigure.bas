Option Explicit

' Render the small SVG dialect emitted by FEMViewer as native Excel shapes.
' This avoids Excel's version-dependent SVG image decoder and does not write a
' temporary file.  The routine deliberately fails on an unsupported element.
Public Function FEMNativeFigureCreate(ByVal ws As Worksheet, ByVal svg As String, _
    ByVal canvasW As Double, ByVal canvasH As Double) As Shape
    Dim doc As Object, root As Object, names As Collection
    Dim shp As Shape, nm As String, i As Long, errNo As Long, errSource As String, errText As String
    If ws Is Nothing Then Err.Raise vbObjectError + 3810, "FEMNativeFigure", "図シートがありません。"
    If canvasW <= 0# Or canvasH <= 0# Then Err.Raise vbObjectError + 3811, "FEMNativeFigure", "図の大きさが不正です。"
    Set names = New Collection
    On Error GoTo Failed
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False
    If Not doc.LoadXML(svg) Then Err.Raise vbObjectError + 3812, "FEMNativeFigure", "SVGの解析に失敗しました。"
    Set root = doc.documentElement
    If root Is Nothing Then Err.Raise vbObjectError + 3812, "FEMNativeFigure", "SVGのルートがありません。"
    If LCase$(FEMNativeLocalName(root)) <> "svg" Then Err.Raise vbObjectError + 3812, "FEMNativeFigure", "SVGのルートが不正です。"
    Set shp = ws.Shapes.AddShape(1, 10#, 10#, CSng(canvasW), CSng(canvasH))
    nm = FEMNativeName(): shp.Name = nm: names.Add nm
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255): shp.Line.Visible = 0
    FEMNativeRenderChildren ws, root, names, 10#, 10#, canvasW, canvasH
    If names.Count = 1 Then Set FEMNativeFigureCreate = shp: Exit Function
    Set FEMNativeFigureCreate = ws.Shapes.Range(FEMNativeNamesArray(names)).Group
    Exit Function
Failed:
    errNo = Err.Number: errSource = Err.Source: errText = Err.Description
    On Error Resume Next
    For i = names.Count To 1 Step -1
        ws.Shapes(CStr(names(i))).Delete
    Next i
    On Error GoTo 0
    Err.Raise errNo, errSource, errText
End Function

Private Sub FEMNativeRenderChildren(ByVal ws As Worksheet, ByVal parent As Object, ByVal names As Collection, ByVal ox As Double, ByVal oy As Double, ByVal canvasW As Double, ByVal canvasH As Double)
    Dim n As Object, tag As String, st As Object
    For Each n In parent.childNodes
        If n.nodeType = 1 Then
            tag = LCase$(FEMNativeLocalName(n))
            Select Case tag
                Case "g": Set st = FEMNativeStyle(n, Nothing): FEMNativeRenderGroup ws, n, names, ox, oy, st, canvasW, canvasH
                Case "rect", "polygon", "polyline", "line", "text", "path"
                    FEMNativeRenderElement ws, n, names, ox, oy, FEMNativeStyle(n, Nothing), canvasW, canvasH
                Case "title", "desc", "metadata": ' ignored metadata
                Case Else: Err.Raise vbObjectError + 3813, "FEMNativeFigure", "未対応のSVG要素: " & tag
            End Select
        End If
    Next n
End Sub

Private Sub FEMNativeRenderGroup(ByVal ws As Worksheet, ByVal parent As Object, ByVal names As Collection, ByVal ox As Double, ByVal oy As Double, ByVal inherited As Object, ByVal canvasW As Double, ByVal canvasH As Double)
    Dim n As Object, tag As String
    For Each n In parent.childNodes
        If n.nodeType = 1 Then
            tag = LCase$(FEMNativeLocalName(n))
            Select Case tag
                Case "g": FEMNativeRenderGroup ws, n, names, ox, oy, FEMNativeStyle(n, inherited), canvasW, canvasH
                Case "rect", "polygon", "polyline", "line", "text", "path"
                    FEMNativeRenderElement ws, n, names, ox, oy, FEMNativeStyle(n, inherited), canvasW, canvasH
                Case "title", "desc", "metadata"
                Case Else: Err.Raise vbObjectError + 3813, "FEMNativeFigure", "未対応のSVG要素: " & tag
            End Select
        End If
    Next n
End Sub

Private Sub FEMNativeRenderElement(ByVal ws As Worksheet, ByVal n As Object, ByVal names As Collection, ByVal ox As Double, ByVal oy As Double, ByVal st As Object, ByVal canvasW As Double, ByVal canvasH As Double)
    Dim tag As String, sh As Shape, nm As String, x As Double, y As Double, x2 As Double, y2 As Double
    Dim pts As Variant, i As Long, txt As String, boxWidth As Double
    Dim errNo As Long, errText As String, phase As String, registered As Boolean
    On Error GoTo Failed
    tag = LCase$(FEMNativeLocalName(n))
    phase = "create"
    Select Case tag
        Case "rect"
            If LCase$(FEMNativeAttr(n, "width")) = "100%" And LCase$(FEMNativeAttr(n, "height")) = "100%" Then Exit Sub
            Set sh = ws.Shapes.AddShape(1, ox + FEMNativeNum(n, "x", 0#), oy + FEMNativeNum(n, "y", 0#), CSng(FEMNativeNum(n, "width", 0#)), CSng(FEMNativeNum(n, "height", 0#)))
        Case "line"
            x = ox + FEMNativeNum(n, "x1", 0#): y = oy + FEMNativeNum(n, "y1", 0#)
            x2 = ox + FEMNativeNum(n, "x2", 0#): y2 = oy + FEMNativeNum(n, "y2", 0#)
            Set sh = ws.Shapes.AddLine(CSng(x), CSng(y), CSng(x2), CSng(y2))
        Case "polygon", "polyline"
            pts = FEMNativePoints(FEMNativeAttr(n, "points"), tag = "polygon", ox, oy)
            Set sh = FEMNativePolyline(ws, pts)
        Case "path"
            pts = FEMNativePathPoints(FEMNativeAttr(n, "d"), ox, oy)
            Set sh = FEMNativePolyline(ws, pts)
        Case "text"
            x = ox + FEMNativeNum(n, "x", 0#): y = oy + FEMNativeNum(n, "y", 0#)
            boxWidth = canvasW - x + 10#: If boxWidth < 1# Then boxWidth = 1#
            Set sh = ws.Shapes.AddTextbox(1, CSng(x), CSng(y - FEMNativeNum(st, "font-size", 12#)), CSng(boxWidth), CSng(FEMNativeNum(st, "font-size", 12#) * 1.25))
            nm = FEMNativeName(): sh.Name = nm: names.Add nm: registered = True
            txt = FEMNativeText(n): sh.TextFrame2.TextRange.Text = txt
            sh.TextFrame2.WordWrap = 0: sh.TextFrame2.AutoSize = 0
            sh.TextFrame2.MarginLeft = 0: sh.TextFrame2.MarginRight = 0: sh.TextFrame2.MarginTop = 0: sh.TextFrame2.MarginBottom = 0
            sh.TextFrame2.TextRange.Font.Name = FEMNativeFirstFont(FEMNativeAttrOr(st, "font-family", "Meiryo"))
            sh.TextFrame2.TextRange.Font.Size = CSng(FEMNativeNum(st, "font-size", 12#))
    End Select
    If tag <> "text" Then
        nm = FEMNativeName(): sh.Name = nm: names.Add nm: registered = True
    End If
    phase = "metadata/style"
    sh.AlternativeText = n.Text
    FEMNativeApplyStyle sh, st, tag = "text"
    Exit Sub
Failed:
    errNo = Err.Number: errText = Err.Description
    On Error Resume Next
    If Not sh Is Nothing Then sh.Delete
    If registered Then names.Remove names.Count
    On Error GoTo 0
    Err.Raise errNo, "FEMNativeFigure", tag & " (" & phase & "): " & errText & " / " & Left$(n.XML, 240)
End Sub

Private Function FEMNativePolyline(ByVal ws As Worksheet, ByRef pts As Variant) As Shape
    Dim builder As FreeformBuilder, i As Long
    Set builder = ws.Shapes.BuildFreeform(msoEditingCorner, CSng(pts(1, 1)), CSng(pts(1, 2)))
    For i = 2 To UBound(pts, 1)
        builder.AddNodes msoSegmentLine, msoEditingCorner, CSng(pts(i, 1)), CSng(pts(i, 2))
    Next i
    Set FEMNativePolyline = builder.ConvertToShape
End Function

Private Sub FEMNativeApplyStyle(ByVal sh As Shape, ByVal st As Object, ByVal isText As Boolean)
    Dim c As Long, sw As Double, fill As String, stroke As String
    fill = LCase$(FEMNativeAttrOr(st, "fill", "black")): stroke = LCase$(FEMNativeAttrOr(st, "stroke", "none"))
    If isText Then
        c = FEMNativeColor(fill): sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = c
        sh.Fill.Visible = 0: sh.Line.Visible = 0
    Else
        If sh.Type <> msoLine Then
            If fill = "none" Then
                sh.Fill.Visible = msoFalse
            Else
                sh.Fill.ForeColor.RGB = FEMNativeColor(fill): sh.Fill.Visible = msoTrue
            End If
        End If
        If stroke = "none" Then sh.Line.Visible = 0 Else sh.Line.ForeColor.RGB = FEMNativeColor(stroke): sh.Line.Visible = -1
        sw = FEMNativeNum(st, "stroke-width", 1#): If sw > 0# Then sh.Line.Weight = CSng(sw)
        If Len(Trim$(FEMNativeAttrOr(st, "stroke-dasharray", ""))) > 0 And LCase$(Trim$(FEMNativeAttrOr(st, "stroke-dasharray", ""))) <> "none" Then sh.Line.DashStyle = 4
    End If
End Sub

Private Function FEMNativeStyle(ByVal n As Object, ByVal parent As Object) As Object
    Dim d As Object, k As Variant, v As String, p As Object
    Set d = CreateObject("Scripting.Dictionary")
    If Not parent Is Nothing Then
        For Each k In parent.Keys
            d(k) = parent(k)
        Next k
    End If
    For Each k In Array("fill", "stroke", "stroke-width", "stroke-dasharray", "font-family", "font-size")
        v = FEMNativeAttr(n, CStr(k)): If Len(v) > 0 Then d(CStr(k)) = v
    Next k
    v = FEMNativeAttr(n, "style")
    If Len(v) > 0 Then FEMNativeParseStyle d, v
    Set FEMNativeStyle = d
End Function

Private Sub FEMNativeParseStyle(ByVal d As Object, ByVal s As String)
    Dim a As Variant, q As Variant, i As Long
    For Each a In Split(s, ";")
        q = Split(CStr(a), ":"): If UBound(q) >= 1 Then d(LCase$(Trim$(q(0)))) = Trim$(q(1))
    Next a
End Sub

Private Function FEMNativePoints(ByVal s As String, ByVal closed As Boolean, ByVal ox As Double, ByVal oy As Double) As Variant
    Dim q As Variant, a As Variant, out() As Single, n As Long, i As Long, x As String
    s = Replace$(Replace$(s, ",", " "), vbTab, " "): q = Split(Application.Trim(s), " ")
    ReDim out(1 To (UBound(q) + 1) \ 2 + IIf(closed, 1, 0), 1 To 2)
    For i = 0 To UBound(q) Step 2: out(n + 1, 1) = CSng(ox + Val(q(i))): out(n + 1, 2) = CSng(oy + Val(q(i + 1))): n = n + 1: Next i
    If closed Then out(n + 1, 1) = out(1, 1): out(n + 1, 2) = out(1, 2)
    FEMNativePoints = out
End Function

Private Function FEMNativePathPoints(ByVal s As String, ByVal ox As Double, ByVal oy As Double) As Variant
    Dim t As Variant, out() As Single, n As Long, i As Long, closed As Boolean
    Dim numericToken As Object, last As Long
    s = Replace$(Replace$(Replace$(s, "M", " M "), "L", " L "), "Z", " Z ")
    t = Split(Application.Trim(s), " "): last = UBound(t)
    If last < 5 Then GoTo InvalidPath
    If t(0) <> "M" Then GoTo InvalidPath
    closed = (t(last) = "Z")
    If closed Then last = last - 1
    If (last + 1) Mod 3 <> 0 Then GoTo InvalidPath
    Set numericToken = CreateObject("VBScript.RegExp")
    numericToken.Pattern = "^[+-]?([0-9]+([.][0-9]*)?|[.][0-9]+)([eE][+-]?[0-9]+)?$"
    ReDim out(1 To (last + 1) \ 3 + IIf(closed, 1, 0), 1 To 2)
    For i = 0 To last Step 3
        If i > 0 And t(i) <> "L" Then GoTo InvalidPath
        If Not numericToken.Test(t(i + 1)) Or Not numericToken.Test(t(i + 2)) Then GoTo InvalidPath
        n = n + 1
        out(n, 1) = CSng(ox + Val(t(i + 1))): out(n, 2) = CSng(oy + Val(t(i + 2)))
    Next i
    If closed Then out(n + 1, 1) = out(1, 1): out(n + 1, 2) = out(1, 2)
    FEMNativePathPoints = out
    Exit Function
InvalidPath:
    Err.Raise vbObjectError + 3814, "FEMNativeFigure", "pathはM/L座標と末尾Zだけに対応しています。"
End Function

Private Function FEMNativeNamesArray(ByVal c As Collection) As Variant
    Dim a() As Variant, i As Long: ReDim a(1 To c.Count)
    For i = 1 To c.Count: a(i) = CStr(c(i)): Next i: FEMNativeNamesArray = a
End Function
Private Function FEMNativeName() As String
    Static n As Long
    n = n + 1
    FEMNativeName = "FEMNative_" & CStr(n) & "_" & CStr(CLng(Timer * 100#))
End Function
Private Function FEMNativeLocalName(ByVal n As Object) As String
    FEMNativeLocalName = n.baseName
    If Len(FEMNativeLocalName) = 0 Then FEMNativeLocalName = n.nodeName
End Function
Private Function FEMNativeAttr(ByVal n As Object, ByVal k As String) As String
    On Error Resume Next
    FEMNativeAttr = CStr(n.getAttribute(k))
    On Error GoTo 0
End Function
Private Function FEMNativeAttrOr(ByVal d As Object, ByVal k As String, ByVal fallback As String) As String
    If d.Exists(k) Then FEMNativeAttrOr = CStr(d(k)) Else FEMNativeAttrOr = fallback
End Function
Private Function FEMNativeNum(ByVal n As Variant, ByVal k As String, ByVal fallback As Double) As Double
    Dim s As String
    If IsObject(n) Then
        If TypeName(n) = "Dictionary" Then
            If n.Exists(k) Then s = CStr(n(k))
        Else
            s = FEMNativeAttr(n, k)
        End If
    End If
    If Len(s) = 0 Then FEMNativeNum = fallback Else FEMNativeNum = Val(s)
End Function
Private Function FEMNativeText(ByVal n As Object) As String
    Dim q As Object
    For Each q In n.childNodes
        If q.nodeType = 3 Then FEMNativeText = FEMNativeText & q.nodeValue
    Next q
End Function
Private Function FEMNativeFirstFont(ByVal s As String) As String
    Dim a As Variant
    a = Split(s, ",")
    FEMNativeFirstFont = Trim$(CStr(a(0)))
    If Len(FEMNativeFirstFont) = 0 Then FEMNativeFirstFont = "Meiryo"
End Function
Private Function FEMNativeColor(ByVal s As String) As Long
    s = LCase$(Trim$(s))
    If s = "white" Then
        FEMNativeColor = RGB(255, 255, 255)
    ElseIf s = "black" Or s = "none" Or Len(s) = 0 Then
        FEMNativeColor = RGB(0, 0, 0)
    ElseIf Left$(s, 1) = "#" And Len(s) >= 7 Then
        FEMNativeColor = RGB(Val("&H" & Mid$(s, 2, 2)), Val("&H" & Mid$(s, 4, 2)), Val("&H" & Mid$(s, 6, 2)))
    Else
        Err.Raise vbObjectError + 3815, "FEMNativeFigure", "色指定が不正です: " & s
    End If
End Function
