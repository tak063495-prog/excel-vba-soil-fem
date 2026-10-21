Option Explicit

Public Function TNativeFigure() As String
    Dim ws As Worksheet, figure As Shape, before As Long, svg As String
    Dim q As Shape, p As Variant, e As Long, k As Variant
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets.Add
    svg = "<svg xmlns='http://www.w3.org/2000/svg'><g font-family='Meiryo, sans-serif'>" & _
        "<polygon points='20,30 60,30 60,70' fill='#123456' stroke='#abcdef'/>" & _
        "<polyline points='70,30 80,40' fill='none' stroke='#654321' stroke-dasharray='3,2'/>" & _
        "<path d='M 20 90 L 60 90' fill='none' stroke='#123456'/>" & _
        "<path d='M 70 90 L 80 100 L 60 100 Z' fill='#ff0000'/>" & _
        "<text x='10' y='20' fill='#263444'>A&amp;B</text>" & _
        "<line x1='90' y1='20' x2='90' y2='70' stroke='#00ff00'/>" & _
        "</g></svg>"
    Set figure = FEMNativeFigureCreate(ws, svg, 160, 140)
    AssertNative figure.Type = msoGroup And figure.GroupItems.Count = 7, "primitive count"
    Set q = figure.GroupItems(2)
    AssertNative q.Nodes.Count = 3, "polygon closure count=" & CStr(q.Nodes.Count)
    p = q.Nodes(1).Points
    AssertNative Abs(p(1, 1) - 30) < 0.001 And Abs(p(1, 2) - 40) < 0.001, "coordinate offset"
    AssertNative q.Fill.ForeColor.RGB = RGB(18, 52, 86), "polygon fill"
    AssertNative figure.GroupItems(3).Line.DashStyle <> msoLineSolid, "inherited/dashed stroke"
    AssertNative figure.GroupItems(4).Nodes.Count = 2, "open path has no extra zero point"
    AssertNative figure.GroupItems(5).Nodes.Count = 3, "closed path count=" & CStr(figure.GroupItems(5).Nodes.Count)
    Set q = figure.GroupItems(6)
    AssertNative q.TextFrame2.TextRange.Text = "A&B" And q.Fill.Visible = msoFalse, "transparent decoded text"
    AssertNative figure.GroupItems(7).Line.ForeColor.RGB = RGB(0, 255, 0), "native line"
    before = ws.Shapes.Count
    For Each k In Array("<circle/>", "<path d='M 10 10 C 20 20 30 30'/>", "<rect width='5' height='5' fill='bad'/>")
        On Error Resume Next
        Set q = FEMNativeFigureCreate(ws, "<svg>" & CStr(k) & "</svg>", 160, 140)
        e = Err.Number: Err.Clear
        On Error GoTo Failed
        AssertNative e <> 0, "unsupported/invalid primitive rejected"
        AssertNative ws.Shapes.Count = before, "failure cleans only new shapes " & CStr(k) & " before=" & CStr(before) & " after=" & CStr(ws.Shapes.Count)
        AssertNative figure.GroupItems.Count = 7, "previous drawing retained"
    Next k
    figure.Delete
    For Each k In Array("zero", "constant", "linear", "axes", "shear", "negative_shear", "thin")
        Set figure = FEMNativeFigureCreate(ws, "<svg>" & TIncrementSvgFixture(CStr(k)) & "</svg>", 160, 140)
        AssertNative figure.Type = msoGroup Or figure.Type = msoAutoShape, "increment fixture " & CStr(k)
        figure.Delete
    Next k
    ws.Delete
    TNativeFigure = "PASS|primitive geometry, colors, dashes, text, failure cleanup, seven increment fixtures"
    Exit Function
Failed:
    TNativeFigure = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Sub AssertNative(ByVal ok As Boolean, ByVal label As String)
    If Not ok Then Err.Raise vbObjectError + 4920, "NativeFigureHarness", label
End Sub
