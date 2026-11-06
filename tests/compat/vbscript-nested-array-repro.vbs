' VBScript access to a UNO sequence<sequence<struct>> returned through the
' OpenOffice OLE bridge (LineShape.PolyPolygon). Needs a running office.
Option Explicit
On Error Resume Next
Dim m, d, doc, page, shape, pts, outer, inner, p, x, i
Sub Report(label)
    If Err.Number <> 0 Then
        WScript.Echo label & ": ERROR " & Hex(Err.Number) & " " & Err.Description
        Err.Clear
    End If
End Sub
Set m = CreateObject("com.sun.star.ServiceManager")
Set d = m.createInstance("com.sun.star.frame.Desktop")
Dim hidden(0)
Set hidden(0) = m.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
hidden(0).Name = "Hidden": hidden(0).Value = True
Set doc = d.loadComponentFromURL("private:factory/scalc", "_blank", 0, hidden)
Set page = doc.getSheets().getByIndex(0).getDrawPage()
Set shape = doc.createInstance("com.sun.star.drawing.LineShape")
page.add shape
Set p = shape.getPosition(): p.X = 1000: p.Y = 500: shape.setPosition p
Set p = shape.getSize(): p.Width = 2000: p.Height = 3000: shape.setSize p

pts = shape.PolyPolygon: Report "get PolyPolygon"
WScript.Echo "VarType(PolyPolygon) = " & VarType(pts) & " (" & TypeName(pts) & ")"
WScript.Echo "IsArray = " & IsArray(pts) & ", UBound = " & UBound(pts): Report "UBound outer"
outer = pts(0): Report "pts(0)"
WScript.Echo "VarType(pts(0)) = " & VarType(outer) & " (" & TypeName(outer) & ")": Report "VarType inner"
WScript.Echo "UBound(pts(0)) = " & UBound(outer): Report "UBound inner"
Set x = outer(0): Report "Set pts(0)(0)"
WScript.Echo "TypeName(pts(0)(0)) = " & TypeName(x): Report "TypeName point"
WScript.Echo "pts(0)(0).X = " & x.X & ", .Y = " & x.Y: Report "point fields"
' Self-assignment from an element of the same variable, as in shapes-line.vbs.
Dim same
same = shape.PolyPolygon
same = same(0): Report "same = same(0)"
WScript.Echo "after same = same(0): VarType " & VarType(same): Report "VarType same"
WScript.Echo "UBound(same) = " & UBound(same): Report "UBound same"
WScript.Echo "same(0).X = " & same(0).X: Report "same(0).X"
i = 0
For Each inner In pts
    i = i + 1
    WScript.Echo "For Each outer item " & i & ": VarType " & VarType(inner)
    For Each x In inner
        WScript.Echo "  point " & x.X & "," & x.Y
    Next
Next
Report "For Each"
doc.close True
WScript.Echo "DONE"
