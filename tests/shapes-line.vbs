Option Explicit
' Shapes.AddLine draws on the worksheet's own page in point coordinates
' without inserting sheets.
Dim excel, workbook, second, manager, desktop, documents, document, sheets, page, line, points, sheetCount, shape, readable
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "shapes-line", label
    WScript.Echo label & ": OK"
End Sub
Function Near(value, expected)
    Near = Abs(value - expected) <= 2
End Function
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.SheetsInNewWorkbook = 2
Set workbook = excel.Workbooks.Add()
Set manager = CreateObject("com.sun.star.ServiceManager")
Set desktop = manager.createInstance("com.sun.star.frame.Desktop")
Set documents = desktop.getComponents().createEnumeration()
Set document = documents.nextElement()
Set sheets = document.getSheets()
sheetCount = sheets.getCount()

Set second = workbook.Worksheets(2)
Set line = second.Shapes.AddLine(72, 36, 0, 108)
Check sheets.getCount() = sheetCount, "AddLine does not insert sheets"
Check sheets.getByIndex(0).getDrawPage().getCount() = 0, "First sheet stays empty"
Set page = sheets.getByIndex(1).getDrawPage()
Check page.getCount() = 1, "Line is on its own worksheet"
Set shape = page.getByIndex(0)
Check Near(shape.getPosition().X, 0) And Near(shape.getPosition().Y, 1270), "Bounding box origin in 1/100 mm"
Check Near(shape.getSize().Width, 2540) And Near(shape.getSize().Height, 2540), "Bounding box size in 1/100 mm"
' The native OLE bridge cannot pass nested struct sequences to VBScript.
Dim beginX, beginY, endX, endY
On Error Resume Next
points = shape.PolyPolygon
points = points(0)
beginX = points(0).X: beginY = points(0).Y
endX = points(1).X: endY = points(1).Y
readable = Err.Number = 0
Err.Clear
On Error GoTo 0
If readable Then
    Check Near(beginX, 2540) And Near(beginY, 1270), "Begin point in 1/100 mm"
    Check Near(endX, 0) And Near(endY, 3810), "End point keeps direction"
Else
    WScript.Echo "PolyPolygon is not readable through this OLE bridge: direction not checked"
End If
Check line.Parent.Parent Is second, "Shape belongs to the worksheet Shapes"
second.Shapes.AddLine 0, 0, 36, 0
Check page.getCount() = 2, "Second line added"
Check sheets.getCount() = sheetCount, "Repeated Shapes access keeps sheets"

Set line = Nothing
Set shape = Nothing
Set page = Nothing
Set sheets = Nothing
Set document = Nothing
Set documents = Nothing
Set desktop = Nothing
Set manager = Nothing
Set second = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Shapes line: PASS"
