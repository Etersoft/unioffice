Option Explicit
' Shapes.AddLine draws on the worksheet's own page in point coordinates
' without inserting sheets.
Dim excel, workbook, second, manager, desktop, documents, document, sheets, page, line, points, sheetCount
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
points = page.getByIndex(0).PolyPolygon
points = points(0)
Check Near(points(0).X, 2540) And Near(points(0).Y, 1270), "Begin point in 1/100 mm"
Check Near(points(1).X, 0) And Near(points(1).Y, 3810), "End point keeps direction"
Check line.Parent.Parent Is second, "Shape belongs to the worksheet Shapes"
second.Shapes.AddLine 0, 0, 36, 0
Check page.getCount() = 2, "Second line added"
Check sheets.getCount() = sheetCount, "Repeated Shapes access keeps sheets"

Set line = Nothing
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
