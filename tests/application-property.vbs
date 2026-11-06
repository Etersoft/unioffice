Option Explicit
' Every object model level returns the same Application object.
Dim excel, workbook, sheet, range, objects, names, i, ok
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "application-property", label
    WScript.Echo label & ": OK"
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set sheet = workbook.Worksheets(1)
Set range = sheet.Range("A1:B2")
workbook.Names.Add "Probe", "=" & sheet.Name & "!$A$1"
names = Array("Application", "Workbooks", "Workbook", "Sheets", "Worksheets", "Worksheet", _
    "Range", "Cells", "Rows", "Font", "Interior", "Borders", "Border", "PageSetup", _
    "Outline", "Names", "Name", "Shapes", "Shape")
objects = Array(excel, excel.Workbooks, workbook, workbook.Sheets, workbook.Worksheets, sheet, _
    range, sheet.Cells, range.Rows, range.Font, range.Interior, range.Borders, range.Borders(7), sheet.PageSetup, _
    sheet.Outline, workbook.Names, workbook.Names(1), sheet.Shapes, sheet.Shapes.AddLine(10, 10, 50, 50))
For i = 0 To UBound(objects)
    ok = objects(i).Application Is excel
    If Not ok Then ok = objects(i).Application.Workbooks.Count = excel.Workbooks.Count And objects(i).Application.SheetsInNewWorkbook = excel.SheetsInNewWorkbook
    Check ok, names(i) & ".Application"
Next
Check sheet.Shapes.Parent Is sheet, "Shapes.Parent"
Check workbook.Names(1).Name = "Probe", "Names(1) uses one-based indexes"
Set range = Nothing
Set sheet = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Application property: PASS"
