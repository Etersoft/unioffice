Option Explicit
Dim excel, workbook, cell
Set excel = CreateObject("Excel.Application")
WScript.Echo "Excel.Application: OK"
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
WScript.Echo "Workbook: OK"
Set cell = excel.Cells(1, 1)
cell.Value = "UniOffice integration test"
If cell.Value <> "UniOffice integration test" Then
    WScript.Echo "Excel cell round trip: FAILED"
    WScript.Quit 1
End If
WScript.Echo "Excel cell round trip: OK"
Set cell = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "UniOffice integration: PASS"
