Option Explicit
Dim excel, workbook, cell, enabled
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "font-script", label
    WScript.Echo label & ": OK"
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set cell = excel.Cells(1, 1)
cell.Value = "script test"
enabled = True
cell.Font.Subscript = enabled
Check cell.Font.Subscript = True, "Enable subscript by reference"
Check cell.Font.Superscript = False, "Subscript excludes superscript"
cell.Font.Subscript = False
Check cell.Font.Subscript = False, "Disable subscript"
cell.Font.Superscript = -1
Check cell.Font.Superscript = True, "Enable superscript by numeric value"
Check cell.Font.Subscript = False, "Superscript excludes subscript"
cell.Font.Subscript = True
Check cell.Font.Superscript = False, "Switch superscript to subscript"
cell.Font.Superscript = True
Check cell.Font.Subscript = False, "Switch subscript to superscript"
cell.Font.Superscript = False
Check cell.Font.Superscript = False, "Disable superscript"
excel.Cells(1, 1).Value = "A1"
excel.Cells(1, 2).Value = "B1"
excel.Cells(2, 1).Value = "A2"
excel.Cells(2, 2).Value = "B2"
excel.Range("A1:B2").Font.Subscript = True
Check excel.Range("A1:B2").Font.Subscript = True, "Set full multi-cell range"
Check excel.Cells(2, 2).Font.Subscript = True, "Last cell in range changed"
excel.Cells(2, 2).Font.Subscript = False
Check IsNull(excel.Range("A1:B2").Font.Subscript), "Mixed range returns Null"
excel.Range("A1:B2").Font.Subscript = False
Check excel.Range("A1:B2").Font.Subscript = False, "Clear full multi-cell range"
Set cell = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Font script integration: PASS"
