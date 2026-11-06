Option Explicit
Dim excel, workbook, sheet, block, cell, row, column, iteration, count, errorNumber
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "cell-index-ownership", label
    WScript.Echo label & ": OK"
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set sheet = excel.ActiveSheet
row = 2
column = "3"
Set cell = excel.Cells(row, column)
count = cell.Borders.Count
For iteration = 1 To 8
    Check column = "3" And VarType(column) = 8, "Preserved Application.Cells column " & iteration
Next
cell.Value = 42
Check excel.Range("C2").Value = 42, "Application.Cells target"
column = "4"
Set cell = sheet.Cells(row, column)
count = cell.Borders.Count
For iteration = 1 To 8
    Check column = "4" And VarType(column) = 8, "Preserved Worksheet.Cells column " & iteration
Next
cell.Value = 84
Check excel.Range("D2").Value = 84, "Worksheet.Cells target"
Set block = excel.Range("F2:I5")
row = CDbl(2)
column = "3"
Set cell = block.Cells(row, column)
count = cell.Borders.Count
For iteration = 1 To 8
    Check column = "3" And VarType(column) = 8 And VarType(row) = 5, "Preserved relative indices " & iteration
Next
cell.Value = 126
Check excel.Range("H3").Value = 126, "Relative Cells target"
row = 2
column = "5"
excel.Cells(row, column) = 168
Check column = "5" And VarType(column) = 8, "Preserved Application.Cells put index"
Check excel.Range("E2").Value = 168, "Application.Cells put target"
column = "6"
sheet.Cells(row, column) = 210
Check column = "6" And VarType(column) = 8, "Preserved Worksheet.Cells put index"
Check excel.Range("F2").Value = 210, "Worksheet.Cells put target"
column = "invalid-column"
On Error Resume Next
Set cell = excel.Cells(row, column)
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
Check errorNumber <> 0 And column = "invalid-column", "Application.Cells get conversion error"
On Error Resume Next
excel.Cells(row, column) = 999
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
Check errorNumber <> 0 And column = "invalid-column", "Application.Cells put conversion error"
On Error Resume Next
Set cell = sheet.Cells(row, column)
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
Check errorNumber <> 0 And column = "invalid-column", "Worksheet.Cells get conversion error"
On Error Resume Next
sheet.Cells(row, column) = 999
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
Check errorNumber <> 0 And column = "invalid-column", "Worksheet.Cells put conversion error"
On Error Resume Next
Set cell = block.Cells(row, column)
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
Check errorNumber <> 0 And column = "invalid-column", "Relative Cells conversion error"
Check excel.Range("E2").Value = 168 And excel.Range("F2").Value = 210, "Invalid indices preserve cell values"
Set cell = Nothing
Set block = Nothing
Set sheet = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Cell index ownership: PASS"
