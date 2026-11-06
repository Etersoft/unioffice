Option Explicit
Dim excel, workbook, sheet, block, cell, row, column, value, errorNumber, checks
checks = 0
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "range-item", label
    checks = checks + 1
    WScript.Echo label & ": OK"
End Sub
Sub Target(targetCell, address, number)
    targetCell.Value = number
    Check excel.Range(address).Value = number, "Target " & address
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set sheet = excel.ActiveSheet
Set block = sheet.Range("B2:D4")
row = CDbl(2)
column = "3"
Set cell = block.Item(row, column)
Target cell, "D3", 101
Check VarType(row) = 5 And row = 2, "Item preserves Double row"
Check VarType(column) = 8 And column = "3", "Item preserves string column"
column = "d"
Set cell = block.Item(row, column)
Target cell, "E3", 102
Check column = "d" And VarType(column) = 8, "Item preserves letter column"
Set cell = block.Item(1, "AA")
Target cell, "AB2", 103
Set cell = block.Item(3, 4)
Target cell, "E4", 104
Set cell = block.Item(1)
Target cell, "B2", 105
Set cell = block.Item(3)
Target cell, "D2", 106
Set cell = block.Item(4)
Target cell, "B3", 107
Set cell = block.Item(10)
Target cell, "B5", 108
row = "2"
Set cell = block.Item(row, 1)
Target cell, "B3", 109
Check VarType(row) = 8 And row = "2", "Item preserves string row"
row = CDbl(2)
column = "3"
Set cell = block(row, column)
Target cell, "D3", 110
Set cell = block(4)
Target cell, "B3", 111
Set cell = block(2, "AA")
Target cell, "AB3", 112
Set cell = block.Cells.Item(4)
Target cell, "B3", 113
Set cell = sheet.Cells.Item(4)
Target cell, "D1", 114
Set cell = excel.Cells.Item(4)
Target cell, "D1", 115
value = "owned-value"
block.Item(row, column) = value
Check excel.Range("D3").Value = value, "Item put target"
Check value = "owned-value" And VarType(value) = 8, "Item put preserves value"
Check row = 2 And VarType(row) = 5 And column = "3" And VarType(column) = 8, "Item put preserves indices"
block.Item(10) = 116
Check excel.Range("B5").Value = 116, "Linear Item put target"
block(2, "D") = 117
Check excel.Range("E3").Value = 117, "Default put target"
block(4) = 118
Check excel.Range("B3").Value = 118, "Linear default put target"
block.Cells(3, 3) = 119
Check excel.Range("D4").Value = 119, "Relative Cells put target"
column = "invalid-column"
On Error Resume Next
Set cell = block.Item(row, column)
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
Check errorNumber <> 0 And column = "invalid-column", "Item get conversion error"
On Error Resume Next
block.Item(row, column) = 999
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
Check errorNumber <> 0 And column = "invalid-column", "Item put conversion error"
On Error Resume Next
Set cell = block.Item()
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
Check errorNumber <> 0, "Item rejects missing row"
On Error Resume Next
Set cell = block.Item(1, 2, 3)
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
Check errorNumber <> 0, "Item rejects extra argument"
Check excel.Range("D3").Value = value And excel.Range("B3").Value = 118, "Failed Item preserves values"
block.Borders.Weight = 2
Check block.Borders.Weight = 2, "Borders Weight after Item"
block.Borders(7).Color = 12345
Check block.Borders(7).Color = 12345, "Border Color after Item"
block.Borders(8).LineStyle = 1
Check block.Borders(8).LineStyle = 1, "Border LineStyle after Item"
Set cell = block.Item(2, 2)
Set block = Nothing
Set sheet = Nothing
Target cell, "C3", 120
Check cell.Parent.Name = excel.ActiveSheet.Name, "Item keeps worksheet parent alive"
Set cell = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Range Item: PASS (" & checks & " checks)"
