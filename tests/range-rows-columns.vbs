Option Explicit
Dim excel, book, sheet, source, rows, columns, part, cells, failures, checks, index
failures = 0
checks = 0
Sub Check(condition, label)
    checks = checks + 1
    If condition Then
        WScript.Echo label & ": OK"
    Else
        failures = failures + 1
        WScript.Echo label & ": FAILED"
    End If
End Sub
Sub Shape(target, firstRow, firstColumn, rowCount, columnCount, label)
    Check target.Row = firstRow, label & " row"
    Check target.Column = firstColumn, label & " column"
    Check target.Rows.Count = rowCount, label & " height"
    Check target.Columns.Count = columnCount, label & " width"
    Check target.Cells.Count = rowCount * columnCount, label & " cells"
End Sub
Sub SelectItem(collection, itemIndex, firstRow, firstColumn, rowCount, columnCount, label)
    Dim selected, errorNumber
    Set selected = Nothing
    On Error Resume Next
    Set selected = collection.Item(itemIndex)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber = 0 And Not (selected Is Nothing), label & " get"
    If Not selected Is Nothing Then Shape selected, firstRow, firstColumn, rowCount, columnCount, label
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set book = excel.Workbooks.Add()
Set sheet = excel.ActiveSheet
Set source = sheet.Range("B2:D4")
Set rows = source.Rows
Set columns = source.Columns
Check rows.Count = 3, "Rows.Count"
Check columns.Count = 3, "Columns.Count"
Shape rows, 2, 2, 3, 3, "Rows bounds"
Shape columns, 2, 2, 3, 3, "Columns bounds"
SelectItem rows, 2, 3, 2, 1, 3, "Rows.Item(2)"
SelectItem columns, 2, 2, 3, 3, 1, "Columns.Item(2)"
SelectItem rows, 5, 6, 2, 1, 3, "Rows outside original bounds"
SelectItem columns, "AA", 2, 28, 3, 1, "Columns letter offset"
index = "2"
SelectItem rows, index, 3, 2, 1, 3, "Rows string index"
Check index = "2" And VarType(index) = 8, "Preserves string index"
On Error Resume Next
Set part = source.Rows(2)
Check Err.Number = 0, "Direct Rows(index)"
Err.Clear
On Error GoTo 0
If Not part Is Nothing Then Shape part, 3, 2, 1, 3, "Direct row"
Set part = Nothing
On Error Resume Next
Set part = source.Columns("B")
Check Err.Number = 0, "Direct Columns(letter)"
Err.Clear
On Error GoTo 0
If Not part Is Nothing Then Shape part, 2, 3, 3, 1, "Direct column"
Set part = Nothing
On Error Resume Next
Set part = columns("B")
Check Err.Number = 0, "Columns default letter"
Err.Clear
On Error GoTo 0
If Not part Is Nothing Then Shape part, 2, 3, 3, 1, "Default column"
Set part = Nothing
On Error Resume Next
Set part = rows.Item(1, 1)
Check Err.Number <> 0, "Rows rejects second Item argument"
Err.Clear
Set part = columns.Item(1, 1)
Check Err.Number <> 0, "Columns rejects second Item argument"
Err.Clear
On Error GoTo 0
SelectItem rows, "2:3", 3, 2, 2, 3, "Relative row slice"
SelectItem columns, "b:d", 2, 3, 3, 3, "Relative column slice"
Set part = rows.Cells(4)
Shape part, 3, 2, 1, 1, "Rows.Cells(4)"
Set part = columns.Cells(2, 2)
Shape part, 3, 3, 1, 1, "Columns.Cells(2,2)"
Set cells = columns.Cells
Set part = cells.Item(4)
Shape part, 3, 2, 1, 1, "Cells resets collection kind"
Check columns.Count = 3, "Cells leaves original collection unchanged"
Check source.Count = 9, "Rows and Columns leave original range unchanged"
Set part = rows.Item(2)
Set rows = Nothing
Set columns = Nothing
Set source = Nothing
Set cells = Nothing
Set sheet = Nothing
Check part.Parent.Name <> "", "Selected row retains worksheet parent"
Shape part, 3, 2, 1, 3, "Selected row lifetime"
Set part = Nothing
book.Close False
Set book = Nothing
excel.Quit
Set excel = Nothing
If failures Then
    WScript.Echo "Range Rows and Columns: " & failures & " failure(s) in " & checks & " checks"
    WScript.Quit 1
End If
WScript.Echo "Range Rows and Columns: PASS (" & checks & " checks)"
