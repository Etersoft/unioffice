Option Explicit
Dim excel, book, sheet, failures, checks, sheetRows, sheetColumns, selected
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
End Sub
Sub Probe(owner, label)
    Dim rows, columns, part, index, errorNumber
    Set rows = owner.Rows
    Set columns = owner.Columns
    Check rows.Count = sheetRows, label & " Rows.Count"
    Check columns.Count = sheetColumns, label & " Columns.Count"
    index = CDbl(2)
    Set part = rows.Item(index)
    Shape part, 2, 1, 1, sheetColumns, label & " Rows.Item"
    Check index = 2 And VarType(index) = 5, label & " Double preserved"
    Set part = columns.Item(index)
    Shape part, 1, 2, sheetRows, 1, label & " Columns.Item"
    Set part = owner.Rows(index)
    Shape part, 2, 1, 1, sheetColumns, label & " Rows(index)"
    Set part = owner.Columns(index)
    Shape part, 1, 2, sheetRows, 1, label & " Columns(index)"
    Check index = 2 And VarType(index) = 5, label & " direct Double preserved"
    index = "B"
    Set part = Nothing
    On Error Resume Next
    Set part = columns.Item(index)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber = 0 And Not (part Is Nothing), label & " Columns.Item(letter)"
    If Not part Is Nothing Then Shape part, 1, 2, sheetRows, 1, label & " column letter"
    Check index = "B" And VarType(index) = 8, label & " letter preserved"
    Set part = rows.Cells(2, 3)
    Shape part, 2, 3, 1, 1, label & " Rows.Cells"
    Set part = columns.Cells(2, 3)
    Shape part, 2, 3, 1, 1, label & " Columns.Cells"
    index = "aa"
    Set part = owner.Columns(index)
    Shape part, 1, 27, sheetRows, 1, label & " direct lowercase letters"
    Check index = "aa" And VarType(index) = 8, label & " lowercase preserved"
    index = "2:4"
    Set part = owner.Rows(index)
    Shape part, 2, 1, 3, sheetColumns, label & " row slice"
    Check part.Count = 3, label & " row slice Count"
    Check index = "2:4" And VarType(index) = 8, label & " row slice preserved"
    Set part = rows.Item(index)
    Shape part, 2, 1, 3, sheetColumns, label & " Rows.Item slice"
    index = "b:d"
    Set part = owner.Columns(index)
    Shape part, 1, 2, sheetRows, 3, label & " column slice"
    Check part.Count = 3, label & " column slice Count"
    Check index = "b:d" And VarType(index) = 8, label & " column slice preserved"
    Set part = columns.Item(index)
    Shape part, 1, 2, sheetRows, 3, label & " Columns.Item slice"
    Set part = owner.Rows(sheetRows).EntireRow
    Shape part, sheetRows, 1, 1, sheetColumns, label & " last EntireRow"
    Set part = owner.Columns(sheetColumns).EntireColumn
    Shape part, 1, sheetColumns, sheetRows, 1, label & " last EntireColumn"
    For Each index In Array(":2", "2:", "2::3", "3:2", "0:2", "1:2147483648", "A0", "")
        On Error Resume Next
        Set part = owner.Rows(index)
        errorNumber = Err.Number
        Err.Clear
        On Error GoTo 0
        Check errorNumber <> 0, label & " rejects Rows(" & index & ")"
        On Error Resume Next
        Set part = owner.Columns(index)
        errorNumber = Err.Number
        Err.Clear
        On Error GoTo 0
        Check errorNumber <> 0, label & " rejects Columns(" & index & ")"
    Next
    On Error Resume Next
    Set part = owner.Rows(1, 2)
    Check Err.Number <> 0, label & " rejects second Rows index"
    Err.Clear
    Set part = owner.Columns(1, 2)
    Check Err.Number <> 0, label & " rejects second Columns index"
    Err.Clear
    On Error GoTo 0
    Set selected = rows.Item(2)
    Set part = Nothing
    Set rows = Nothing
    Set columns = Nothing
    Check selected.Parent.Name = sheet.Name, label & " selected parent"
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set book = excel.Workbooks.Add()
Set sheet = excel.ActiveSheet
sheetRows = sheet.Cells.Rows.Count
sheetColumns = sheet.Cells.Columns.Count
WScript.Echo "Backend dimensions: " & sheetRows & " rows, " & sheetColumns & " columns"
Probe sheet, "Worksheet"
Probe excel, "Application"
Set sheet = Nothing
Check selected.Parent.Name <> "", "Selected row retains worksheet"
Shape selected, 2, 1, 1, sheetColumns, "Selected row lifetime"
Set selected = Nothing
book.Close False
Set book = Nothing
excel.Quit
Set excel = Nothing
If failures Then
    WScript.Echo "Sheet Rows and Columns: " & failures & " failure(s) in " & checks & " checks"
    WScript.Quit 1
End If
WScript.Echo "Sheet Rows and Columns: PASS (" & checks & " checks)"
