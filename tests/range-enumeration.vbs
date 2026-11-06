Option Explicit
Dim excel, book, sheet, source, item, outer, inner, retained, axis, index, count, nested, checks, failures
checks = 0
failures = 0
Sub Check(condition, label)
    checks = checks + 1
    If Not condition Then
        failures = failures + 1
        WScript.Echo label & ": FAILED"
    End If
End Sub
Sub Shape(target, firstRow, firstColumn, height, width, label)
    Check target.Row = firstRow, label & " row"
    Check target.Column = firstColumn, label & " column"
    Check target.Rows.Count = height, label & " height"
    Check target.Columns.Count = width, label & " width"
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set book = excel.Workbooks.Add()
Set sheet = excel.ActiveSheet
Set source = sheet.Range("B2:D3")
index = 0
For Each item In source
    Shape item, 2 + (index \ 3), 2 + (index Mod 3), 1, 1, "Range cell " & index
    item.Value = index + 10
    index = index + 1
Next
Check index = 6, "Range enumeration count"
Check sheet.Range("B2").Value = 10, "First cell write"
Check sheet.Range("D3").Value = 15, "Last cell write"
index = 0
For Each item In source.Rows
    Shape item, 2 + index, 2, 1, 3, "Row " & index
    Check item.Count = 1, "Yielded row retains collection kind"
    index = index + 1
Next
Check index = 2, "Rows enumeration count"
index = 0
For Each item In source.Columns
    Shape item, 2, 2 + index, 2, 1, "Column " & index
    Check item.Count = 1, "Yielded column retains collection kind"
    index = index + 1
Next
Check index = 3, "Columns enumeration count"
index = 0
For Each item In source.Columns.Cells
    Shape item, 2 + (index \ 3), 2 + (index Mod 3), 1, 1, "Columns.Cells " & index
    Check item.Value = index + 10, "Cell enumeration reads values"
    index = index + 1
Next
Check index = 6, "Cells conversion count"
nested = 0
For Each outer In source
    count = 0
    For Each inner In source
        Check inner.Row = 2 + (count \ 3) And inner.Column = 2 + (count Mod 3), "Nested cursor order"
        count = count + 1
        nested = nested + 1
    Next
    Check count = 6, "Independent nested cursor"
Next
Check nested = 36, "Nested enumeration count"
For Each item In source
    Exit For
Next
index = 0
For Each item In source
    index = index + 1
Next
Check index = 6, "Early exit leaves fresh enumeration at start"
For Each axis In Array("Rows", "Columns")
    If axis = "Rows" Then
        Set retained = sheet.Rows
    Else
        Set retained = sheet.Columns
    End If
    index = 0
    For Each item In retained
        If axis = "Rows" Then
            Shape item, index + 1, 1, 1, sheet.Cells.Columns.Count, "Whole worksheet row"
        Else
            Shape item, 1, index + 1, sheet.Cells.Rows.Count, 1, "Whole worksheet column"
        End If
        index = index + 1
        If index = 2 Then Exit For
    Next
    Check index = 2, "Lazy worksheet " & axis
Next
index = 0
For Each item In excel.Cells
    Shape item, 1, index + 1, 1, 1, "Application.Cells lazy item"
    index = index + 1
    If index = 2 Then Exit For
Next
Check index = 2, "Lazy Application.Cells"
index = 0
For Each item In sheet.Range("D3")
    Shape item, 3, 4, 1, 1, "Singleton"
    index = index + 1
Next
Check index = 1, "Singleton count"
For Each item In source.Rows
    Set retained = item
Next
Set source = Nothing
Set sheet = Nothing
Set item = Nothing
Set outer = Nothing
Set inner = Nothing
Shape retained, 3, 2, 1, 3, "Retained yielded row"
Check retained.Parent.Name <> "", "Yielded row retains worksheet"
Set retained = Nothing
book.Close False
excel.Quit
Set book = Nothing
Set excel = Nothing
If failures Then
    WScript.Echo "Range enumeration: " & failures & " failure(s) in " & checks & " checks"
    WScript.Quit 1
End If
WScript.Echo "Range enumeration: PASS (" & checks & " checks)"
