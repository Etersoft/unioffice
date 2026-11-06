Option Explicit
Dim excel, workbook, sheet, failures, checks, cell
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
Sub Probe(owner, label, firstValue)
    Dim cell, collection, index, row, column, value, errorNumber
    index = CDbl(4)
    Set cell = Nothing
    On Error Resume Next
    Set cell = owner.Cells(index)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber = 0 And Not (cell Is Nothing), label & " single-index get"
    If Not cell Is Nothing Then
        cell.Value = firstValue
        Check sheet.Range("D1").Value = firstValue, label & " single-index target"
        Check cell.Parent.Name = sheet.Name, label & " cell parent"
    End If
    Check index = 4 And VarType(index) = 5, label & " preserves Double index"
    value = label & "-value"
    On Error Resume Next
    owner.Cells(index) = value
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber = 0, label & " single-index put"
    If errorNumber = 0 Then Check sheet.Range("D1").Value = value, label & " single-index put target"
    Check value = label & "-value" And VarType(value) = 8, label & " preserves assigned string"
    Check index = 4 And VarType(index) = 5, label & " put preserves Double index"
    index = "5"
    Set cell = Nothing
    On Error Resume Next
    Set cell = owner.Cells(index)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber = 0 And Not (cell Is Nothing), label & " numeric-string get"
    If Not cell Is Nothing Then
        cell.Value = firstValue + 1
        Check sheet.Range("E1").Value = firstValue + 1, label & " numeric-string target"
    End If
    Check index = "5" And VarType(index) = 8, label & " preserves string index"
    On Error Resume Next
    owner.Cells(index) = firstValue + 2
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber = 0, label & " numeric-string put"
    If errorNumber = 0 Then Check sheet.Range("E1").Value = firstValue + 2, label & " numeric-string put target"
    row = "2"
    column = "AA"
    Set cell = Nothing
    On Error Resume Next
    Set cell = owner.Cells(row, column)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber = 0 And Not (cell Is Nothing), label & " string row and letter column get"
    If Not cell Is Nothing Then
        cell.Value = firstValue + 3
        Check sheet.Range("AA2").Value = firstValue + 3, label & " two-index target"
    End If
    Check row = "2" And column = "AA", label & " preserves two indices"
    On Error Resume Next
    owner.Cells(row, column) = firstValue + 4
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber = 0, label & " two-index put"
    If errorNumber = 0 Then Check sheet.Range("AA2").Value = firstValue + 4, label & " two-index put target"
    index = "invalid-index"
    On Error Resume Next
    Set cell = owner.Cells(index)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber <> 0 And index = "invalid-index", label & " invalid get"
    On Error Resume Next
    owner.Cells(index) = 999
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber <> 0 And index = "invalid-index", label & " invalid put"
    On Error Resume Next
    Set cell = owner.Cells(1, 2, 3)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber <> 0, label & " rejects extra get argument"
    On Error Resume Next
    owner.Cells(1, 2, 3) = 999
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber <> 0, label & " rejects extra put argument"
    Set collection = owner.Cells
    Set cell = collection.Item(6)
    Set collection = Nothing
    cell.Value = firstValue + 5
    Check sheet.Range("F1").Value = firstValue + 5, label & " collection child lifetime"
    Check cell.Parent.Name = sheet.Name, label & " collection child parent"
    Set cell = Nothing
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set sheet = excel.ActiveSheet
Probe excel, "Application.Cells", 100
Probe sheet, "Worksheet.Cells", 200
Set cell = sheet.Cells.Item(1)
cell.Value = 301
Check sheet.Range("A1").Value = 301, "Explicit Item remains usable"
Set cell = Nothing
Set sheet = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
If failures Then
    WScript.Echo "Cells indexing: " & failures & " failure(s) in " & checks & " checks"
    WScript.Quit 1
End If
WScript.Echo "Cells indexing: PASS (" & checks & " checks)"
