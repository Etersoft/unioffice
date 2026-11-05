Option Explicit
Dim excel, workbook, borders, outer, inner, expected, outerCount, innerCount, total
Dim key, seen(12), aliases, aliasTargets, errorNumber
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "borders-enumeration", label
    WScript.Echo label & ": OK"
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set borders = excel.Range("A1:D4").Borders
expected = borders.Count
Check expected = 8, "Eight distinct borders"
For key = 5 To 12
    borders(key).LineStyle = 1
    borders(key).ColorIndex = key
    seen(key) = 0
Next
For Each outer In borders
    key = outer.ColorIndex
    Check key >= 5 And key <= 12, "Enumerated border color"
    seen(key) = seen(key) + 1
Next
For key = 5 To 12
    Check seen(key) = 1, "Border returned once, key " & key
Next
aliases = Array(1, 2, 3, 4)
aliasTargets = Array(7, 10, 8, 9)
For key = 0 To 3
    Check borders(aliases(key)).ColorIndex = borders(aliasTargets(key)).ColorIndex, "Legacy alias " & aliases(key)
Next
borders.LineStyle = 1
borders.ColorIndex = 10
For key = 7 To 12
    Check borders(key).LineStyle = 1, "Aggregate line style, key " & key
    Check borders(key).ColorIndex = 10, "Aggregate color, key " & key
Next
For key = 5 To 6
    Check borders(key).ColorIndex = key, "Aggregate preserves diagonal, key " & key
Next
For Each key In Array(0, -1, 13, -2147483648, 2147483647)
    Set outer = Nothing
    On Error Resume Next
    Set outer = borders(key)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber <> 0, "Rejected index " & key
Next
outerCount = 0
total = 0
For Each outer In borders
    outerCount = outerCount + 1
    innerCount = 0
    For Each inner In borders
        innerCount = innerCount + 1
        total = total + 1
    Next
    Check innerCount = expected, "Inner enumeration " & outerCount
Next
Check outerCount = expected, "Outer enumeration count"
Check total = expected * expected, "Nested enumeration count"
outerCount = 0
For Each outer In borders
    outerCount = outerCount + 1
Next
Check outerCount = expected, "Repeated enumeration count"
Set inner = Nothing
Set outer = Nothing
Set borders = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Borders enumeration integration: PASS"
