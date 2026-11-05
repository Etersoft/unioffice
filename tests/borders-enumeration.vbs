Option Explicit
Dim excel, workbook, borders, outer, inner, expected, outerCount, innerCount, total
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
