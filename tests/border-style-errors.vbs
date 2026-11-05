Option Explicit
Dim excel, workbook, borders, border, key, style, errorNumber
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "border-style-errors", label
    WScript.Echo label & ": OK"
End Sub
Sub PutStyle(target, propertyName, value)
    On Error Resume Next
    If propertyName = "LineStyle" Then
        target.LineStyle = value
    Else
        target.Value = value
    End If
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set borders = excel.Range("A1:D4").Borders
For key = 1 To 12
    borders(key).LineStyle = 1
    borders(key).Weight = 2
    borders(key).ColorIndex = 10
Next
For Each style In Array(-4115, 4, 5, -4118, 13, 12345)
    For key = 1 To 12
        Set border = borders(key)
        PutStyle border, "LineStyle", style
        Check errorNumber <> 0, "Rejected individual style " & style & ", key " & key
        Check border.LineStyle = 1 And border.Weight = 2 And border.ColorIndex = 10, "Unchanged individual, key " & key
    Next
    PutStyle borders, "LineStyle", style
    Check errorNumber <> 0, "Rejected aggregate LineStyle " & style
    PutStyle borders, "Value", style
    Check errorNumber <> 0, "Rejected aggregate Value " & style
    For key = 5 To 12
        Check borders(key).LineStyle = 1 And borders(key).Weight = 2 And borders(key).ColorIndex = 10, "Unchanged aggregate member, key " & key
    Next
Next
borders.Value = -4119
Check borders.Value = -4119 And borders.LineStyle = -4119, "Supported double alias"
For key = 7 To 12
    Check borders(key).LineStyle = -4119 And borders(key).Weight = 2, "Double alias member, key " & key
Next
Check borders(5).LineStyle = 1 And borders(6).LineStyle = 1, "Alias preserves diagonals"
borders.LineStyle = 1
Check borders.Value = 1, "Supported continuous alias"
Set border = Nothing
Set borders = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Border style errors: PASS"
