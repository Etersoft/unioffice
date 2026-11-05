Option Explicit
Dim excel, workbook, manager, desktop, documents, document, sheet, target
Dim border, borders, line, key, weight, previous(12), actual, errorNumber
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "border-weight", label
    WScript.Echo label & ": OK"
End Sub
Function UnoLine(index)
    Dim table
    Select Case index
        Case 5: Set UnoLine = target.getPropertyValue("DiagonalTLBR")
        Case 6: Set UnoLine = target.getPropertyValue("DiagonalBLTR")
        Case Else
            Set table = target.getPropertyValue("TableBorder")
            Select Case index
                Case 1, 7: Set UnoLine = table.LeftLine
                Case 2, 10: Set UnoLine = table.RightLine
                Case 3, 8: Set UnoLine = table.TopLine
                Case 4, 9: Set UnoLine = table.BottomLine
                Case 11: Set UnoLine = table.VerticalLine
                Case 12: Set UnoLine = table.HorizontalLine
            End Select
    End Select
End Function
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set manager = CreateObject("com.sun.star.ServiceManager")
Set desktop = manager.createInstance("com.sun.star.frame.Desktop")
Set documents = desktop.getComponents().createEnumeration()
Set document = documents.nextElement()
Set sheet = document.Sheets.getByIndex(0)
Set target = sheet.getCellRangeByName("A1:D4")
Set borders = excel.Range("A1:D4").Borders
For Each weight In Array(1, 2, -4138, 4)
    For key = 1 To 12
        Set border = borders(key)
        border.LineStyle = 1
        border.ColorIndex = 10
        border.Weight = weight
        Check border.Weight = weight, "Single Weight " & weight & ", key " & key
        Check border.LineStyle = 1 And border.ColorIndex = 10, "Single style/color, key " & key
        Set line = UnoLine(key)
        actual = line.InnerLineWidth + line.OuterLineWidth + line.LineDistance
        Check line.OuterLineWidth > 0 And line.InnerLineWidth = 0 And line.LineDistance = 0, "UNO single geometry, key " & key
        Check actual > previous(key), "Increasing UNO width, key " & key
        previous(key) = actual
    Next
Next
For key = 5 To 12
    Set border = borders(key)
    For Each weight In Array(1, 2, -4138, 4)
        border.LineStyle = -4119
        border.Weight = weight
        Check border.Weight = weight And border.LineStyle = -4119, "Double Weight " & weight & ", key " & key
        Set line = UnoLine(key)
        Check line.InnerLineWidth > 0 And line.OuterLineWidth > 0 And line.LineDistance > 0, "UNO double geometry, key " & key
        border.LineStyle = 1
        Check border.Weight = weight And border.LineStyle = 1, "Weight retained by LineStyle, key " & key
    Next
    border.LineStyle = -4142
    Check border.LineStyle = -4142, "Clear line, key " & key
    border.Weight = 2
    Check border.Weight = 2 And border.LineStyle = 1, "Weight makes line visible, key " & key
Next
borders(5).Weight = 1
borders(6).Weight = 4
For Each weight In Array(1, 2, -4138, 4)
    borders.Weight = weight
    Check borders.Weight = weight, "Aggregate Weight " & weight
    For key = 7 To 12
        Check borders(key).Weight = weight, "Aggregate member, key " & key
        Set line = UnoLine(key)
        Check line.OuterLineWidth > 0, "Aggregate UNO width, key " & key
    Next
    Check borders(5).Weight = 1 And borders(6).Weight = 4, "Aggregate preserves diagonals"
Next
For Each weight In Array(0, 3, -1, 2147483647)
    On Error Resume Next
    borders(7).Weight = weight
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber <> 0, "Rejected individual Weight " & weight
    On Error Resume Next
    borders.Weight = weight
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber <> 0, "Rejected aggregate Weight " & weight
    Check borders(7).Weight = 4, "Invalid Weight preserves previous value"
Next
' A single edge update must preserve the other cells' differing edge widths.
Set border = excel.Range("A1").Borders(7)
border.Weight = 1
Set border = excel.Range("A2").Borders(7)
border.Weight = 4
excel.Range("A1:D4").Borders(8).Weight = 2
Check excel.Range("A1").Borders(7).Weight = 1, "Preserve mixed left edge A1"
Check excel.Range("A2").Borders(7).Weight = 4, "Preserve mixed left edge A2"
Set line = Nothing
Set border = Nothing
Set borders = Nothing
Set target = Nothing
Set sheet = Nothing
Set document = Nothing
Set documents = Nothing
Set desktop = Nothing
Set manager = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Border weight integration: PASS"
