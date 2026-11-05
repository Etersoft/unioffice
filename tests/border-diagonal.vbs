Option Explicit
Dim excel, workbook, manager, desktop, documents, document, sheet, target, line, border, key, name
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "border-diagonal", label
    WScript.Echo label & ": OK"
End Sub
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
For key = 1 To 12
    Set border = excel.Range("A1:D4").Borders(key)
    border.LineStyle = 1
    border.ColorIndex = 10
    Check border.LineStyle = 1, "LineStyle round trip, key " & key
    Check border.ColorIndex = 10, "ColorIndex round trip, key " & key
    If key = 5 Or key = 6 Then
        If key = 5 Then name = "DiagonalTLBR" Else name = "DiagonalBLTR"
        Set line = target.getPropertyValue(name)
        Check line.Color = 32768, "UNO diagonal color, key " & key
        Check line.InnerLineWidth + line.OuterLineWidth > 0, "UNO diagonal width, key " & key
    End If
Next
For key = 5 To 6
    Set border = excel.Range("A1:D4").Borders(key)
    border.LineStyle = -4119
    Check border.LineStyle = -4119, "Double diagonal, key " & key
    border.LineStyle = -4142
    Check border.LineStyle = -4142, "Clear diagonal, key " & key
Next
Set line = Nothing
Set border = Nothing
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
WScript.Echo "Border integration: PASS"
