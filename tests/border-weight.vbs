Option Explicit
' Known failing probe: Weight is still a stub; do not include in passing suites.
Dim excel, workbook, border
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set border = excel.Range("A1:D4").Borders(7)
border.LineStyle = 1
border.Weight = 2
If border.Weight <> 2 Then Err.Raise vbObjectError + 1, "border-weight", "Weight round trip"
Set border = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Border weight integration: PASS"
