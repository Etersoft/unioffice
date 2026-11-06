Option Explicit
Dim excel, workbook, manager, desktop, documents, document, sheet, rows, columns, outline
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "outline-levels", label
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
Set sheet = document.getSheets().getByIndex(0)
sheet.group sheet.getCellRangeByName("B2:D4").getRangeAddress(), 1
sheet.group sheet.getCellRangeByName("B2:D4").getRangeAddress(), 0
Set rows = sheet.getRows()
Set columns = sheet.getColumns()
Set outline = excel.Sheets(1).Outline
Check outline.ShowLevels(1, 1) = True, "ShowLevels returns an Automation result"
Check rows.getByIndex(1).IsVisible = False, "Collapse rows"
Check columns.getByIndex(1).IsVisible = False, "Collapse columns in same call"
outline.ShowLevels 2, 0
Check rows.getByIndex(1).IsVisible = True, "Expand rows"
Check columns.getByIndex(1).IsVisible = False, "Zero leaves columns unchanged"
outline.ShowLevels , 2
Check columns.getByIndex(1).IsVisible = True, "Omitted rows allow column-only update"
outline.ShowLevels 0, 1
Check rows.getByIndex(1).IsVisible = True, "Zero leaves rows unchanged"
Check columns.getByIndex(1).IsVisible = False, "Collapse columns only"
' SummaryRow/SummaryColumn default to Excel's Below/Right and persist per sheet.
Check outline.SummaryRow = 1, "SummaryRow defaults to xlSummaryBelow"
Check outline.SummaryColumn = -4152, "SummaryColumn defaults to xlSummaryOnRight"
outline.SummaryRow = 0
outline.SummaryColumn = -4131
Check excel.Sheets(1).Outline.SummaryRow = 0, "SummaryRow persists for the sheet"
Check excel.Sheets(1).Outline.SummaryColumn = -4131, "SummaryColumn persists for the sheet"
Check document.getDocumentProperties().getUserDefinedProperties().getPropertyValue("UniOffice.Outline.SummaryRow." & sheet.getName()) = 0, "SummaryRow is stored in the document"
Dim failed
On Error Resume Next
outline.SummaryRow = 5
failed = Err.Number <> 0
Err.Clear
On Error GoTo 0
Check failed, "Invalid SummaryRow is rejected"
outline.SummaryRow = 1
Check excel.Sheets(1).Outline.SummaryRow = 1, "SummaryRow returns to Below"
Set outline = Nothing
Set rows = Nothing
Set columns = Nothing
Set sheet = Nothing
Set document = Nothing
Set documents = Nothing
Set desktop = Nothing
Set manager = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Outline integration: PASS"
