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
