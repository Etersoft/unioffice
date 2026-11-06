Option Explicit
' UsedRange, PageSetup print ranges, Font.OutlineFont, Workbook.Saved and
' Application.CentimetersToPoints checked against the backend document.
Dim excel, workbook, sheet, manager, desktop, documents, document, oosheet, used, setup, areas, titles
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "sheet-properties", label
    WScript.Echo label & ": OK"
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set sheet = workbook.Worksheets(1)
Set manager = CreateObject("com.sun.star.ServiceManager")
Set desktop = manager.createInstance("com.sun.star.frame.Desktop")
Set documents = desktop.getComponents().createEnumeration()
Set document = documents.nextElement()
Set oosheet = document.getSheets().getByIndex(0)

Check Abs(excel.CentimetersToPoints(2.54) - 72) < 0.000001, "CentimetersToPoints(2.54)"
Check Abs(excel.CentimetersToPoints(1) - 28.3464566929) < 0.000001, "CentimetersToPoints(1)"

' Workbook.Saved mirrors the backend modified flag in both directions.
workbook.Saved = True
Check document.isModified() = False, "Saved=True clears modified"
Check workbook.Saved = True, "Saved reads True"
sheet.Range("C3").Value = 1
Check workbook.Saved = False, "Edit makes Saved False"
workbook.Saved = True
Check document.isModified() = False, "Saved=True after edit"
workbook.Saved = False
Check document.isModified() = True, "Saved=False marks modified"

' UsedRange starts at the first used cell and reaches the last one.
sheet.Range("E7").Value = 2
sheet.Range("D10").Value = 3
Set used = sheet.UsedRange
Check used.Row = 3 And used.Column = 3, "UsedRange origin"
Check used.Rows.Count = 8 And used.Columns.Count = 3, "UsedRange.Rows/Columns.Count"
Set used = Nothing

' OutlineFont is stored as contoured characters.
Check sheet.Range("C3").Font.OutlineFont = False, "OutlineFont default"
sheet.Range("C3:D4").Font.OutlineFont = True
Check oosheet.getCellByPosition(3, 3).CharContoured = True, "OutlineFont applies to range"
Check sheet.Range("C3").Font.OutlineFont = True, "OutlineFont reads back"
sheet.Range("C3:D4").Font.OutlineFont = False
Check oosheet.getCellByPosition(2, 2).CharContoured = False, "OutlineFont clears"

' PrintArea and PrintTitleRows reach the sheet print settings.
Set setup = sheet.PageSetup
Check setup.PrintArea = "", "No PrintArea by default"
setup.PrintArea = "B2:D5"
areas = oosheet.getPrintAreas()
Check UBound(areas) = 0, "PrintArea sets one area"
Check areas(0).StartColumn = 1 And areas(0).EndRow = 4, "PrintArea bounds"
Check setup.PrintArea = "$B$2:$D$5", "PrintArea reads back"
setup.PrintArea = "A1"
Check setup.PrintArea = "$A$1", "Single-cell PrintArea"
setup.PrintArea = ""
Check UBound(oosheet.getPrintAreas()) = -1, "Empty PrintArea clears areas"
Check setup.PrintTitleRows = "", "No title rows by default"
setup.PrintTitleRows = "$1:$2"
Set titles = oosheet.getTitleRows()
Check oosheet.getPrintTitleRows() = True, "PrintTitleRows enables titles"
Check titles.StartRow = 0 And titles.EndRow = 1, "PrintTitleRows bounds"
Check setup.PrintTitleRows = "$1:$2", "PrintTitleRows reads back"
setup.PrintTitleRows = "A3"
Check setup.PrintTitleRows = "$3:$3", "Cell reference selects its row"
setup.PrintTitleRows = ""
Check oosheet.getPrintTitleRows() = False, "Empty PrintTitleRows disables titles"
Set titles = Nothing
Set setup = Nothing

' Full-width ranges group rows using the backend's real sheet width.
Dim lastColumn
lastColumn = oosheet.getColumns().getCount()
sheet.Range(sheet.Cells(12, 1), sheet.Cells(13, lastColumn)).Group
sheet.Outline.ShowLevels 1
Check oosheet.getRows().getByIndex(11).IsVisible = False, "Full-width Group groups rows"
Check oosheet.getColumns().getByIndex(1).IsVisible = True, "Full-width Group leaves columns"
sheet.Outline.ShowLevels 2
sheet.Range(sheet.Cells(12, 1), sheet.Cells(13, lastColumn)).Ungroup
sheet.Outline.ShowLevels 1
Check oosheet.getRows().getByIndex(11).IsVisible = True, "Full-width Ungroup removes row group"

Set oosheet = Nothing
Set document = Nothing
Set documents = Nothing
Set desktop = Nothing
Set manager = Nothing
Set sheet = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Sheet properties: PASS"
