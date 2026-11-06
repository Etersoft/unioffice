Option Explicit
' Range.Hidden, CurrentRegion, NumberFormat, FormulaR1C1 and FormulaLocal
' checked through the backend document, not only through successful calls.
Dim excel, workbook, sheet, manager, desktop, documents, document, oosheet, region, cell
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "range-properties", label
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

' Hidden changes row/column visibility and leaves cell protection alone.
sheet.Columns(2).Hidden = True
Check oosheet.getColumns().getByIndex(1).IsVisible = False, "Columns(2).Hidden hides column B"
Check oosheet.getColumns().getByIndex(2).IsVisible = True, "Neighbour column stays visible"
Check oosheet.getCellByPosition(1, 0).CellProtection.IsHidden = False, "Hidden does not touch cell protection"
Check sheet.Columns(2).Hidden = True, "Hidden column reads back"
Check sheet.Range("B1:B1048576").Hidden = True, "Whole-column range addresses columns"
sheet.Range("B1:B1048576").Hidden = False
Check oosheet.getColumns().getByIndex(1).IsVisible = True, "Whole-column range shows column"
sheet.Rows(3).Hidden = True
Check oosheet.getRows().getByIndex(2).IsVisible = False, "Rows(3).Hidden hides row 3"
Check sheet.Range("A3").EntireRow.Hidden = True, "EntireRow reads hidden row"
Check sheet.Range("C3").Hidden = True, "Cell range reads its hidden row"
Check IsNull(sheet.Range("A2:A3").Hidden), "Partly hidden rows read Null"
sheet.Range("A2:A3").Hidden = False
Check oosheet.getRows().getByIndex(2).IsVisible = True, "Cell range shows its rows"
Check sheet.Rows(3).Hidden = False, "Shown row reads back"

' CurrentRegion expands across adjacent non-empty cells.
sheet.Range("B5").Value = 1
sheet.Range("C5").Value = 2
sheet.Range("C6").Value = 3
sheet.Range("F9").Value = 4
Set region = sheet.Range("B5").CurrentRegion
Check region.Row = 5 And region.Column = 2, "CurrentRegion origin"
Check region.Rows.Count = 2 And region.Columns.Count = 2, "CurrentRegion size"
Set region = sheet.Range("F9").CurrentRegion
Check region.Row = 9 And region.Column = 6 And region.Count = 1, "Isolated CurrentRegion"
Set region = Nothing

' NumberFormat stores an English format code on the backend cell.
Check sheet.Range("B5").NumberFormat = "General", "Default NumberFormat is General"
sheet.Range("B5:C6").NumberFormat = "0.00"
Check sheet.Range("B5").NumberFormat = "0.00", "NumberFormat reads back"
Check oosheet.getCellByPosition(2, 5).getString() = "3.00" Or oosheet.getCellByPosition(2, 5).getString() = "3,00", "NumberFormat applies to the whole range"
sheet.Range("B5").NumberFormat = "General"
Check sheet.Range("B5").NumberFormat = "General", "General restores the standard format"
sheet.Range("D1").NumberFormat = "@"
sheet.Range("D1").Value = "007"
Check sheet.Range("D1").NumberFormat = "@", "Text format reads back"

' FormulaR1C1 compiles relative references for every target cell.
sheet.Range("A10").Value = 10
sheet.Range("A11").Value = 20
sheet.Range("B10:B11").FormulaR1C1 = "=RC[-1]*2"
Check oosheet.getCellByPosition(1, 9).getValue() = 20, "FormulaR1C1 first cell value"
Check oosheet.getCellByPosition(1, 10).getValue() = 40, "FormulaR1C1 shifts for second cell"
Check sheet.Range("B11").FormulaR1C1 = "=RC[-1]*2", "FormulaR1C1 reads back relative"
Check sheet.Range("B11").Formula = "=A11*2", "R1C1 formula reads back in A1"
sheet.Range("C10").FormulaR1C1 = "=SUM(R10C1:R11C1,R[1]C[-2])"
Check sheet.Range("C10").Value = 50, "FormulaR1C1 with absolute area and argument list"
Check sheet.Range("C10").FormulaR1C1 = "=SUM(R10C1:R11C1,R[1]C[-2])", "FormulaR1C1 absolute reads back"
sheet.Range("C11").FormulaR1C1 = "=IF(R10C1>5,""R1C1, text"",0)"
Check oosheet.getCellByPosition(2, 10).getString() = "R1C1, text", "FormulaR1C1 keeps string literals"
sheet.Range("C12").FormulaR1C1 = 15
Check sheet.Range("C12").Value = 15, "FormulaR1C1 accepts constants"
Check sheet.Range("C12").FormulaR1C1 = "15", "Constant FormulaR1C1 reads back"

' FormulaLocal uses the backend's localized formula syntax.
sheet.Range("D10").Formula = "=A10+A11"
Check sheet.Range("D10").FormulaLocal = oosheet.getCellByPosition(3, 9).FormulaLocal, "FormulaLocal reads the backend local formula"
sheet.Range("E10:E11").FormulaLocal = sheet.Range("D10").FormulaLocal
Check oosheet.getCellByPosition(4, 9).getValue() = 30, "FormulaLocal first cell value"
Check oosheet.getCellByPosition(4, 10).getFormula() = "=A11+A12", "FormulaLocal shifts for second cell"
sheet.Range("E12").FormulaLocal = "text"
Check sheet.Range("E12").Value = "text", "FormulaLocal accepts constants"

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
WScript.Echo "Range properties: PASS"
