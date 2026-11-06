Option Explicit
' Range.Copy, Range.PasteSpecial and Application.CutCopyMode checked
' against the backend cells.
Dim excel, workbook, sheet, manager, desktop, documents, document, oosheet, failed
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "paste-special", label
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

sheet.Range("A1").Value = 2
sheet.Range("A2").Formula = "=A1*10"
sheet.Range("A1:A2").Font.Bold = True
Check excel.CutCopyMode = 0, "CutCopyMode is False initially"
sheet.Range("A1:A2").Copy
Check excel.CutCopyMode = 1, "Copy sets xlCopy"

' Values only: no formula, no formatting.
sheet.Range("C1").PasteSpecial -4163
Check oosheet.getCellByPosition(2, 1).getValue() = 20, "xlPasteValues copies the result"
Check oosheet.getCellByPosition(2, 1).getFormula() = "20", "xlPasteValues drops the formula"
Check oosheet.getCellByPosition(2, 0).CharWeight < 150, "xlPasteValues drops formatting"

' Formats only.
sheet.Range("D1").Value = 7
sheet.Range("D1").PasteSpecial -4122
Check oosheet.getCellByPosition(3, 0).CharWeight >= 150, "xlPasteFormats copies bold"
Check oosheet.getCellByPosition(3, 0).getValue() = 7, "xlPasteFormats keeps the value"

' Everything, with a relative formula.
sheet.Range("E1").PasteSpecial
Check oosheet.getCellByPosition(4, 1).getFormula() = "=E1*10", "xlPasteAll shifts the formula"
Check oosheet.getCellByPosition(4, 0).CharWeight >= 150, "xlPasteAll copies formatting"

' Arithmetic operation and transpose.
sheet.Range("F1").Value = 100
sheet.Range("F2").Value = 100
sheet.Range("F1").PasteSpecial -4163, 2
Check oosheet.getCellByPosition(5, 0).getValue() = 102 And oosheet.getCellByPosition(5, 1).getValue() = 120, "Operation xlAdd"
sheet.Range("G1").PasteSpecial -4163, -4142, False, True
Check oosheet.getCellByPosition(6, 0).getValue() = 2 And oosheet.getCellByPosition(7, 0).getValue() = 20, "Transpose"

' Unsupported paste kinds fail instead of pasting something else.
On Error Resume Next
sheet.Range("J1").PasteSpecial 8
failed = Err.Number <> 0
Err.Clear
On Error GoTo 0
Check failed, "xlPasteColumnWidths is refused"

' CutCopyMode = False ends the copy mode.
excel.CutCopyMode = False
Check excel.CutCopyMode = 0, "CutCopyMode = False"
On Error Resume Next
sheet.Range("K1").PasteSpecial -4163
failed = Err.Number <> 0
Err.Clear
On Error GoTo 0
Check failed And oosheet.getCellByPosition(10, 0).getType() = 0, "PasteSpecial fails after CutCopyMode = False"

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
WScript.Echo "Paste special: PASS"
