Option Explicit
Dim excel, workbook, sheet, block, address, first, last, iteration, count, errorNumber
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "worksheet-range-ownership", label
    WScript.Echo label & ": OK"
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set sheet = excel.ActiveSheet
address = "F2:I5"
Set block = sheet.Range(address)
count = block.Borders.Count
For iteration = 1 To 8
    Check address = "F2:I5" And VarType(address) = 8, "Preserved worksheet address " & iteration
Next
Check block.Count = 16, "Worksheet address extent"
first = "F2"
last = "I5"
Set block = sheet.Range(first, last)
count = block.Borders.Count
For iteration = 1 To 8
    Check first = "F2" And last = "I5", "Preserved worksheet endpoints " & iteration
Next
Check block.Count = 16, "Worksheet endpoint extent"
Set block = Nothing
Set first = sheet.Range("F2")
Set last = sheet.Range("I5")
first.Value = 42
last.Value = 84
Set block = sheet.Range(first, last)
Check block.Count = 16, "Object endpoint extent"
Check first.Value = 42 And last.Value = 84, "Object arguments remain usable"
Set block = sheet.Range("F2", last)
Check block.Count = 16 And last.Value = 84, "String and object endpoints"
Set block = sheet.Range(first, "I5")
Check block.Count = 16 And first.Value = 42, "Object and string endpoints"
For iteration = 1 To 16
    Set block = sheet.Range("F2", "I5")
    Check block.Count = 16, "Repeated owned endpoint cleanup " & iteration
Next
On Error Resume Next
Set block = sheet.Range(42, last)
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
Check errorNumber <> 0 And last.Value = 84, "Reject numeric endpoint"
On Error Resume Next
Set block = sheet.Range(first, excel)
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
Check errorNumber <> 0 And first.Value = 42, "Reject non-Range object endpoint"
Set block = Nothing
Set first = Nothing
Set last = Nothing
Set sheet = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Worksheet Range ownership: PASS"
