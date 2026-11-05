Option Explicit
' Isolate Range/Borders argument lifetime without invoking Weight.
Dim excel, workbook, addresses, address, expected, target, iteration, count
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
addresses = Array("F2", "F2:I2", "F2:F5", "F2:I5")
For Each address In addresses
    expected = CStr(address)
    Set target = excel.Range(address).Borders
    count = target.Count
    For iteration = 1 To 4
        WScript.Echo "Range argument: " & address
        If address <> expected Then Err.Raise vbObjectError + 1, "range-address-variable", "Changed Range argument"
    Next
    Set target = Nothing
Next
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Range address lifetime: PASS"
