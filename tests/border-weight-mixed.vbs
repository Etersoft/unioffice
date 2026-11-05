Option Explicit
Dim excel, workbook, borders, key, weight, address, expectedAddress
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "border-weight-mixed", label
    WScript.Echo label & ": OK"
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set borders = excel.Range("A1:D4").Borders
For Each weight In Array(1, 2, -4138, 4)
    borders.Weight = weight
    Check VarType(borders.Weight) = 3 And borders.Weight = weight, "Uniform Long Weight " & weight
Next
borders.Weight = 2
borders(5).Weight = 1
borders(6).Weight = 4
Check borders.Weight = 2, "Ignore diagonal differences"
For key = 7 To 12
    borders(key).Weight = 4
    Check IsNull(borders.Weight), "Mixed member key " & key
    borders(key).Weight = 2
    Check borders.Weight = 2, "Uniform again key " & key
Next
excel.Range("A1").Borders(7).Weight = 1
excel.Range("A2").Borders(7).Weight = 4
Check IsNull(borders.Weight), "Mixed cells on left edge"
borders.Weight = 2
excel.Range("A1").Borders(7).ColorIndex = 5
excel.Range("A2").Borders(7).ColorIndex = 6
Check borders.Weight = 2, "Mixed colors with same Weight"
excel.Range("A1").Borders(7).LineStyle = -4119
Check borders.Weight = 2, "Mixed styles with same Weight"
borders.Weight = 2
excel.Range("B2").Borders(7).Weight = 4
Check IsNull(borders.Weight), "Mixed inside vertical segments"
borders.Weight = 2
excel.Range("B2").Borders(8).Weight = 4
Check IsNull(borders.Weight), "Mixed inside horizontal segments"
borders.Weight = 2
Check borders.Weight = 2, "Restore uniform after cell mixtures"
For Each address In Array("F2", "F2:I2", "F2:F5", "F2:I5")
    expectedAddress = Left(address, Len(address))
    Set borders = excel.Range(address).Borders
    For Each weight In Array(1, 2, -4138, 4)
        borders.Weight = weight
        Check VarType(borders.Weight) = 3 And borders.Weight = weight, "Non-origin " & address & ", Weight " & weight
        Check address = expectedAddress, "Range address argument preserved"
    Next
Next
Set borders = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Mixed border weight integration: PASS"
