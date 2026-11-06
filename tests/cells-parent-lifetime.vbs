Option Explicit
Dim excel, workbook, sheet, cells, cell, parent, expectedName, iteration
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "cells-parent-lifetime", label
    WScript.Echo label & ": OK"
End Sub
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set workbook = excel.Workbooks.Add()
Set sheet = excel.ActiveSheet
expectedName = sheet.Name
Set sheet = Nothing
Set cells = excel.Cells
Set parent = cells.Parent
Check parent.Name = expectedName, "Application.Cells retains parent"
Set parent = Nothing
For iteration = 1 To 16
    Set sheet = excel.ActiveSheet
    Set cells = sheet.Cells
    Set sheet = Nothing
    Set parent = cells.Parent
    Check parent.Name = expectedName, "Worksheet.Cells retains parent " & iteration
    Set parent = Nothing
    Set cell = cells(2, 3)
    Set cells = Nothing
    cell.Value = iteration
    Check excel.Range("C2").Value = iteration, "Cell survives collection release " & iteration
    Set parent = cell.Parent
    Check parent.Name = expectedName, "Child retains parent " & iteration
    Set parent = Nothing
    Set cell = Nothing
Next
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Cells parent lifetime: PASS"
