Option Explicit
' Names.Add and Name properties checked against UNO named ranges.
Dim excel, workbook, sheet, manager, desktop, documents, document, ranges, nm, failed
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "names", label
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
Set ranges = document.NamedRanges

Set nm = workbook.Names.Add("Data", "=" & sheet.Name & "!$A$1:$B$2")
Check ranges.hasByName("Data"), "Names.Add creates a named range"
Check ranges.getByName("Data").getContent() = "$" & sheet.Name & ".$A$1:$B$2", "Names.Add stores UNO content"
Check nm.Name = "Data", "Added Name object"
Check nm.RefersTo = "=" & sheet.Name & "!$A$1:$B$2", "RefersTo reads Excel syntax"
Check nm.Value = nm.RefersTo, "Value equals RefersTo"
Check nm = nm.RefersTo, "Default member is RefersTo"
Check nm.RefersToR1C1 = "=" & sheet.Name & "!R1C1:R2C2", "RefersToR1C1 reads R1C1"
Check nm.NameLocal = "Data", "NameLocal"
Check nm.Visible = True, "Visible reads True"
Check nm.RefersToRange.Rows.Count = 2, "RefersToRange still resolves"
Check workbook.Names.Count = 1, "Names.Count after Add"
Check nm.Index = 1, "Index"

sheet.Range("C3").Value = 7
nm.RefersTo = "=" & sheet.Name & "!$C$3"
Check ranges.getByName("Data").getContent() = "$" & sheet.Name & ".$C$3", "RefersTo assignment"
nm.RefersToR1C1 = "=" & sheet.Name & "!R4C4:R5C5"
Check ranges.getByName("Data").getContent() = "$" & sheet.Name & ".$D$4:$E$5", "RefersToR1C1 assignment"

Set nm = workbook.Names.Add("Constant", "=10*2")
Check ranges.getByName("Constant").getContent() = "10*2", "Formula content"
Set nm = workbook.Names.Add("Data", "=" & sheet.Name & "!$A$1")
Check ranges.getByName("Data").getContent() = "$" & sheet.Name & ".$A$1", "Names.Add replaces an existing name"
Set nm = workbook.Names.Add(, , , , , , "Local", , , "=" & sheet.Name & "!R1C2")
Check ranges.getByName("Local").getContent() = "$" & sheet.Name & ".$B$1", "NameLocal with RefersToR1C1"
Check workbook.Names.Count = 3, "Names.Count after three names"
Check workbook.Names(workbook.Names("Local").Index).Name = "Local", "Index selects the same name"

On Error Resume Next
workbook.Names.Add "Hidden", "=1", False
failed = Err.Number <> 0
Err.Clear
On Error GoTo 0
Check failed And Not ranges.hasByName("Hidden"), "Hidden names are rejected, not ignored"

workbook.Names("Constant").Delete
Check Not ranges.hasByName("Constant"), "Name.Delete removes the named range"
Check workbook.Names.Count = 2, "Names.Count after Delete"

Set nm = Nothing
Set ranges = Nothing
Set document = Nothing
Set documents = Nothing
Set desktop = Nothing
Set manager = Nothing
Set sheet = Nothing
workbook.Close False
Set workbook = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Names: PASS"
