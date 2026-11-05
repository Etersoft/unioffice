Option Explicit
Dim manager, desktop, document, target, info, table, line, name, value, errorNumber
Dim reflection, description
Set manager = CreateObject("com.sun.star.ServiceManager")
Set desktop = manager.createInstance("com.sun.star.frame.Desktop")
Set document = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, Array())
Set target = document.Sheets.getByIndex(0).getCellRangeByName("A1:D4")
Set info = target.getPropertySetInfo()
For Each name In Array("TableBorder", "TableBorder2", "LeftBorder2", "DiagonalTLBR", "DiagonalTLBR2")
    WScript.Echo name & ": " & CStr(info.hasPropertyByName(name))
Next
If Not info.hasPropertyByName("TableBorder") Then Err.Raise vbObjectError + 1, "border-capabilities", "Missing TableBorder"
Set table = target.getPropertyValue("TableBorder")
Set line = table.LeftLine
On Error Resume Next
value = line.LineStyle
errorNumber = Err.Number
Err.Clear
On Error GoTo 0
WScript.Echo "Classic BorderLine.LineStyle error: " & Hex(errorNumber)
Set reflection = manager.createInstance("com.sun.star.reflection.CoreReflection")
For Each name In Array("com.sun.star.table.BorderLine", "com.sun.star.table.BorderLine2", "com.sun.star.table.BorderLineStyle")
    Set description = Nothing
    On Error Resume Next
    Set description = reflection.forName(name)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    If description Is Nothing Then
        WScript.Echo name & ": unavailable; error " & Hex(errorNumber)
    Else
        WScript.Echo name & ": available"
    End If
Next
Set description = Nothing
Set reflection = Nothing
Set line = Nothing
Set table = Nothing
Set info = Nothing
Set target = Nothing
document.close True
Set document = Nothing
desktop.terminate
Set desktop = Nothing
Set manager = Nothing
WScript.Echo "OpenOffice border capabilities: PASS"
