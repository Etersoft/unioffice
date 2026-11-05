Option Explicit
Dim manager, desktop, document, sheet, attempt
On Error Resume Next
For attempt = 1 To 30
    Err.Clear
    Set manager = CreateObject("com.sun.star.ServiceManager")
    If Err.Number = 0 Then Exit For
    WScript.Sleep 1000
Next
If Err.Number <> 0 Then
    WScript.Echo "ServiceManager: FAILED " & Hex(Err.Number) & " " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "ServiceManager: OK"
Set desktop = manager.createInstance("com.sun.star.frame.Desktop")
WScript.Echo "Desktop: OK"
Set document = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, Array())
Set sheet = document.Sheets.getByIndex(0)
sheet.getCellByPosition(0, 0).setString "UniOffice backend test"
If sheet.getCellByPosition(0, 0).getString() <> "UniOffice backend test" Then
    WScript.Echo "Cell round trip: FAILED"
    document.close True
    WScript.Quit 1
End If
WScript.Echo "Calc cell round trip: OK"
document.close True
Set sheet = Nothing
Set document = Nothing
Set manager = Nothing
desktop.terminate
Set desktop = Nothing
WScript.Echo "OpenOffice COM backend: PASS"
