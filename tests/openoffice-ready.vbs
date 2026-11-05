Option Explicit
Dim manager, attempt
WScript.Echo "Script host: " & WScript.FullName
On Error Resume Next
For attempt = 1 To 30
    Err.Clear
    Set manager = CreateObject("com.sun.star.ServiceManager")
    If Err.Number = 0 Then Exit For
    WScript.Sleep 1000
Next
If Err.Number <> 0 Then
    WScript.Echo "OpenOffice readiness: FAILED " & Hex(Err.Number) & " " & Err.Description
    WScript.Quit 1
End If
Set manager = Nothing
WScript.Echo "OpenOffice readiness: PASS"
