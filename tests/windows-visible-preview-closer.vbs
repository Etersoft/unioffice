' Helper for windows-visible.vbs: records that the workbook entered page
' preview, then closes the preview so the modal PrintPreview call returns.
Dim manager, desktop, documents, document, controller, fso, report, i, state, isView
Set fso = CreateObject("Scripting.FileSystemObject")
Set report = fso.CreateTextFile(WScript.Arguments(0), True)
Set manager = CreateObject("com.sun.star.ServiceManager")
Set desktop = manager.createInstance("com.sun.star.frame.Desktop")
state = "not-found"
For i = 1 To 40
    WScript.Sleep 500
    Set documents = desktop.getComponents().createEnumeration()
    Do While documents.hasMoreElements()
        Set document = documents.nextElement()
        If document.getSheets().getByIndex(0).getCellByPosition(0, 0).getString() = WScript.Arguments(1) Then
            Set controller = document.getCurrentController()
            ' The preview controller has no XServiceInfo in OpenOffice 4.
            On Error Resume Next
            isView = controller.supportsService("com.sun.star.sheet.SpreadsheetView")
            If Err.Number <> 0 Then isView = False
            Err.Clear
            On Error GoTo 0
            If Not isView Then
                state = "preview"
                ' Report before closing: closing lets PrintPreview return.
                report.WriteLine state
                report.Close
                manager.createInstance("com.sun.star.frame.DispatchHelper").executeDispatch _
                    controller.getFrame(), ".uno:ClosePreview", "", 0, Array()
                WScript.Quit 0
            End If
        End If
    Loop
Next
report.WriteLine state
report.Close
