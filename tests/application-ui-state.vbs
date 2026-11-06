Option Explicit
' Application.ScreenUpdating and Interactive without an active workbook,
' for existing and later workbooks, with balanced controller locks.
Dim excel, first, second, manager, desktop
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "application-ui-state", label
    WScript.Echo label & ": OK"
End Sub
Function Window(book)
    Set Window = BackendDocument(book).getCurrentController().getFrame().getContainerWindow()
End Function
Function BackendDocument(book)
    Dim documents, document
    Set documents = desktop.getComponents().createEnumeration()
    Do While documents.hasMoreElements()
        Set document = documents.nextElement()
        If document.getSheets().getByIndex(0).getCellByPosition(0, 0).getString() = book.Name Then
            Set BackendDocument = document
            Exit Function
        End If
    Loop
    Err.Raise vbObjectError + 2, "application-ui-state", "Backend document not found"
End Function
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set manager = CreateObject("com.sun.star.ServiceManager")
Set desktop = manager.createInstance("com.sun.star.frame.Desktop")

excel.ScreenUpdating = False
Check excel.ScreenUpdating = False, "ScreenUpdating=False without workbook"
excel.Interactive = False
Check excel.Interactive = False, "Interactive=False without workbook"

Set first = excel.Workbooks.Add()
first.Worksheets(1).Range("A1").Value = first.Name
Check BackendDocument(first).hasControllersLocked() = True, "New workbook starts locked"
Check Window(first).isEnabled() = False, "New workbook window disabled"

excel.ScreenUpdating = False
excel.ScreenUpdating = True
Check BackendDocument(first).hasControllersLocked() = False, "Repeated False needs one True"
excel.Interactive = True
Check Window(first).isEnabled() = True, "Interactive=True enables window"
Check excel.Interactive = True, "Interactive reads True"

Set second = excel.Workbooks.Add()
second.Worksheets(1).Range("A1").Value = second.Name
Check BackendDocument(second).hasControllersLocked() = False, "Workbook added while updating is unlocked"
excel.ScreenUpdating = False
excel.Interactive = False
Check BackendDocument(first).hasControllersLocked() = True, "False locks first workbook"
Check BackendDocument(second).hasControllersLocked() = True, "False locks second workbook"
Check Window(second).isEnabled() = False, "Interactive=False disables every window"
excel.ScreenUpdating = True
excel.Interactive = True
Check BackendDocument(first).hasControllersLocked() = False, "True unlocks first workbook"
Check BackendDocument(second).hasControllersLocked() = False, "True unlocks second workbook"
Check Window(second).isEnabled() = True, "Interactive=True enables every window"

second.Close False
first.Close False
Set second = Nothing
Set first = Nothing
Set desktop = Nothing
Set manager = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Application UI state: PASS"
