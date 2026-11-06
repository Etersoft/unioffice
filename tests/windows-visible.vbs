Option Explicit
' Window members that need a visible document window: geometry, state,
' frozen panes, splits, scrolling and the modal print preview.
Dim excel, book, win, manager, desktop, shell, fso, marker, report, before, started, failed
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "windows-visible", label
    WScript.Echo label & ": OK"
End Sub
Function Near(value, expected)
    Near = Abs(value - expected) <= 2
End Function
Function Controller()
    Dim documents, document
    Set documents = desktop.getComponents().createEnumeration()
    Do While documents.hasMoreElements()
        Set document = documents.nextElement()
        If document.getSheets().getByIndex(0).getCellByPosition(0, 0).getString() = marker Then
            Set Controller = document.getCurrentController()
            Exit Function
        End If
    Loop
    Err.Raise vbObjectError + 2, "windows-visible", "Backend document not found"
End Function
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
Set book = excel.Workbooks.Add()
marker = "visible-window-probe"
book.Worksheets(1).Range("A1").Value = marker
Set manager = CreateObject("com.sun.star.ServiceManager")
Set desktop = manager.createInstance("com.sun.star.frame.Desktop")
Set win = excel.ActiveWindow

Check win.Visible = True, "Window.Visible"
' Window managers apply state changes asynchronously.
excel.WindowState = -4143
WScript.Sleep 1000
win.Left = 150
win.Top = 100
win.Width = 480
win.Height = 360
WScript.Sleep 500
WScript.Echo "Geometry: state=" & win.WindowState & " left=" & win.Left & " top=" & win.Top & " width=" & win.Width & " height=" & win.Height
Check Near(win.Width, 480) And Near(win.Height, 360), "Width/Height in points"
Check Near(win.Left, 150) And Near(win.Top, 100), "Left/Top in points"
Check win.UsableWidth > 0 And win.UsableHeight > 0, "UsableWidth/UsableHeight"

win.WindowState = -4137
WScript.Sleep 1000
Check win.WindowState = -4137, "Window.WindowState maximized"
Check excel.WindowState = -4137, "Application.WindowState follows the active window"
excel.WindowState = -4143
WScript.Sleep 1000
Check win.WindowState = -4143, "Application.WindowState normal"
win.WindowState = -4140
WScript.Sleep 1000
Check win.WindowState = -4140, "Window.WindowState minimized"
win.WindowState = -4143
WScript.Sleep 1000
Check win.WindowState = -4143, "Restore from minimized"
win.WindowState = -4137
WScript.Sleep 1000

' Freeze panes at C4: two columns and three rows stay visible.
win.ScrollRow = 1
win.ScrollColumn = 1
book.Worksheets(1).Range("C4").Select
win.FreezePanes = True
Check Controller().hasFrozenPanes() = True, "FreezePanes"
Check win.FreezePanes = True, "FreezePanes reads back"
Check win.SplitRow = 3 And win.SplitColumn = 2, "SplitRow/SplitColumn of frozen panes"
Check win.Split = True, "Split is True with frozen panes"
win.SplitRow = 5
Check win.SplitRow = 5 And win.SplitColumn = 2, "SplitRow moves the frozen rows"
win.FreezePanes = False
Check Controller().hasFrozenPanes() = False, "FreezePanes = False"
win.Split = False
Check win.Split = False, "Split = False"

' Visible range and page scrolling.
win.ScrollRow = 1
before = win.VisibleRange.Rows.Count
Check before > 1 And win.VisibleRange.Row = 1, "VisibleRange"
win.LargeScroll 1
Check win.ScrollRow > 1, "LargeScroll moves down by a page"
win.LargeScroll , 1
Check win.ScrollRow = 1, "LargeScroll back up"

' PrintPreview is modal: the helper observes the preview and closes it.
report = shell.ExpandEnvironmentStrings("%TEMP%") & "\unioffice-preview-" & CStr(Int(Timer * 1000)) & ".txt"
shell.Run "cscript.exe //nologo """ & fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "windows-visible-preview-closer.vbs") & """ """ & report & """ """ & marker & """", 0, False
started = Timer
book.PrintPreview
Check fso.FileExists(report), "PrintPreview returned after the helper finished"
Check Trim(fso.OpenTextFile(report).ReadLine()) = "preview", "PrintPreview showed the page preview"
Check Timer - started >= 0.5, "PrintPreview waited for the preview to close"
Check Controller().getActiveSheet().getName() = book.Worksheets(1).Name, "Normal view after PrintPreview"
fso.DeleteFile report

Set win = Nothing
book.Close False
Set book = Nothing
Set desktop = Nothing
Set manager = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Visible windows: PASS"
