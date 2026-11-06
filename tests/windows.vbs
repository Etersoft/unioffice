Option Explicit
' Application/Workbook Windows, ActiveWindow and Window view settings
' checked against the backend controller and frame.
Dim excel, first, second, manager, desktop, windows, win, item, count, names, cell, failed
Sub Check(condition, label)
    If Not condition Then Err.Raise vbObjectError + 1, "windows", label
    WScript.Echo label & ": OK"
End Sub
Function Controller(book)
    Set Controller = BackendDocument(book).getCurrentController()
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
    Err.Raise vbObjectError + 2, "windows", "Backend document not found"
End Function
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set manager = CreateObject("com.sun.star.ServiceManager")
Set desktop = manager.createInstance("com.sun.star.frame.Desktop")

Check excel.ActiveWindow Is Nothing, "ActiveWindow is Nothing without workbooks"
Check excel.Windows.Count = 0, "No windows without workbooks"
Set first = excel.Workbooks.Add()
first.Worksheets(1).Range("A1").Value = first.Name
Set second = excel.Workbooks.Add()
second.Worksheets(1).Range("A1").Value = second.Name

Set windows = excel.Windows
Check windows.Count = 2, "Application.Windows.Count"
Check windows(1).Parent.Name = first.Name, "Windows(1) shows the first workbook"
Check windows.Item(2).Parent.Name = second.Name, "Windows.Item(2)"
Check windows(second.Name).Index = 2, "Windows by name"
count = 0
For Each item In windows
    count = count + 1
Next
Check count = 2, "For Each over Windows"
Check first.Windows.Count = 1, "Workbook.Windows.Count"
Check first.Windows(1).Parent.Name = first.Name, "Workbook.Windows(1)"
Check windows(1).Application Is excel, "Window.Application"
Check windows(1).Type = 1, "Window.Type is xlWorkbook"
Check windows(1).WindowNumber = 1, "WindowNumber"

' Activate changes the active workbook and ActiveWindow.
windows(1).Activate
Check excel.ActiveWorkbook.Name = first.Name, "Window.Activate selects the workbook"
Check excel.ActiveWindow.Parent.Name = first.Name, "ActiveWindow follows activation"
excel.ActiveWindow.ActivateNext
Check excel.ActiveWorkbook.Name = second.Name, "ActivateNext"
excel.ActiveWindow.ActivatePrevious
Check excel.ActiveWorkbook.Name = first.Name, "ActivatePrevious"
second.Activate
Check excel.ActiveWorkbook.Name = second.Name, "Workbook.Activate"
first.Activate

Set win = first.Windows(1)
win.Caption = "Probe caption"
Check win.Caption = "Probe caption", "Caption reads back"
Check Controller(first).getFrame().Title = "Probe caption", "Caption sets the frame title"

' View settings map to SpreadsheetView properties.
win.DisplayGridlines = False
Check Controller(first).ShowGrid = False, "DisplayGridlines"
Check win.DisplayGridlines = False, "DisplayGridlines reads back"
win.DisplayGridlines = True
win.DisplayFormulas = True
Check Controller(first).ShowFormulas = True, "DisplayFormulas"
win.DisplayFormulas = False
win.DisplayHeadings = False
Check Controller(first).HasColumnRowHeaders = False, "DisplayHeadings"
win.DisplayHeadings = True
win.DisplayHorizontalScrollBar = False
Check Controller(first).HasHorizontalScrollBar = False, "DisplayHorizontalScrollBar"
win.DisplayVerticalScrollBar = False
Check Controller(first).HasVerticalScrollBar = False, "DisplayVerticalScrollBar"
win.DisplayWorkbookTabs = False
Check Controller(first).HasSheetTabs = False, "DisplayWorkbookTabs"
win.DisplayZeros = False
Check Controller(first).ShowZeroValues = False, "DisplayZeros"
win.DisplayOutline = False
Check Controller(first).IsOutlineSymbolsSet = False, "DisplayOutline"
win.GridlineColor = RGB(255, 0, 0)
Check Controller(first).GridColor = &HFF0000, "GridlineColor converts to UNO RGB"
Check win.GridlineColor = RGB(255, 0, 0), "GridlineColor reads back"
win.Zoom = 150
Check Controller(first).ZoomValue = 150, "Zoom"
Check win.Zoom = 150, "Zoom reads back"
On Error Resume Next
win.Zoom = 5
failed = Err.Number <> 0
Err.Clear
On Error GoTo 0
Check failed, "Zoom outside 10..400 is rejected"
win.ScrollRow = 11
Check Controller(first).getFirstVisibleRow() = 10, "ScrollRow"
Check win.ScrollRow = 11, "ScrollRow reads back"
win.ScrollColumn = 3
Check win.ScrollColumn = 3, "ScrollColumn"
win.SmallScroll 2, , 1
Check win.ScrollRow = 13 And win.ScrollColumn = 4, "SmallScroll down and right"
win.SmallScroll , 20, , 10
Check win.ScrollRow = 1 And win.ScrollColumn = 1, "SmallScroll clamps at A1"

Check win.VisibleRange.Row = win.ScrollRow And win.VisibleRange.Column = win.ScrollColumn, "VisibleRange starts at the scroll position"

' Selection, ActiveCell and ActiveSheet come from the controller.
first.Worksheets(1).Range("C3:D4").Select
Check win.RangeSelection.Address = "$C$3:$D$4", "RangeSelection"
Check win.Selection.Address = "$C$3:$D$4", "Selection"
Set cell = win.ActiveCell
Check cell.Address = "$C$3", "ActiveCell"
Check win.ActiveSheet.Name = first.Worksheets(1).Name, "Window.ActiveSheet"
Check excel.ActiveWindow.Parent.Name = first.Name, "ActiveWindow after Select"

' Workbook.Name is the file name, independent of the window caption.
Dim path
Check InStr(second.Name, " - ") = 0, "Unsaved Workbook.Name has no window suffix"
path = "C:\users\Public\uo window name.xls"
second.SaveAs path
Check second.Name = "uo window name.xls", "Saved Workbook.Name is the decoded file name"
Check excel.Workbooks("uo window name.xls").Name = second.Name, "Workbooks by name"
second.Worksheets(1).Range("A1").Value = second.Name

' Closing the window closes its workbook.
second.Windows(1).Close False
Check excel.Workbooks.Count = 1, "Window.Close closes the workbook"
Check excel.Windows.Count = 1, "Windows.Count after Close"

Set cell = Nothing
Set win = Nothing
Set item = Nothing
Set windows = Nothing
first.Close False
Set first = Nothing
Set second = Nothing
Set desktop = Nothing
Set manager = Nothing
excel.Quit
Set excel = Nothing
WScript.Echo "Windows: PASS"
