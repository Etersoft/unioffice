'++++++++++++++++++++++++++++++++++++
'test
'     IPageSetup
'++++++++++++++++++++++++++++++++++++
Dim WshShell
Dim fso
Dim gsuccess, gfailed, show_excel, next_script
Dim success, failed
Dim otchetFile


'Объект для запуска других скриптов
Set WshShell = Wscript.CreateObject("Wscript.Shell")
'объект для работы с файлами
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

success = 0
failed = 0

If Wscript.Arguments.Count = 3 Then
    next_script = WScript.Arguments(0)
    gsuccess = WScript.Arguments(1)
    gfailed = WScript.Arguments(2)
    show_excel = false
    Set otchetFile = fso.OpenTextFile("otchet.txt", 8)
else 
    gsuccess = 0
    gfailed = 0
    show_excel = true
    Set otchetFile = fso.CreateTextFile("otchet.txt", True)
End If

'Не забыть скрыть Excel если этопотоковый запуск.
otchetFile.WriteLine("++++++ TEST 4 (IPageSetup) ++++++")
On Error Resume Next

'Создаем объект Excel.Application
Set Excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] create Excel.Application")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] create Excel.Application")  
    success = success + 1   
End If 

If show_excel then
'Показываем Excel
    Excel.Visible = TRUE
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] PUT Excel.Visible")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] PUT Excel.Visible")  
        success = success + 1   
    End If 
Else
'Не показываем Excel. По умолчанию он и так не покажеться, но для проверки свойства выполняем
    Excel.Visible = FALSE
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] PUT Excel.Visible")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] PUT Excel.Visible")  
        success = success + 1   
    End If 
End If

'Отключаем предупреждения
Excel.DisplayAlerts = FALSE
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT Excel.DisplayAlerts")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT Excel.DisplayAlerts")  
    success = success + 1 
End If

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Здесь помещается текст тестового скрипта
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Excel.SheetsInNewWorkbook = 2
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT Excel.SheetsInNewWorkbook")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT Excel.SheetsInNewWorkbook")  
    success = success + 1 
End If

Set WB = Excel.Workbooks.Add
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] Excel.Workbooks.Add")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] Excel.Workbooks.Add")  
    success = success + 1 
End If

Dim count
count = Excel.Sheets.count
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] GET Excel.Sheets.count")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] GET Excel.Sheets.count")  
    success = success + 1 
End If
If count<>2 then
    otchetFile.WriteLine("[FAILED] NOT EQUAL Excel.Sheets.count = 2")  
    failed = failed + 1
Else
   otchetFile.WriteLine("[SUCCESS] EQUAL Excel.Sheets.count = 2")  
    success = success + 1 
End If

Set otarget = Excel.Sheets(1)
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] GET Excel.Sheets(1)")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] GET Excel.Sheets(1)")  
    success = success + 1 
End If

'LeftMargin
otarget.PageSetup.LeftMargin = 10
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.LeftMargin")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.LeftMargin")  
    success = success + 1 
End If
If otarget.PageSetup.LeftMargin=10 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.LeftMargin")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.LeftMargin")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.LeftMargin  = " + otarget.PageSetup.LeftMargin)  
    failed = failed + 1
End If

'RightMargin
otarget.PageSetup.RightMargin = 20
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.RightMargin")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.RightMargin")  
    success = success + 1 
End If
If otarget.PageSetup.RightMargin=20 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.RightMargin")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.RightMargin")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.RightMargin  = " + otarget.PageSetup.RightMargin)  
    failed = failed + 1
End If

'TopMargin
otarget.PageSetup.TopMargin = 30
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.TopMargin")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.TopMargin")  
    success = success + 1 
End If
If otarget.PageSetup.TopMargin=30 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.TopMargin")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.TopMargin")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.TopMargin  = " + otarget.PageSetup.TopMargin)  
    failed = failed + 1
End If

'BottomMargin
otarget.PageSetup.BottomMargin = 40
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.BottomMargin")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.BottomMargin")  
    success = success + 1 
End If
If otarget.PageSetup.BottomMargin=40 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.BottomMargin")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.BottomMargin")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.BottomMargin  = " + otarget.PageSetup.BottomMargin)  
    failed = failed + 1
End If

'Orientation
otarget.PageSetup.Orientation = 2
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.Orientation")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.Orientation")  
    success = success + 1 
End If
If otarget.PageSetup.Orientation=2 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.Orientation")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.Orientation")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.Orientation  = " + otarget.PageSetup.Orientation)  
    failed = failed + 1
End If

'Zoom
otarget.PageSetup.Zoom = 200
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.Zoom")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.Zoom")  
    success = success + 1 
End If
If otarget.PageSetup.Zoom=200 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.Zoom")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.Zoom")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.Zoom  = " + otarget.PageSetup.Zoom)  
    failed = failed + 1
End If

'CenterHorizontally
otarget.PageSetup.CenterHorizontally = TRUE
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.CenterHorizontally")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.CenterHorizontally")  
    success = success + 1 
End If
If otarget.PageSetup.CenterHorizontally=TRUE then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.CenterHorizontally")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.CenterHorizontally")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.CenterHorizontally")  
    failed = failed + 1
End If

'CenterVertically
otarget.PageSetup.CenterVertically = TRUE
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.CenterVertically")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.CenterVertically")  
    success = success + 1 
End If
If otarget.PageSetup.CenterVertically=TRUE then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.CenterVertically")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.CenterVertically")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.CenterVertically")  
    failed = failed + 1
End If

'FooterMargin
otarget.PageSetup.FooterMargin = 50
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.FooterMargin")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.FooterMargin")  
    success = success + 1 
End If
If otarget.PageSetup.FooterMargin=50 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.FooterMargin")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.FooterMargin")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.FooterMargin  = " + otarget.PageSetup.FooterMargin)  
    failed = failed + 1
End If

'HeaderMargin
otarget.PageSetup.HeaderMargin = 60
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.HeaderMargin")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.HeaderMargin")  
    success = success + 1 
End If
If otarget.PageSetup.HeaderMargin=60 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.HeaderMargin")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.HeaderMargin")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.HeaderMargin  = " + otarget.PageSetup.HeaderMargin)  
    failed = failed + 1
End If

'FitToPagesTall
otarget.PageSetup.FitToPagesTall = 3
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.FitToPagesTall")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.FitToPagesTall")  
    success = success + 1 
End If
If otarget.PageSetup.FitToPagesTall=3 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.FitToPagesTall")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.FitToPagesTall")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.FitToPagesTall  = " + otarget.PageSetup.FitToPagesTall)  
    failed = failed + 1
End If

'FitToPagesWide
otarget.PageSetup.FitToPagesWide = 2
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IPageSetup.FitToPagesWide")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IPageSetup.FitToPagesWide")  
    success = success + 1 
End If
If otarget.PageSetup.FitToPagesWide=2 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IPageSetup.FitToPagesWide")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IPageSetup.FitToPagesWide")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IPageSetup.FitToPagesWide  = " + otarget.PageSetup.FitToPagesWide)  
    failed = failed + 1
End If

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'конец кода теста
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

if show_excel then
   MsgBox "Тест пройден. Результат в файле otchet.txt"
End If

'Закрываем Excel
Excel.Quit
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] Excel.Quit")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] Excel.Quit")  
    success = success + 1   
End If 


otchetFile.WriteLine("Всего тестов - " & success+failed & "  Успешно - " & success & "  Провалено - " & failed)
otchetFile.Close

'Если это потоковый запуск, то запускает следующий тест.

if show_excel=false then   
    gsuccess = gsuccess + success
    gfailed = gfailed + failed
    next_script = next_script + 1
    Dim param, filename
'заполняем параметры запуска
    param = " " & next_script & " " & gsuccess & " " & gfailed
    filename = "test_" & next_script & ".vbs"
    'запускаем скрипт на исполнение    
    if fso.FileExists(filename) then 
        WshShell.Run "cscript.exe /E:vbscript "& filename & " " & param, 0, FALSE
    Else
        Set otchetFile = fso.OpenTextFile("otchet.txt", 8)  
        otchetFile.WriteLine("Всего тестов - " & gsuccess+gfailed & "  Успешно - " & gsuccess & "  Провалено - " & gfailed)
        otchetFile.Close 
        MsgBox "Все тесты пройдены. Результаты в файле otchet.txt"
    End If
End If