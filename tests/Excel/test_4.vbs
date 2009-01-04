'++++++++++++++++++++++++++++++++++++
'test
'     IPageSetup
'++++++++++++++++++++++++++++++++++++

Dim WshShell
Dim fso
Dim gsuccess, gfailed, show_excel, next_script
Dim success, failed
Dim otchetFile

Function ERROR_MES ( mes_err )
    otchetFile.WriteLine("[FAILED] " +  mes_err )    
    failed = failed + 1
End Function


Function OK_MES( mes_ok )
    otchetFile.WriteLine("[SUCCESS] " +  mes_ok ) 
    success = success + 1
End Function

Function TEST_NAME( name_test )
    otchetFile.WriteLine("+++++   " +  name_test  + "    +++++") 
End Function

success = 0
failed = 0

'Объект для запуска других скриптов
Set WshShell = Wscript.CreateObject("Wscript.Shell")
'объект для работы с файлами
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

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

'Не забыть скрыть Excel если это потоковый запуск.
TEST_NAME("TEST 4 (IPageSetup)")
On Error Resume Next

'Создаем объект Excel.Application
Set Excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    ERROR_MES("create Excel.Application")  
    Err.Clear
Else 
    OK_MES("create Excel.Application")    
End If 

If show_excel then
'Показываем Excel
    Excel.Visible = TRUE
    If Err.Number <> 0 Then
        ERROR_MES ("PUT Excel.Visible")  
        Err.Clear
    Else 
        OK_MES ("PUT Excel.Visible")     
    End If 
Else
'Не показываем Excel. По умолчанию он и так не покажеться, но для проверки свойства выполняем
    Excel.Visible = FALSE
    If Err.Number <> 0 Then
        ERROR_MES ("PUT Excel.Visible")  
        Err.Clear
    Else 
        OK_MES ("PUT Excel.Visible")   
    End If 
End If

'Отключаем предупреждения
Excel.DisplayAlerts = FALSE
If Err.Number <> 0 Then
    ERROR_MES ("PUT Excel.DisplayAlerts")  
    Err.Clear
Else 
    OK_MES ("PUT Excel.DisplayAlerts")   
End If

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Здесь помещается текст тестового скрипта
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Excel.SheetsInNewWorkbook = 2
If Err.Number <> 0 Then
    ERROR_MES ("PUT Excel.SheetsInNewWorkbook")  
    Err.Clear
Else 
    OK_MES ("PUT Excel.SheetsInNewWorkbook")  
End If

Set WB = Excel.Workbooks.Add
If Err.Number <> 0 Then
    ERROR_MES ("Excel.Workbooks.Add")  
    Err.Clear
Else 
    OK_MES ("Excel.Workbooks.Add")  
End If

Dim count
count = Excel.Sheets.count
If Err.Number <> 0 Then
    ERROR_MES ("GET Excel.Sheets.count")  
    Err.Clear
Else 
    OK_MES ("GET Excel.Sheets.count")  
End If
If count<>2 then
    ERROR_MES ("NOT EQUAL Excel.Sheets.count = 2")  
Else
    OK_MES ("EQUAL Excel.Sheets.count = 2")   
End If

Set otarget = Excel.Sheets(1)
If Err.Number <> 0 Then
    ERROR_MES ("GET Excel.Sheets(1)")  
    Err.Clear
Else 
    OK_MES ("GET Excel.Sheets(1)")  
End If

'LeftMargin
otarget.PageSetup.LeftMargin = 10
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.LeftMargin")  
    Err.Clear
Else 
    OK_MES ("PUT IPageSetup.LeftMargin")  
End If
If otarget.PageSetup.LeftMargin=10 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.LeftMargin")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.LeftMargin")     
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.LeftMargin  = " + otarget.PageSetup.LeftMargin)  
End If

'RightMargin
otarget.PageSetup.RightMargin = 20
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.RightMargin")  
    Err.Clear
Else 
    OK_MES ("PUT IPageSetup.RightMargin")  
End If
If otarget.PageSetup.RightMargin=20 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.RightMargin")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.RightMargin")       
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.RightMargin  = " + otarget.PageSetup.RightMargin)  
End If

'TopMargin
otarget.PageSetup.TopMargin = 30
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.TopMargin")  
    Err.Clear
Else 
    OK_MES ("PUT IPageSetup.TopMargin")  
End If
If otarget.PageSetup.TopMargin=30 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.TopMargin")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.TopMargin")     
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.TopMargin  = " + otarget.PageSetup.TopMargin)  
End If

'BottomMargin
otarget.PageSetup.BottomMargin = 40
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.BottomMargin")  
    Err.Clear
Else 
    OK_MES ("PUT IPageSetup.BottomMargin")  
End If
If otarget.PageSetup.BottomMargin=40 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.BottomMargin")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.BottomMargin")     
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.BottomMargin  = " + otarget.PageSetup.BottomMargin)  
End If

'Orientation
otarget.PageSetup.Orientation = 2
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.Orientation")  
    Err.Clear
Else 
    OK_MES("PUT IPageSetup.Orientation")  
End If
If otarget.PageSetup.Orientation=2 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.Orientation")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.Orientation")      
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.Orientation  = " + otarget.PageSetup.Orientation)  
End If

'Zoom
otarget.PageSetup.Zoom = 200
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.Zoom")  
    Err.Clear
Else 
    OK_MES ("PUT IPageSetup.Zoom")  
    success = success + 1 
End If
If otarget.PageSetup.Zoom=200 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.Zoom")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.Zoom")       
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.Zoom  = " + otarget.PageSetup.Zoom)  
End If

'CenterHorizontally
otarget.PageSetup.CenterHorizontally = TRUE
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.CenterHorizontally")  
    Err.Clear
Else 
    OK_MES ("PUT IPageSetup.CenterHorizontally")  
End If
If otarget.PageSetup.CenterHorizontally=TRUE then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.CenterHorizontally")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.CenterHorizontally")       
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.CenterHorizontally")  
End If

'CenterVertically
otarget.PageSetup.CenterVertically = TRUE
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.CenterVertically")  
    Err.Clear
Else 
    OK_MES ("PUT IPageSetup.CenterVertically")  
End If
If otarget.PageSetup.CenterVertically=TRUE then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.CenterVertically")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.CenterVertically")      
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.CenterVertically")  
End If

'FooterMargin
otarget.PageSetup.FooterMargin = 50
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.FooterMargin")  
    Err.Clear
Else 
    OK_MES ("PUT IPageSetup.FooterMargin")  
End If
If otarget.PageSetup.FooterMargin=50 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.FooterMargin")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.FooterMargin")       
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.FooterMargin  = " + otarget.PageSetup.FooterMargin)  
End If

'HeaderMargin
otarget.PageSetup.HeaderMargin = 60
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.HeaderMargin")  
    Err.Clear
Else 
    OK_MES ("PUT IPageSetup.HeaderMargin")  
End If
If otarget.PageSetup.HeaderMargin=60 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.HeaderMargin")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.HeaderMargin")      
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.HeaderMargin  = " + otarget.PageSetup.HeaderMargin)  
End If

'FitToPagesTall
otarget.PageSetup.FitToPagesTall = 3
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.FitToPagesTall")  
    Err.Clear
Else 
    OK_MES ("PUT IPageSetup.FitToPagesTall")  
End If
If otarget.PageSetup.FitToPagesTall=3 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.FitToPagesTall")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.FitToPagesTall")     
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.FitToPagesTall  = " + otarget.PageSetup.FitToPagesTall)  
End If

'FitToPagesWide
otarget.PageSetup.FitToPagesWide = 2
If Err.Number <> 0 Then
    ERROR_MES ("PUT IPageSetup.FitToPagesWide")  
    Err.Clear
Else 
    OK_MES ("PUT IPageSetup.FitToPagesWide")  
End If
If otarget.PageSetup.FitToPagesWide=2 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IPageSetup.FitToPagesWide")  
        Err.Clear
    Else 
        OK_MES ("GET IPageSetup.FitToPagesWide")       
   End If
Else
    ERROR_MES ("NOT EQUAL IPageSetup.FitToPagesWide  = " + otarget.PageSetup.FitToPagesWide)  
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
    ERROR_MES ("Excel.Quit")  
    Err.Clear
Else 
    OK_MES ("Excel.Quit")    
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
