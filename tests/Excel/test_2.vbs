'+++++++++++++++++++++++++++++++++++++++++++++
' test 
'         IRange->Group
'         IRange->UnGroup
'         IWorksheet->Outline
'         IOutline->Showlevel
'         IApplication.SheetsInNewWorkbook
'+++++++++++++++++++++++++++++++++++++++++++++


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

'Не забыть скрыть Excel если этопотоковый запуск.
TEST_NAME ("TEST 2")
On Error Resume Next

'Создаем объект Excel.Application
Set Excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    ERROR_MES ("create Excel.Application")  
    Err.Clear
Else 
    OK_MES ("create Excel.Application")     
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
    OK_MES("PUT Excel.SheetsInNewWorkbook")   
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

Excel.Sheets(1).Range("B1:H9").Group
Excel.Sheets(1).Range("C2:G8").Group
Excel.Sheets(1).Range("B11:B15").Rows.Group
Excel.Sheets(1).Range("C12:D14").Rows.Group
If Err.Number <> 0 Then
    ERROR_MES ("GET IRange.Group")  
    Err.Clear
Else 
    OK_MES ("GET IRange.Group")  
End If

Excel.Sheets(2).Range("B1:H9").Group
Excel.Sheets(2).Range("C2:G8").Group
Excel.Sheets(2).Range("B11:B15").Rows.Group
Excel.Sheets(2).Range("C12:D14").Rows.Group
If Err.Number <> 0 Then
    ERROR_MES ("GET IRange.Group")  
    Err.Clear
Else 
    OK_MES("GET IRange.Group")  
End If

Excel.Sheets(2).Range("C2:G8").Ungroup
Excel.Sheets(2).Range("C12:D14").Rows.Ungroup
If Err.Number <> 0 Then
    ERROR_MES ("GET IRange.Ungroup")  
    Err.Clear
Else 
    OK_MES ("GET IRange.Ungroup")  
End If

Set outline = Excel.Sheets(1).Outline
If Err.Number <> 0 Then
    ERROR_MES ("GET IWorksheet.Outline")  
    Err.Clear
Else 
    OK_MES ("GET IWorksheet.Outline")   
End If

outline.Showlevels 1,1
If Err.Number <> 0 Then
    ERROR_MES ("GET IOutline.Showlevels")  
    Err.Clear
Else 
    OK_MES ("GET IOutline.Showlevels")   
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
