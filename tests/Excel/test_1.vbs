'++++++++++++++++++++++++++++++++++++++++++++++
' Этот модуль тестирует все свойства и методы 
' интерфейса IFont.
'++++++++++++++++++++++++++++++++++++++++++++++

Dim WshShell
Dim fso
Dim gsuccess, gfailed, show_excel, next_script
Dim success, failed
Dim otchetFile

Function ERROR_MES ( mes_err )
    otchetFile.WriteLine("!![FAILED]!! " +  mes_err )    
    failed = failed + 1
End Function


Function OK_MES( mes_ok )
    otchetFile.WriteLine("[SUCCESS] " +  mes_ok ) 
    success = success + 1
End Function

Function TEST_NAME( name_test )
    otchetFile.WriteLine("==== " +  name_test  + " ====") 
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
TEST_NAME( "TEST 1 (IFont)" )
On Error Resume Next

'Создаем объект Excel.Application
Set Excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    ERROR_MES( "create Excel.Application" )
    Err.Clear
Else 
    OK_MES ( "create Excel.Application" )
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

Set WB = Excel.Workbooks.Add
If Err.Number <> 0 Then
    ERROR_MES ("create Excel.Workbooks.Add")  
    Err.Clear
Else 
    OK_MES ("create Excel.Workbooks.Add")   
End If

'IFont.Italic
Excel.Cells(1,1).Value = "Italic"
Excel.Cells(1,1).Font.Italic = TRUE
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.Italic")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.Italic")   
End If
If Excel.Cells(1,1).Font.Italic=TRUE then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IFont.Italic")  
    Err.Clear
    Else
        OK_MES ("GET IFont.Italic")      
   End If
Else
    ERROR_MES ("NOT EQUAL IFont.Italic")  
End If

'IFont.Bold
Excel.Cells(2,1).Value = "Bold"
Excel.Cells(2,1).Font.Bold = TRUE
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.Bold")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.Bold")  
End If
If Excel.Cells(2,1).Font.Bold=TRUE then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IFont.Bold")  
        Err.Clear
    Else 
        OK_MES ("GET IFont.Bold")       
   End If
Else
    ERROR_MES ("NOT EQUAL IFont.Bold")  
End If

'IFont.Strikethrough
Excel.Cells(3,1).Value = "Strikethrough"
Excel.Cells(3,1).Font.Strikethrough = TRUE
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.Strikethrough")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.Strikethrough")   
End If
If Excel.Cells(3,1).Font.Strikethrough=TRUE then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IFont.Strikethrough")  
        Err.Clear
    Else
        OK_MES ("GET IFont.Strikethrough")       
    End If
Else
    ERROR_MES ("NOT EQUAL IFont.Strikethrough")  
End If

'IFont.Underline
Excel.Cells(4,1).Value = "Underline" 
Excel.Cells(4,1).Font.Underline = 2
Excel.Cells(4,2).Value = "Underline"
Excel.Cells(4,2).Font.Underline = 4		
Excel.Cells(4,3).Value = "Underline" 
Excel.Cells(4,3).Font.Underline = 5
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.Underline")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.Underline")   
End If
If Excel.Cells(4,1).Font.Underline=2 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IFont.Underline")  
        Err.Clear
    Else
        OK_MES ("GET IFont.Underline")      
    End If
Else
    ERROR_MES ("NOT EQUAL IFont.Underline")  
End If

'SubScript 
Excel.Cells(5,1).Value = "SubScript"
Excel.Cells(5,1).Font.SubScript = TRUE
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.SubScript")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.SubScript")  
End If
If Excel.Cells(5,1).Font.SubScript=TRUE then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IFont.SubScript")  
        Err.Clear
    Else
        OK_MES ("GET IFont.SubScript")       
    End If    
Else
    ERROR_MES ("NOT EQUAL IFont.SubScript")  
End If

'SuperScript 
Excel.Cells(6,1).Value = "SuperScript"
Excel.Cells(6,1).Font.SuperScript = TRUE
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.SuperScript")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.SuperScript")  
End If
If Excel.Cells(6,1).Font.SuperScript=TRUE then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IFont.SuperScript")  
        Err.Clear
    Else
        OK_MES ("GET IFont.SuperScript")    
    End If
Else
    ERROR_MES ("NOT EQUAL IFont.SuperScript")  
End If

'FontStyle
Excel.Cells(7,1).Value = "FontStyle" 
Excel.Cells(7,1).Font.FontStyle = "Bold"  
Excel.Cells(7,2).Value = "FontStyle" 
Excel.Cells(7,2).Font.FontStyle = "Italic"
Excel.Cells(7,3).Value = "FontStyle" 
Excel.Cells(7,3).Font.FontStyle = "Regular"
Excel.Cells(7,4).Value = "FontStyle" 
Excel.Cells(7,4).Font.FontStyle = "Underline"
Excel.Cells(7,5).Value = "FontStyle" 
Excel.Cells(7,5).Font.FontStyle = "Strikeout" 
Excel.Cells(7,6).Value = "FontStyle" 
Excel.Cells(7,6).Font.FontStyle = "полужирный"
Excel.Cells(7,7).Value = "FontStyle" 
Excel.Cells(7,7).Font.FontStyle = "курсив"
Excel.Cells(7,8).Value = "FontStyle" 
Excel.Cells(7,8).Font.FontStyle = "Bold Italic"
Excel.Cells(7,9).Value = "FontStyle" 
Excel.Cells(7,9).Font.FontStyle = "полужирный курсив"
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.FontStyle")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.FontStyle")  
End If


Excel.Cells(7,1).Value = Excel.Cells(7,1).Font.FontStyle  
Excel.Cells(7,2).Value = Excel.Cells(7,2).Font.FontStyle
Excel.Cells(7,3).Value = Excel.Cells(7,3).Font.FontStyle
Excel.Cells(7,4).Value = Excel.Cells(7,4).Font.FontStyle
Excel.Cells(7,5).Value = Excel.Cells(7,5).Font.FontStyle
Excel.Cells(7,6).Value = Excel.Cells(7,6).Font.FontStyle
Excel.Cells(7,7).Value = Excel.Cells(7,7).Font.FontStyle
Excel.Cells(7,8).Value = Excel.Cells(7,8).Font.FontStyle
Excel.Cells(7,9).Value = Excel.Cells(7,9).Font.FontStyle
If Err.Number <> 0 Then
    ERROR_MES ("GET IFont.FontStyle")  
    Err.Clear
Else 
    OK_MES ("GET IFont.FontStyle")  
End If


If Excel.Cells(7,1).Font.Bold=TRUE then
    If  Err.Number <> 0 Then
        ERROR_MES ("GET IFont.Bold")  
        Err.Clear
    Else
        OK_MES ("EQUAL IFont.FontStyle")      
    End If 
Else
    ERROR_MES ("NOT EQUAL IFont.FontStyle")  
End If

'Name 
Excel.Cells(8,1).Value = "Name" 
Excel.Cells(8,1).Font.Name = "Times New Roman" 
Excel.Cells(8,2).Value = "Arial" 
Excel.Cells(8,2).Font.Name = "Arial" 
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.Name")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.Name")  
End If
If Excel.Cells(8,1).Font.Name="Times New Roman" then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IFont.Name")  
        Err.Clear
    Else
        OK_MES ("GET IFont.Name")     
    End If
Else
    ERROR_MES ("NOT EQUAL IFont.Name")  
End If

'Shadow 
Excel.Cells(9,1).Value = "Shadow" 
Excel.Cells(9,1).Font.Shadow = TRUE
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.Shadow")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.Shadow")  
End If
If Excel.Cells(9,1).Font.Shadow=TRUE then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IFont.Shadow")  
        Err.Clear
    Else
        OK_MES ("GET IFont.Shadow")      
    End If
Else
    ERROR_MES ("NOT EQUAL IFont.Shadow")  
End If

'Size
Excel.Cells(10,1).Value = "Size" 
Excel.Cells(10,1).Font.Size = 18    
Excel.Cells(10,2).Value = "Size" 
Excel.Cells(10,2).Font.Size = 14
Excel.Cells(10,3).Value = "Size" 
Excel.Cells(10,3).Font.Size = 10
Excel.Cells(10,4).Value = "Size" 
Excel.Cells(10,4).Font.Size = 8
Excel.Cells(10,5).Value = "Size" 
Excel.Cells(10,5).Font.Size = 6
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.Size")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.Size")  
End If
If Excel.Cells(10,1).Font.Size=18 then
    If Err.Number <> 0 Then
         ERROR_MES ("GET IFont.Size")  
         Err.Clear
    Else
         OK_MES ("GET IFont.Size")     
    End If
Else
    ERROR_MES ("NOT EQUAL IFont.Size")  
End If

'Color
Excel.Cells(11,1).Value = "Color" 
Excel.Cells(11,1).Font.Color = 255
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.Color")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.Color")  
End If
If Excel.Cells(11,1).Font.Color=255 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IFont.Color")  
        Err.Clear
    Else
        OK_MES ("GET IFont.Color")      
    End If
Else
    ERROR_MES ("NOT EQUAL IFont.Color")  
End If

'ColorIndex
Excel.Cells(12,1).Value = "ColorIndex"
Excel.Cells(12,1).Font.ColorIndex = 10
Excel.Cells(12,2).Value = "ColorIndex" 
Excel.Cells(12,2).Font.ColorIndex = 1
Excel.Cells(12,3).Value = "ColorIndex" 
Excel.Cells(12,3).Font.ColorIndex = 52
Excel.Cells(12,4).Value = "ColorIndex" 
Excel.Cells(12,4).Font.ColorIndex = 25
Excel.Cells(12,5).Value = "ColorIndex" 
Excel.Cells(12,5).Font.ColorIndex = 45
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.ColorIndex")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.ColorIndex")  
End If
If Excel.Cells(12,1).Font.ColorIndex=10 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IFont.ColorIndex")  
        Err.Clear
    Else
        OK_MES ("GET IFont.ColorIndex")      
    End If
Else
    ERROR_MES ("NOT EQUAL IFont.ColorIndex")  
End If

'OutlineFont
Excel.Cells(13,1).Value = "OutlineFont" 
Excel.Cells(13,1).Font.OutlineFont = FALSE
If Err.Number <> 0 Then
    ERROR_MES ("PUT IFont.OutlineFont")  
    Err.Clear
Else 
    OK_MES ("PUT IFont.OutlineFont")   
End If
'Такое поведение теста обусловлено тем, что
'даже в windows это свойство ни на что не влияет
If Excel.Cells(13,1).Font.OutlineFont=FALSE then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IFont.OutlineFont")  
        Err.Clear
    Else
       OK_MES ("GET IFont.OutlineFont")       
   End If
Else
    ERROR_MES ("NOT EQUAL IFont.OutlineFont")  
End If

'Background
'Application 
'Creator   
'Parent 

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
