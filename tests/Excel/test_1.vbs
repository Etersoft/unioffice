'++++++++++++++++++++++++++++++++++++++++++++++
' Этот модуль тестирует все свойства и методы 
' интерфейса IFont.
'++++++++++++++++++++++++++++++++++++++++++++++

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
otchetFile.WriteLine("++++++ TEST 1 (IFont)++++++")
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

Set WB = Excel.Workbooks.Add
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] create Excel.Workbooks.Add")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] create Excel.Workbooks.Add")  
    success = success + 1 
End If

'IFont.Italic
Excel.Cells(1,1).Value = "Italic"
Excel.Cells(1,1).Font.Italic = TRUE
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IFont.Italic")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.Italic")  
    success = success + 1 
End If
If Excel.Cells(1,1).Font.Italic=TRUE then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.Italic")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IFont.Italic")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.Italic")  
    failed = failed + 1
End If

'IFont.Bold
Excel.Cells(2,1).Value = "Bold"
Excel.Cells(2,1).Font.Bold = TRUE
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IFont.Bold")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.Bold")  
    success = success + 1 
End If
If Excel.Cells(2,1).Font.Bold=TRUE then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.Bold")  
        failed = failed + 1
        Err.Clear
    Else 
        otchetFile.WriteLine("[SUCCESS] GET IFont.Bold")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.Bold")  
    failed = failed + 1
End If

'IFont.Strikethrough
Excel.Cells(3,1).Value = "Strikethrough"
Excel.Cells(3,1).Font.Strikethrough = TRUE
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IFont.Strikethrough")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.Strikethrough")  
    success = success + 1 
End If
If Excel.Cells(3,1).Font.Strikethrough=TRUE then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.Strikethrough")  
        failed = failed + 1
        Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IFont.Strikethrough")  
        success = success + 1     
    End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.Strikethrough")  
    failed = failed + 1
End If

'IFont.Underline
Excel.Cells(4,1).Value = "Underline" 
Excel.Cells(4,1).Font.Underline = 2
Excel.Cells(4,2).Value = "Underline"
Excel.Cells(4,2).Font.Underline = 4		
Excel.Cells(4,3).Value = "Underline" 
Excel.Cells(4,3).Font.Underline = 5
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IFont.Underline")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.Underline")  
    success = success + 1 
End If
If Excel.Cells(4,1).Font.Underline=2 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.Underline")  
        failed = failed + 1
        Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IFont.Underline")  
        success = success + 1     
    End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.Underline")  
    failed = failed + 1
End If

'SubScript 
Excel.Cells(5,1).Value = "SubScript"
Excel.Cells(5,1).Font.SubScript = TRUE
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IFont.SubScript")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.SubScript")  
    success = success + 1 
End If
If Excel.Cells(5,1).Font.SubScript=TRUE then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.SubScript")  
        failed = failed + 1
        Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IFont.SubScript")  
        success = success + 1     
    End If    
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.SubScript")  
    failed = failed + 1
End If

'SuperScript 
Excel.Cells(6,1).Value = "SuperScript"
Excel.Cells(6,1).Font.SuperScript = TRUE
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IFont.SuperScript")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.SuperScript")  
    success = success + 1 
End If
If Excel.Cells(6,1).Font.SuperScript=TRUE then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.SuperScript")  
        failed = failed + 1
        Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IFont.SuperScript")  
        success = success + 1     
    End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.SuperScript")  
    failed = failed + 1
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
    otchetFile.WriteLine("[FAILED] PUT IFont.FontStyle")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.FontStyle")  
    success = success + 1 
End If
If Excel.Cells(7,1).Font.Bold=TRUE then
    If  Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.FontStyle")  
        failed = failed + 1
        Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IFont.FontStyle")  
        success = success + 1     
    End If 
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.FontStyle")  
    failed = failed + 1
End If

'Name 
Excel.Cells(8,1).Value = "Name" 
Excel.Cells(8,1).Font.Name = "Times New Roman" 
Excel.Cells(8,2).Value = "Arial" 
Excel.Cells(8,2).Font.Name = "Arial" 
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IFont.Name")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.Name")  
    success = success + 1 
End If
If Excel.Cells(8,1).Font.Name="Times New Roman" then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.Name")  
        failed = failed + 1
        Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IFont.Name")  
        success = success + 1     
    End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.Name")  
    failed = failed + 1
End If

'Shadow 
Excel.Cells(9,1).Value = "Shadow" 
Excel.Cells(9,1).Font.Shadow = TRUE
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IFont.Shadow")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.Shadow")  
    success = success + 1 
End If
If Excel.Cells(9,1).Font.Shadow=TRUE then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.Shadow")  
        failed = failed + 1
        Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IFont.Shadow")  
        success = success + 1     
    End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.Shadow")  
    failed = failed + 1
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
    otchetFile.WriteLine("[FAILED] PUT IFont.Size")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.Size")  
    success = success + 1 
End If
If Excel.Cells(10,1).Font.Size=18 then
    If Err.Number <> 0 Then
         otchetFile.WriteLine("[FAILED] GET IFont.Size")  
         failed = failed + 1
         Err.Clear
    Else
         otchetFile.WriteLine("[SUCCESS] GET IFont.Size")  
         success = success + 1     
    End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.Size")  
    failed = failed + 1
End If

'Color
Excel.Cells(11,1).Value = "Color" 
Excel.Cells(11,1).Font.Color = 255
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IFont.Color")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.Color")  
    success = success + 1 
End If
If Excel.Cells(11,1).Font.Color=255 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.Color")  
        failed = failed + 1
        Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IFont.Color")  
        success = success + 1     
    End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.Color")  
    failed = failed + 1
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
    otchetFile.WriteLine("[FAILED] PUT IFont.ColorIndex")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.ColorIndex")  
    success = success + 1 
End If
If Excel.Cells(12,1).Font.ColorIndex=10 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.ColorIndex")  
        failed = failed + 1
        Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IFont.ColorIndex")  
        success = success + 1     
    End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.ColorIndex")  
    failed = failed + 1
End If

'OutlineFont
Excel.Cells(13,1).Value = "OutlineFont" 
Excel.Cells(13,1).Font.OutlineFont = FALSE
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IFont.OutlineFont")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IFont.OutlineFont")  
    success = success + 1 
End If
'Такое поведение теста обусловлено тем, что
'даже в windows это свойство ни на что не влияет
If Excel.Cells(13,1).Font.OutlineFont=FALSE then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IFont.OutlineFont")  
        failed = failed + 1
        Err.Clear
    Else
       otchetFile.WriteLine("[SUCCESS] GET IFont.OutlineFont")  
       success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IFont.OutlineFont")  
    failed = failed + 1
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