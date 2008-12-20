'Проверка записи и чтения из ячейки, ячеек.
'IRange:    Value
'           Formula
'         
'
'
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
otchetFile.WriteLine("++++++ TEST 5 ++++++")
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
    otchetFile.WriteLine("[FAILED] PUT Excel.Workbooks.Add")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT Excel.Workbooks.Add")  
    success = success + 1 
End If


'Запись в ячейку значений через свойство Value 

Excel.Cells(1,1).Value = "Value"
Excel.Cells(1,2).Value = 10
Excel.Cells(1,3).Value = 20
Excel.Cells(1,4).Value = "=B1+C1"
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IRange.Value")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IRange.Value")  
    success = success + 1 
End If

'Проверка чтения через свойство Value
If Excel.Cells(1,1).Value="Value" then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange.Value")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange.Value")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange.Value")  
    failed = failed + 1
End If

If Excel.Cells(1,2).Value=10 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange.Value")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange.Value")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange.Value")  
    failed = failed + 1
End If

If Excel.Cells(1,4).Value=30 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange.Value")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange.Value")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange.Value")  
    failed = failed + 1
End If



'Запись в ячейку значений через свойство Formula 

Excel.Cells(2,1).Value = "Formula"
Excel.Cells(2,2).Value = 10
Excel.Cells(2,3).Value = 20
Excel.Cells(2,4).Value = "=B2+C2"
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IRange.Formula")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IRange.Formula")  
    success = success + 1 
End If

'Проверка чтения через свойство Formula
If Excel.Cells(2,1).Formula="Formula" then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange.Formula")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange.Formula")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange.Formula")  
    failed = failed + 1
End If

If Excel.Cells(2,3).Formula=20 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange.Formula")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange.Formula")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange.Formula")  
    failed = failed + 1
End If

If Excel.Cells(2,4).Formula="=B2+C2" then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange.Formula")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange.Formula")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange.Formula")  
    failed = failed + 1
End If

'++++++++++++++++++++
'Check Value with array
'++++++++++++++++++++

Excel.Cells(2,6) = "Formula"
Excel.Cells(2,7) = 10
Excel.Cells(2,8) = 20
Excel.Cells(2,9) = "=B2+C2"
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IRange = value")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IRange = value")  
    success = success + 1 
End If

'++++++++++++++++++++
'Check Value with array
'++++++++++++++++++++
Dim MyArray(3, 3)

for i=0 to 2
	for j=0 to 2
		MyArray(i, j) = i+j
	next 
next 


Excel.Range("A3:C5").Value = MyArray
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT IRange.Value Two Demension Array")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT IRange.Value Two Demension Array")  
    success = success + 1 
End If

Excel.Range("B6:D8") = MyArray
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] PUT Range =  Two Demension Array")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] PUT Range =  Two Demension Array")  
    success = success + 1 
End If








'++++++   IRange -> EntireRow   

Set tmp = Excel.Range("C3:F4")
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] GET Excel - > Range")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] GET Excel -> Range")  
    success = success + 1 
End If

If tmp.EntireRow.Row=3 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange -> EntireRow")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange -> EntireRow")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange -> EntireRow")  
    failed = failed + 1
End If

'++++++   IRange -> EntireColumn

If tmp.EntireColumn.Column=3 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange -> EntireColumn")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange -> EntireColumn")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange -> EntireColumn")  
    failed = failed + 1
End If

'++++++   IRange -> Columns

If tmp.Columns.Column=3 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange -> Columns")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange -> Columns")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange -> Columns")  
    failed = failed + 1
End If
 
'++++++   IRange -> Rows

If tmp.Rows.Row=3 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange -> Rows")  
        failed = failed + 1
    	Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange -> Rows")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange -> Rows")  
    failed = failed + 1
End If

'++++++   IRange -> Offset 

Set tmp = Excel.Range("C3:F4")
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] GET Excel - > Range")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] GET Excel -> Range")  
    success = success + 1 
End If

Set tmp = tmp.Offset(3)
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] GET Range -> Offset")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] GET Range -> Offset")  
    success = success + 1 
End If


If tmp.Row=6 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange -> Offset")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange -> Offset")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange -> Offset")  
    failed = failed + 1
End If
 
	
Set tmp = tmp.Offset(0, 3)
If Err.Number <> 0 Then
    otchetFile.WriteLine("[FAILED] GET Range -> Offset")  
    failed = failed + 1
    Err.Clear
Else 
    otchetFile.WriteLine("[SUCCESS] GET Range -> Offset")  
    success = success + 1 
End If


If tmp.Column=6 then
    If Err.Number <> 0 Then
        otchetFile.WriteLine("[FAILED] GET IRange -> Offset")  
        failed = failed + 1
    Err.Clear
    Else
        otchetFile.WriteLine("[SUCCESS] GET IRange -> Offset")  
        success = success + 1     
   End If
Else
    otchetFile.WriteLine("[FAILED] NOT EQUAL IRange -> Offset")  
    failed = failed + 1
End If


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'конец кода теста
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Err.Clear

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
