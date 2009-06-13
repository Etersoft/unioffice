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
TEST_NAME ("TEST 5")
On Error Resume Next

'Создаем объект Excel.Application
Set Excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    ERROR_MES("create Excel.Application")  
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

Set WB = Excel.Workbooks.Add
If Err.Number <> 0 Then
    ERROR_MES ("PUT Excel.Workbooks.Add")  
    Err.Clear
Else 
    OK_MES ("PUT Excel.Workbooks.Add")  
End If


'Запись в ячейку значений через свойство Value 

Excel.Cells(1,1).Value = "Value"
Excel.Cells(1,2).Value = 10
Excel.Cells(1,3).Value = 20
Excel.Cells(1,4).Value = "=B1+C1"
If Err.Number <> 0 Then
    ERROR_MES ("PUT IRange.Value")  
    Err.Clear
Else 
    OK_MES ("PUT IRange.Value")  
End If

'Проверка чтения через свойство Value
If Excel.Cells(1,1).Value="Value" then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IRange.Value")  
    Err.Clear
    Else
        OK_MES ("GET IRange.Value")      
   End If
Else
    ERROR_MES ("NOT EQUAL IRange.Value")  
End If

If Excel.Cells(1,2).Value=10 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IRange.Value")  
    Err.Clear
    Else
        OK_MES ("GET IRange.Value")     
   End If
Else
    ERROR_MES ("NOT EQUAL IRange.Value")  
End If

If Excel.Cells(1,4).Value=30 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IRange.Value")  
    Err.Clear
    Else
        OK_MES ("GET IRange.Value")      
   End If
Else
    ERROR_MES ("NOT EQUAL IRange.Value")  
End If



'Запись в ячейку значений через свойство Formula 

Excel.Cells(2,1).Value = "Formula"
Excel.Cells(2,2).Value = 10
Excel.Cells(2,3).Value = 20
Excel.Cells(2,4).Value = "=B2+C2"
If Err.Number <> 0 Then
    ERROR_MES ("PUT IRange.Formula")  
    Err.Clear
Else 
    OK_MES ("PUT IRange.Formula")  
End If

'Проверка чтения через свойство Formula
If Excel.Cells(2,1).Formula="Formula" then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IRange.Formula")  
    Err.Clear
    Else
        OK_MES ("GET IRange.Formula")    
   End If
Else
    ERROR_MES ("NOT EQUAL IRange.Formula")  
End If

If Excel.Cells(2,3).Formula=20 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IRange.Formula")  
    Err.Clear
    Else
        OK_MES ("GET IRange.Formula")   
   End If
Else
    ERROR_MES ("NOT EQUAL IRange.Formula")  
End If

If Excel.Cells(2,4).Formula="=B2+C2" then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IRange.Formula")  
    Err.Clear
    Else
        OK_MES ("GET IRange.Formula")    
   End If
Else
    ERROR_MES ("NOT EQUAL IRange.Formula")  
End If

'++++++++++++++++++++
'Check Value with array
'++++++++++++++++++++

Excel.Cells(2,6) = "Formula"
Excel.Cells(2,7) = 10
Excel.Cells(2,8) = 20
Excel.Cells(2,9) = "=B2+C2"
If Err.Number <> 0 Then
    ERROR_MES ("PUT IRange = value")  
    Err.Clear
Else 
    OK_MES ("PUT IRange = value")   
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
    ERROR_MES ("PUT IRange.Value Two Demension Array")  
    Err.Clear
Else 
    OK_MES ("PUT IRange.Value Two Demension Array")  
End If

Excel.Range("B6:D8") = MyArray
If Err.Number <> 0 Then
    ERROR_MES ("PUT Range =  Two Demension Array")  
    Err.Clear
Else 
    OK_MES ("PUT Range =  Two Demension Array")   
End If





'++++++   IRange -> EntireRow   

Set tmp = Excel.Range("C3:F4")
If Err.Number <> 0 Then
    ERROR_MES ("GET Excel - > Range")  
    Err.Clear
Else 
    OK_MES ("GET Excel -> Range")  
End If

If tmp.EntireRow.Row=3 then
    If Err.Number <> 0 Then
        ERROR_MES ( "GET IRange -> EntireRow")  
    	Err.Clear
    Else
        OK_MES ("GET IRange -> EntireRow")      
   End If
Else
    ERROR_MES ("NOT EQUAL IRange -> EntireRow")  
End If

'++++++   IRange -> EntireColumn

If tmp.EntireColumn.Column=3 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IRange -> EntireColumn")  
    Err.Clear
    Else
        OK_MES ("GET IRange -> EntireColumn")     
   End If
Else
    ERROR_MES ("NOT EQUAL IRange -> EntireColumn")  
End If

'++++++   IRange -> Columns

If tmp.Columns.Column=3 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IRange -> Columns")  
    Err.Clear
    Else
        OK_MES ("GET IRange -> Columns")     
   End If
Else
    ERROR_MES ("NOT EQUAL IRange -> Columns")  
End If
 
'++++++   IRange -> Rows

If tmp.Rows.Row=3 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IRange -> Rows")  
    	Err.Clear
    Else
        OK_MES ("GET IRange -> Rows")      
   End If
Else
    ERROR_MES ("NOT EQUAL IRange -> Rows")  
End If

'++++++   IRange -> Offset 

Set tmp = Excel.Range("C3:F4")
If Err.Number <> 0 Then
    ERROR_MES ("GET Excel - > Range")  
    Err.Clear
Else 
    OK_MES ("GET Excel -> Range")  
End If

Set tmp = tmp.Offset(3)
If Err.Number <> 0 Then
    ERROR_MES ("GET Range -> Offset")  
    Err.Clear
Else 
    OK_MES ("GET Range -> Offset")  
End If


If tmp.Row=6 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IRange -> Offset")  
    Err.Clear
    Else
        OK_MES ("GET IRange -> Offset")      
   End If
Else
    ERROR_MES ("NOT EQUAL IRange -> Offset")  
End If
 
	
Set tmp = tmp.Offset(0, 3)
If Err.Number <> 0 Then
    ERROR_MES ("GET Range -> Offset")  
    Err.Clear
Else 
    OK_MES ("GET Range -> Offset")  
End If


If tmp.Column=6 then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IRange -> Offset")  
    Err.Clear
    Else
        OK_MES ("GET IRange -> Offset")      
   End If
Else
    ERROR_MES ("NOT EQUAL IRange -> Offset")  
End If


'+++++++++++++++ Range->Select()

Excel.Range("F15").Select()
Excel.Range("F15").Value = "TEST"
If Err.Number <> 0 Then
    ERROR_MES ("GET Range -> Select")  
    Err.Clear
Else 
    OK_MES ("GET Range -> Select")  
End If

'+++++++++++++++ Application->Selection()

If Excel.Selection.Value = "TEST" then
    If Err.Number <> 0 Then
        ERROR_MES ("GET IApplication -> Selection")  
    Err.Clear
    Else
        OK_MES ("GET IApplication -> Selection")     
   End If
Else
    ERROR_MES ("NOT EQUAL IApplication -> Selection")  
End If


'+++++++++++ Range[range, range]

Excel.Range(Excel.Cells(19,5), Excel.Cells(20,7)).Select()
If Err.Number <> 0 Then
    ERROR_MES ("GET Range[range, range]")  
    Err.Clear
Else 
    OK_MES ("GET Range[range, range]")   
End If
	

'+++++++++++ Range->Resize(row,col) 

Set tmp = Excel.Range(Excel.Cells(19,1), Excel.Cells(19,1))
tmp.Resize(3,3).Interior.ColorIndex = 12
If Err.Number <> 0 Then
    ERROR_MES ("IRange -> Resize ")  
    Err.Clear
Else 
    OK_MES ("IRange -> Resize")   
End If

Set tmp = Excel.Range(Excel.Cells(22,1), Excel.Cells(25,7))
tmp.Resize(2,2).Interior.ColorIndex = 10
If Err.Number <> 0 Then
    ERROR_MES ("IRange -> Resize ")  
    Err.Clear
Else 
    OK_MES ("IRange -> Resize")  
End If

Set tmp = Excel.Range(Excel.Cells(25,4), Excel.Cells(28,10))
tmp.Resize(2).Interior.ColorIndex = 8
If Err.Number <> 0 Then
    ERROR_MES ("IRange -> Resize ")  
    Err.Clear
Else 
    OK_MES ("IRange -> Resize")   
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
