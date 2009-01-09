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

'������ ��� ������� ������ ��������
Set WshShell = Wscript.CreateObject("Wscript.Shell")
'������ ��� ������ � �������
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

'�� ������ ������ Excel ���� ������������ ������.
TEST_NAME ("TEST 6 (ISheets)")
On Error Resume Next

'������� ������ Excel.Application
Set Excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    ERROR_MES ("create Excel.Application")  
    Err.Clear
Else 
    OK_MES ("create Excel.Application")     
End If 

If show_excel then
'���������� Excel
    Excel.Visible = TRUE
    If Err.Number <> 0 Then
        ERROR_MES ("PUT Excel.Visible")  
        Err.Clear
    Else 
        OK_MES ("PUT Excel.Visible")     
    End If 
Else
'�� ���������� Excel. �� ��������� �� � ��� �� ����������, �� ��� �������� �������� ���������
    Excel.Visible = FALSE
    If Err.Number <> 0 Then
        ERROR_MES ("PUT Excel.Visible")  
        Err.Clear
    Else 
        OK_MES ("PUT Excel.Visible")    
    End If 
End If

'��������� ��������������
Excel.DisplayAlerts = FALSE
If Err.Number <> 0 Then
    ERROR_MES ("PUT Excel.DisplayAlerts")  
    Err.Clear
Else 
    OK_MES ("PUT Excel.DisplayAlerts")  
End If

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'����� ���������� ����� ��������� �������
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Excel.SheetsInNewWorkbook = 2
If Err.Number <> 0 Then
    ERROR_MES ("PUT Excel.SheetsInNewWorkbook")  
    Err.Clear
Else 
    OK_MES ("PUT Excel.SheetsInNewWorkbook")  
    success = success + 1 
End If

Set WB = Excel.Workbooks.Add
If Err.Number <> 0 Then
    ERROR_MES ("PUT Excel.Workbooks.Add")  
    Err.Clear
Else 
    OK_MES ("PUT Excel.Workbooks.Add")   
End If


'+++++++   ISheets->Count

Excel.Cells(1,1).Value = "Count"
Excel.Cells(1,2).Value = Excel.Sheets.Count 
If Err.Number <> 0 Then
    ERROR_MES ("ISheets -> Count ")  
    Err.Clear
Else 
    OK_MES ("ISheets -> Count")  
End If

'+++++++   ISheets->Delete

while Excel.Sheets.Count > 1
	Excel.Sheets(Excel.Sheets.Count).Delete()
	If Err.Number <> 0 Then
    		ERROR_MES ("ISheets -> Delete Count = " + Excel.Sheets.Count)  
    		Err.Clear
	Else 
    		OK_MES ("ISheets -> Delete ")  
	End If	
wend

'+++++++ Rename list

Excel.Sheets.Item(1).Name = "���� 1" 
If Err.Number <> 0 Then
    ERROR_MES ("ISheets.item(1).Name ")  
    Err.Clear
Else 
    OK_MES ("ISheets.item(1).Name")  
End If

'+++++++++++++ ISheets -> Add

Set tmp = Excel.Sheets.Add() 
If Err.Number <> 0 Then
    ERROR_MES ("ISheets.Add ")  
    Err.Clear
Else 
    OK_MES ("ISheets.Add")  
End If

Excel.Sheets.Add()
If Err.Number <> 0 Then
    ERROR_MES ("ISheets.Add second list ")  
    Err.Clear
Else 
    OK_MES ("ISheets.Add second list")   
End If


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'����� ���� �����
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

if show_excel then
   MsgBox "���� �������. ��������� � ����� otchet.txt"
End If

'��������� Excel
Excel.Quit
If Err.Number <> 0 Then
    ERROR_MES ("Excel.Quit")  
    Err.Clear
Else 
    OK_MES ("Excel.Quit")  
End If 


otchetFile.WriteLine("����� ������ - " & success+failed & "  ������� - " & success & "  ��������� - " & failed)
otchetFile.Close

'���� ��� ��������� ������, �� ��������� ��������� ����.

if show_excel=false then   
    gsuccess = gsuccess + success
    gfailed = gfailed + failed
    next_script = next_script + 1
    Dim param, filename
'��������� ��������� �������
    param = " " & next_script & " " & gsuccess & " " & gfailed
    filename = "test_" & next_script & ".vbs"
    '��������� ������ �� ����������    
    if fso.FileExists(filename) then 
        WshShell.Run "cscript.exe /E:vbscript "& filename & " " & param, 0, FALSE
    Else
        Set otchetFile = fso.OpenTextFile("otchet.txt", 8)  
        otchetFile.WriteLine("����� ������ - " & gsuccess+gfailed & "  ������� - " & gsuccess & "  ��������� - " & gfailed)
        otchetFile.Close 
        MsgBox "��� ����� ��������. ���������� � ����� otchet.txt"
    End If
End If
