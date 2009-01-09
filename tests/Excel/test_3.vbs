'++++++++++++++++++++++++++++++++++++++++++++
'test
'     IBorders
'     IBorder
'++++++++++++++++++++++++++++++++++++++++++++

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
TEST_NAME("TEST 3 (IBorders, IBorder)")
On Error Resume Next

'������� ������ Excel.Application
Set Excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    ERROR_MES("create Excel.Application")  
    Err.Clear
Else 
    OK_MES("create Excel.Application")     
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

Set WB = Excel.Workbooks.Add
If Err.Number <> 0 Then
    ERROR_MES ("PUT Excel.Workbooks.Add")  
    Err.Clear
Else 
    OK_MES ("PUT Excel.Workbooks.Add")   
End If

For Each V in Excel.Range("A1:D4").Borders
     V.LineStyle = 1
     If Err.Number <> 0 Then
        ERROR_MES ("PUT Border->LineStyle")  
        Err.Clear
     Else 
        OK_MES ("PUT Border->LineStyle")  
    End If     

     V.ColorIndex = 10
     If Err.Number <> 0 Then
        ERROR_MES ( "PUT Border->ColorIndex")  
        Err.Clear
     Else 
        OK_MES ("PUT Border->ColorIndex)")   
    End If 
Next
If Err.Number <> 0 Then
    ERROR_MES ("Borders (IEnumVARIANT)")  
    Err.Clear
Else 
    OK_MES ("Borders (IEnumVARIANT)")  
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
