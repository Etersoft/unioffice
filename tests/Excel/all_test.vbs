Dim WshShell
Dim fso
Dim i
Dim success, failed
Dim param

'������ ��� ������� ������ ��������
Set WshShell = Wscript.CreateObject("Wscript.Shell")
'������ ��� ������ � �������
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

i=1
success = 0
failed = 0

if fso.FileExists("otchet.txt") then
    fso.DeleteFile("otchet.txt")
End If
Set otchetFile = fso.CreateTextFile("otchet.txt", true)
otchetFile.Close

filename = "test_" & i & ".vbs"
if fso.FileExists(filename) then
'��������� ��������� �������
    param = i & " " & success & " " & failed
'��������� ������ �� ����������    
    WshShell.Run "cscript.exe /E:vbscript "& filename & " " & param, 0, FALSE
End If
