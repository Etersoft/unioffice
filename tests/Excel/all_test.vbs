Dim WshShell
Dim fso
Dim i
Dim success, failed
Dim param

'Объект для запуска других скриптов
Set WshShell = Wscript.CreateObject("Wscript.Shell")
'объект для работы с файлами
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
'заполняем параметры запуска
    param = i & " " & success & " " & failed
'запускаем скрипт на исполнение    
    WshShell.Run "cscript.exe /E:vbscript "& filename & " " & param, 0, FALSE
End If
