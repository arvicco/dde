'Не запускайте этот файл самостоятельно
'Он запускается через БыстрыйСтарт.vbs
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject ("WSCript.shell")

HOST_EXE=Wscript.FullName
VBS=WScript.ScriptFullName
Set VBS_FILE=fso.GetFile(WScript.ScriptFullName)
WHERE_WE=Replace(VBS_FILE.ShortPath,"\" &  VBS_FILE.ShortName,"")
WHERE_QUIK=fso.GetFolder(WHERE_WE).ParentFolder.ShortPath
Set VBS_FILE=Nothing


If InStr(1,HOST_EXE,"wscript.exe",1) Then
   MsgBox "Этот скрипт вызывается из БыстрыйСтарт.vbs"
   Wscript.Quit
End If

Wscript.StdOut.WriteLine "Запустите квик..."
Wscript.StdOut.WriteLine "Настройте нужные вам таблицы для вывода ДДЕ..."
Wscript.StdOut.WriteLine "где DDE Server:excel"
Wscript.StdOut.WriteLine "Ждем квик таблицы..."

Set DDE_EXE=WshShell.Exec(WHERE_WE & "\dde.exe")
Do While Not DDE_EXE.StdOut.AtEndOfStream
   WHAT = DDE_EXE.StdOut.ReadLine
   Wscript.StdOut.WriteLine (WHAT)
Loop

Function WHAT_PATH(WHAT)
PATH=Split(WHAT,"\")
For i=0 To Ubound(PATH)-1
    WHAT_PATH= WHAT_PATH & "\" & PATH(i)
Next
WHAT_PATH=Mid(WHAT_PATH,2)
End Function

