Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject ("WSCript.shell")
DELIMETER=Chr(1)

HOST_EXE=Wscript.FullName
VBS=WScript.ScriptFullName
Set VBS_FILE=fso.GetFile(WScript.ScriptFullName)
WHERE_WE=Replace(VBS_FILE.ShortPath,"\" &  VBS_FILE.ShortName,"")
Set VBS_FILE=Nothing
Set P = CreateObject("Scripting.Dictionary")
P("WHERE_WE")=WHERE_WE
P("SN")=WScript.Arguments.Named.Item("SN")
P("T")=WScript.Arguments.Named.Item("T")

Set ANALYS_WSC=GetObject("script:" & WHERE_WE & "\ANALYS.wsc")
Set ANALYS_WSC.DIC=P

Do 
   WshShell.RegWrite "HKCU\inPIPE\", 1, "REG_DWORD"
   On Error Resume Next
   WHAT=WScript.StdIn.ReadLine
    
   If Err.Number=62 Then Exit Do
   WshShell.RegWrite "HKCU\inPIPE\", 0, "REG_DWORD"
   On Error GoTo 0
   'Запустим анализ
   ANALYS_WSC.GetData4Analysis P("T"), WHAT, WHAT
Loop
Wscript.Quit

