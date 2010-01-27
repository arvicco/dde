Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject ("WSCript.shell")

HOST_EXE=Wscript.FullName
VBS=WScript.ScriptFullName
Set VBS_FILE=fso.GetFile(WScript.ScriptFullName)
WHERE_WE=Replace(VBS_FILE.ShortPath,"\" &  VBS_FILE.ShortName,"")
WHERE_QUIK=fso.GetFolder(WHERE_WE).ParentFolder.ShortPath
Set VBS_FILE=Nothing


Set GearBox=GetObject("script:" & WHERE_WE & "\GearBox.wsc")
GearBox.WhereWe=WHERE_WE 'установим в GearBox путь
On Error Resume Next
fso.DeleteFile GearBox.WhereInQUIK & "\*.*"
On Error GoTo 0

'Определим переменную Dictionary(словарь) где будем хранить параметры команд кпайла
Set P=CreateObject("Scripting.Dictionary")

'Выполним кпайл команду DELETE_ALL_LABELS
'Удалим все метки с графика RI
'Пример(1) выполнения одной строкой, без использования Dictionary переменной
GearBox.RunInQUIK "DELETE_ALL_LABELS","$P1=RI"

'Выполним кпайл команду DELETE_ALL_LABELS
'Удалим все метки с графика SBER
'Пример(2) c использованием Dictionary переменной
cmd="DELETE_ALL_LABELS"
Set P(cmd)=CreateObject("Scripting.Dictionary")
P(cmd)("$P1")="SBER"
RC=GearBox.RunInQUIK(cmd,P(cmd))

'Выполним кпайл команду MESSAGE
cmd="MESSAGE"
Set P(cmd)=CreateObject("Scripting.Dictionary")
P(cmd)("$P1")="Привет Кпайлу от VBS"
P(cmd)("$P2")=1
RC=GearBox.RunInQUIK(cmd,P(cmd))

'Выполним кпайл команду ADD_LABEL (нарисуем метку)
'по цене (ось y)=95250
'по времени (ось x)= на сегодн. дату и время
cmd="ADD_LABEL"
Set P(cmd)=CreateObject("Scripting.Dictionary")
P(cmd)("$P1")="RI"
P(cmd)("TEXT")="Привет метке от VBS" 
P(cmd)("ALIGNMENT")="TOP"
P(cmd)("YVALUE")= 95250
P(cmd)("DATE")=Year(Now()) & FT(Month(Now())) & FT(Day(Now()))   
P(cmd)("TIME")= FT(Hour(Now())) & FT(Minute(Now())) & "00" 
P(cmd)("R")=0
P(cmd)("G")=255
P(cmd)("B")=255
P(cmd)("TRANSPARENCY")=0
P(cmd)("FONT_FACE_NAME")="wingdings"
P(cmd)("FONT_HEIGHT")=20
P(cmd)("IMAGE_PATH")="" 
P(cmd)("HINT")=""
RC=GearBox.RunInQUIK(cmd,P(cmd))

'Выполним кпайл команду GET_INFO_PARAM(VERSION), и покажем результат 
cmd="GET_INFO_PARAM"
Set P(cmd)=CreateObject("Scripting.Dictionary")
P(cmd)("$P1")="VERSION"
RC=GearBox.RunInQUIKA(cmd,P(cmd))
MsgBox P(cmd)("$ANSWER")

'Выполним кпайл команду GET_INFO_PARAM(TRADEDATE), и покажем результат 
P(cmd)("$P1")="TRADEDATE"
RC=GearBox.RunInQUIKA(cmd,P(cmd))
MsgBox P(cmd)("$ANSWER")
Set GearBox=Nothing

Function FT(WHAT) 
FT = WHAT
If WHAT < 10 Then FT = "0" & FT
End Function
