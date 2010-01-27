'Задача
'Имеем в квике  таблицы 
'-ТВС(табл всех сделк) для акций (TVS_A) 
'-ТВС для фюч, эти данные будем хранить в БД (TVS_F)
'-ТИП (таблица изм параметров) для фуч ГАЗПРОМА (TIP_GAZ)
'-ТИП (таблица изм параметров) для фуч СБЕР (TIP_SBER)
'-ТС (таблица сделок) (TS)
'-TЗ (таблица заявок) (TZ)
'-ТЛБ (табл лимитов по бумаге) (TLB)
'-ТТП (табл тек параметров) (TTP)
'Создадим 4 ДДЕ сервера (далее сервер) с именами
'ddeTVS,ddeGAZ,ddeSBER,ddeTABLES
'в ddeTVS направим данные от TVS_A и TVS_F и в дальнейшем будем хранить в бд
'  и не анализировать,
'  но если хотите можно и анализировать
'в ddeGAZ и ddeSBER данные TIP_GAZ и TIP_SBER и будем их анализировать("писать робота")
'  в бд писать не будем так как у нас уже есть эти данные см выше, 
'  но если хотите можно и хранить в бд
'в ddeTABLES TS,TZ,TLB,TTP и в дальнейшем будем хранить в бд, 
'  и не будем их анализировать,
'  но если хотите можно и анализировать
'Что такое анализировать? 
'Сервер получив данные от квика делает 2 такта:
'1 в зависимости от настройки передает или не передает данные в канал (SELF)
'  ВНЕШНЯЯ(не сервер) программа(далее PIPE_EXE) считывает эти данне и сохраняет их.
'  см dde.vbs   
'2 После передачи данных в канал сервер в параллельном потоке запускает (или не запускает
'  в зависимости от настройки) другую ВНЕШНЮЮ (COM_EXE) программу, если она не была 
'  запущена ранее для анализа (но не записи, так как данные уже переданы на сохранение в шаге 1), 
'  если она была запущена и выполняется то вызов не происходит. 
'  COM_EXE анализируя данные не мешает дальнейшему приему данных (такт 1) для записи 
'  в канал и вызов COM_EXE не происходит так как он работает.
'  Можно провести аналогию вы приняли решение выставить заявку, вы выставляете ее, 
'  и в этот короткий миг времени для вас уже не имеет значения какие идут данные в ТВС
'  так как от вас уже ничего не зависит и заявка выставляется, но это не значит 
'  что данные прекратили свое поступление.
'  

Dim HTA,LOG,YAKOR
Dim IsLOG
Dim fso,WshShell

Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject ("WSCript.shell")

HOST_EXE=Wscript.FullName
VBS=WScript.ScriptFullName
Set VBS_FILE=fso.GetFile(WScript.ScriptFullName)
WHERE_WE=Replace(VBS_FILE.ShortPath,"\" &  VBS_FILE.ShortName,"")
WHERE_QUIK=fso.GetFolder(WHERE_WE).ParentFolder.ShortPath
Set VBS_FILE=Nothing

IsLOG=False

'Проверим, все ли файлы, которые нам нужны, на месте
Set TESTME=CreateObject("Scripting.Dictionary")
TESTME(0)=WHERE_QUIK & "\info.exe" 
TESTME(2)=WHERE_WE   & "\LOG.HTA"
'TESTME(3)=WHERE_WE   & "\TRANS2QUIK.dll"
TESTME(4)=WHERE_WE   & "\FindHTA68.exe"
'TESTME(5)=WHERE_WE   & "\QUIK.mdb"
TESTME(7)=WHERE_WE   & "\dde.exe"
TESTME(8)=WHERE_WE   & "\DDE.wsc"
TESTME(9)=WHERE_WE   & "\DDE.vbs"
TESTME(10)=WHERE_WE   & "\QUIK.udl"
For Each Key In TESTME.Keys
    If Not fso.FileExists(TESTME(Key)) Then
       MsgBox "Нет файла" & vbCrLf & TESTME(Key)
       Wscript.Quit
    End If
Next
Set TESTME=Nothing

'Проверим подключение к БД
On Error Resume Next
Set cnn=CreateObject("ADODB.Connection")
cnn.Open "File Name=" & WHERE_WE & "\QUIK.UDL"
If err.Number<>0 Then
   MsgBox "Не открыть ADODB.Connection" & vbCrLF & "Настройте QUIK.udl"
   Wscript.Quit
End If
cnn.Close
Set cnn=Nothing
On Error GoTo 0
'Запустим ЛОГ окно
WshShell.Run WHERE_WE & "\FINDHTA68.EXE",0,True
Set LOG=Nothing
Set HTA = CreateObject("FindHTA68.HTA68")
Set LOG=HTA.START_HTA(WHERE_WE & "\LOG.HTA", "LOG")
If Not(LOG Is Nothing) Then IsLOG=True
Set YAKOR=LOG.frames("GRID").document.all("YAKOR") 
WriteLOG "Начало работы..."

CreateArxiv("*")
Set P=CreateObject("Scripting.Dictionary")

'Запустим 5 серверов ddeST,ddeTVS,ddeGAZ,ddeSBER,ddeTBL 
'
'ddeST
'после запуска Квика
'Для стканов. На наш взгляд имеет смысл хранить данные стакана
'лучший спрос и лучшее предложение. Для этого настроим окно стакана в квике:
'Вид котировочного окна:^$$^
'Лучшие спрос и предложения видны всегда:Да(V)
'Заголовки столбцов:Покупка,Цена покупки,Цена продажи,Продажа
'В квике для СТКАНА GAZ в настройке вывода дде укажем 
'DDE сервер: ddeST
'Рабочая книга:ST
'Лист:GAZ (для сбера SBER)
'Вывод после создания:Да   [V]
'С заголовками строк:Нет   [ ]
'С заголовками столбцов:Нет[ ]

SN="ddeST"          'имя сервера    
P("SN")=SN
P("RN")=1           'количество строк вывода   
StartDDE "wscript.exe " & WHERE_WE & "\DDE.VBS //D " & SN,P 'StartDDE это польз ф-ия см ниже


'ddeTVS после запуска Квика
'В квике для TVS_A в настройке вывода дде укажем 
'DDE сервер: ddeTVS
'Рабочая книга:TVS
'Лист:A
'Вывод после создания:Да  [V]
'С заголовками строк:Да   [V]
'С заголовками столбцов:Да[V]
'[НАЧАТь ВЫВОД]
'В квике для TVS_F в настройке вывода дде укажем 
'DDE сервер: ddeTVS
'Рабочая книга:TVS
'Лист:F
'Вывод после создания:Да  [V]
'С заголовками строк:Да   [V]
'С заголовками столбцов:Да[V]
'[НАЧАТь ВЫВОД]
SN="ddeTVS"              'имя сервера    
P("SN")=SN
P("RN")=0
P("T4A")="[TVS]F"
P("COM")="script:" & WHERE_WE & "\analys.wsc" 'см dde.wsc
StartDDE "wscript.exe " & WHERE_WE & "\DDE.VBS //D " & SN,P 'StartDDE это польз ф-ия см ниже

'ddeGAZ после запуска Квика
'В квике для TIP_GAZ в настройке вывода дде укажем 
'DDE сервер: ddeGAZ
'Рабочая книга:TIP
'Лист:GAZ
'Вывод после создания:Да  [V]
'С заголовками строк:Да   [V]
'С заголовками столбцов:Да[V]
'[НАЧАТь ВЫВОД]
SN="ddeGAZ"           'имя сервера    
P("SN")=SN   
P("T4A")="[TIP]GAZ"                           '[ИМЯ_ТАБЛИЦЫ]ЛИСТ данные из окна настройки ДДЕ квика которые надо анализировать
P("COM")="script:" & WHERE_WE & "\analys.wsc" 'см dde.wsc
P("P")=""                                     'вывод не нужен
StartDDE "wscript.exe " & WHERE_WE & "\DDE.VBS //D " & SN,P 'StartDDE это польз ф-ия см ниже

'ddeSBER после запуска Квика
'В квике для TIP_SBER в настройке вывода дде укажем 
'DDE сервер: ddeSBER
'Рабочая книга:TIP
'Лист:SBER
'Вывод после создания:Да  [V]
'С заголовками строк:Да   [V]
'С заголовками столбцов:Да[V]
'[НАЧАТь ВЫВОД]
SN="ddeSBER"           'имя сервера    
P("SN")=SN   
P("T4A")="[TIP]SBER"                          '[ИМЯ_ТАБЛИЦЫ]ЛИСТ данные из окна настройки ДДЕ квика которые надо анализировать
P("COM")="script:" & WHERE_WE & "\analys.wsc" 'см dde.wsc
P("P")=""                                     'вывод не нужен
StartDDE "wscript.exe " & WHERE_WE & "\DDE.VBS //D " & SN,P 'StartDDE это польз ф-ия см ниже

'ddeTBL после запуска Квика
'В квике для TC,ТЗ,ТТП,ТЛБ в настройке вывода дде укажем 
'DDE сервер: ddeTBL
'Рабочая книга:TBL
'Лист:TS     (TZ для TZ,TTP для ТТП ,TLB для ТЛБ)    
'Вывод после создания:Да  [V]
'С заголовками строк:Да   [V]
'С заголовками столбцов:Да[V]
'[НАЧАТь ВЫВОД]
P.removeAll
SN="ddeTBL"              'имя сервера    
P("SN")=SN
StartDDE "wscript.exe " & WHERE_WE & "\DDE.VBS //D " & SN,P 'StartDDE это польз ф-ия см ниже

WriteLOG "Запускаем " & WHERE_QUIK & "\info.exe (КВИК)"
PID_QUIK=START_QUIK_EXE 'это польз ф-ия см ниже

'Подождем, когда завершится квик
Set WMI = GetObject("winmgmts:\\.\root\cimv2")
Set colProcesses = WMI.ExecNotificationQuery ("Select * From __InstanceDeletionEvent Within 1 Where TargetInstance ISA 'Win32_Process'")
Do Until i = 999
    Set objProcess = colProcesses.NextEvent
    If objProcess.TargetInstance.ProcessID = PID_QUIK Then
       Exit Do
    End If
Loop

'теперь убьем все DDE.EXE
Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("Select * from Win32_Process Where Name = 'dde.exe'")
For Each DDE_EXE in WMI
    DDE_EXE.Terminate
Next

Wscript.Quit

Function CreateArxiv(what_del)
'Если надо, создадим архивный каталог
what=WHERE_WE & "\ARXIV" 'Путь для сохранения КВИК ТАБЛИЦ
CreateIfNotExists what   'это польз ф-ия см ниже
what=what & "\" & Year(Date)
CreateIfNotExists what
what=what & "\" & Month(Date) 
CreateIfNotExists what
what=what & "\" & Day(Date)
CreateIfNotExists what
'Проверим есть ли там csv файлы
If fso.GetFolder(what).Files.Count > 0 Then
   On Error Resume Next
   fso.DeleteFile what & "\" & what_del & ".CSV"
   If Err.Number <> 53 And Err.Number <> 0 Then
      MsgBox Err.Number & "-" & Err.Description
      Wscript.Quit
   End If
End If
WriteLOG "Подготовили архивный каталог " & what   
End Function

Function CreateIfNotExists(what)
If Not (fso.FolderExists(what)) Then 
   fso.CreateFolder(what)
End If 
End Function

'Запуск ДДЕ сервера
Function StartDDE(WHAT_RUN,P)
'Сформируем строку параметров видк
' /SN=ddeTVS;/RN=1;...
For Each Key In P.Keys
    P_String=P_String &  ";" & Key & "=" & P(Key) 
Next
P_String=Mid(P_String,2)
'Запустим 
WshShell.Run WHAT_RUN & " " & P_String,1,False
WriteLog "Запустили " & WHAT_RUN  & " " & P_String
P.RemoveAll
End Function

'Запуск Квика
Function START_QUIK_EXE
Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("Select * from Win32_Process Where Name = 'info.exe'")
If WMI.Count=0 Then
   WshShell.CurrentDirectory = WHERE_QUIK
   QUIK_CMD = WHERE_QUIK & "\info.exe /i" '& App.Path & "\info.ini"
   START_QUIK_EXE=WshShell.Exec(QUIK_CMD).ProcessID
  Else
   For Each pr In WMI
       START_QUIK_EXE=pr.ProcessId
   Next
   Set pr=Nothing    
End If   
Set WMI=Nothing
End Function

'Запись в лог окно
Function WriteLOG(what)
On Error Resume Next
YAKOR.insertAdjacentText "BeforeEnd","[start.vbs]" & what & vbCrLf 
YAKOR.Document.Body.DoScroll
End Function
