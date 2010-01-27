'Этот файл эмуляция dde.exe
'1 Сохранить ранее накопленную из квика SQL таблицу всех сделок или другую в файл
'2 Читать этот файл и а) записывать в базу
'                     б) запускать анализ
Dim LOG
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject ("WSCript.shell")
WshShell.RegWrite "HKCU\inPIPE\", 0, "REG_DWORD"
HOST_EXE=Wscript.FullName
VBS=WScript.ScriptFullName

If InStr(1,HOST_EXE,"wscript.exe",1) Then
   HOST_EXE=Replace(HOST_EXE,"wscript.exe","cscript.exe",1,1,1)
   WScript.CreateObject("WScript.Shell").Run HOST_EXE & " //NOLOGO //D " & VBS,1,False
   Wscript.Quit
End If

Set VBS_FILE=fso.GetFile(WScript.ScriptFullName)
WHERE_WE=Replace(VBS_FILE.ShortPath,"\" &  VBS_FILE.ShortName,"")
WHERE_QUIK=fso.GetFolder(WHERE_WE).ParentFolder.ShortPath
Set VBS_FILE=Nothing
'Запустим ЛОГ
Set HTA = CreateObject("FindHTA68.HTA68")
Set LOG=HTA.START_HTA(WHERE_WE & "\LOG.HTA", "LOG")

DELIMETER=Chr(1)

Set cnn=CreateObject("ADODB.Connection")
Set rst=CreateObject("ADODB.Recordset")
Set rst_out=CreateObject("ADODB.Recordset")
Set cmd=CreateObject("ADODB.Command") 
Set FIELDS = CreateObject("Scripting.Dictionary")

cnn.Open "File Name=" & WHERE_WE & "\QUIK.UDL"
cnn.CursorLocation=3

SN="ddeTVS" 'Имя дде сервера
T="TVS"     'имя файла
L="F"       'имя листа

SQL_TBL=SN & "_" & T & "_" & L
rst.Open SQL_TBL, cnn, 0, 1
For Each fld In rst.Fields
    If (cnn.Execute("SELECT COLUMNPROPERTY( OBJECT_ID('" & SQL_TBL & "'),'" & fld.Name & "','IsIdentity')").Fields(0).Value) + _
       (cnn.Execute("SELECT COLUMNPROPERTY( OBJECT_ID('" & SQL_TBL & "'),'" & fld.Name & "','IsComputed')").Fields(0).Value)=0  Then
       FIELDS(fld.Name)=fld.Name
    End If 
Next
On Error Resume Next 
   fso.DeleteFile SQL_TBL & ".adtg"
   Err.Clear 
   rst.Save SQL_TBL & ".adtg",0 'сохраним SQL_TBL в файл 
   If Err.Number<>0 Then
      MsgBox "Ошибка в момент сохранения файла " &  SQL_TBL & ".adtg" & vbCrLf & _
             "Продолжение невозможно"
      Wscript.Quit
   End If
On Error GoTo 0
rst.Close

rst.Open SQL_TBL & ".adtg", "Provider=MSPersist;", 0, 4, 256
'Теперь так как мы сохранили данные в файле
'можно удалить данные из таблицы, чтоб в нее писать тики из файла
'(Строки закомментированы на всякий случай и данные будут писаться в таблицу "ddeTVS_TVS_F1")
'cnn.Execute "SET ARITHABORT ON",&H80
'cnn.Execute "delete from " & SQL_TBL,&H80
'rst_out.Open SQL_TBL, cnn, 2, 3
cnn.Execute  "delete T",&H80
cnn.Execute  "delete ddeTVS_TVS_F1",&H80
rst_out.Open "ddeTVS_TVS_F1", cnn, 2, 3

L="F1"
Set ANALYS_VBS=WshShell.Exec("cscript.exe //d /nologo " & WHERE_WE & "\analys.vbs /SN:" & SN & " /T:[" & T & "]" & L)

SECOND1=-1
inPIPE=True
Do While Not rst.Eof
   NEWROW="" 
   rst_out.Addnew
   For Each fld In rst.Fields
       NEWROW=NEWROW & DELIMETER & Trim(fld.Value)
       If FIELDS.Exists(fld.Name) Then rst_out.Fields(fld.Name)=fld.Value
   Next 
   rst_out.Update
   NEWROW=Mid(NEWROW,2)
   If SECOND1<>Second(rst("Время").Value) Then
      SECOND1=Second(rst("Время").Value)
      Wscript.Sleep(100)
     Else
      Wscript.Sleep(5)
   End If   
   If  WshShell.RegRead("HKCU\inPIPE\")=1 Then
       ANALYS_VBS.StdIn.WriteLine NEWROW
       WshShell.RegWrite "HKCU\inPIPE\", 0, "REG_DWORD"
   End If
   Wscript.StdOut.WriteLine NEWROW
   rst.MoveNext
Loop

rst_out.Close
rst.Close
