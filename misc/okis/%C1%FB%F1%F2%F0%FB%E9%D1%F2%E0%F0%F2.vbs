Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject ("WSCript.shell")

HOST_EXE=Wscript.FullName
VBS=WScript.ScriptFullName
Set VBS_FILE=fso.GetFile(WScript.ScriptFullName)
WHERE_WE=Replace(VBS_FILE.ShortPath,"\" &  VBS_FILE.ShortName,"")
WHERE_QUIK=fso.GetFolder(WHERE_WE).ParentFolder.ShortPath
Set VBS_FILE=Nothing

'убьем все DDE.EXE
Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("Select * from Win32_Process Where Name = 'dde.exe'")
For Each DDE_EXE in WMI
    DDE_EXE.Terminate
Next

WshShell.Run "cscript.exe //NOLOGO """ & WHERE_WE & "\dde2Console.vbs""", 1, true

'убьем все DDE.EXE
Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2").ExecQuery("Select * from Win32_Process Where Name = 'dde.exe'")
For Each DDE_EXE in WMI
    DDE_EXE.Terminate
Next

