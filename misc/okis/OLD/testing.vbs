Set WshShell = CreateObject ("WSCript.shell")
RC=WshShell.Run("ping www.quik.ru", 1, True)
MsgBox Now + RC