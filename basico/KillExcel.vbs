'**************************************************************
'*****This Script would close Open Applicatopns*******

Main()
sub Main()
Set oShell=createobject("wscript.shell")
Oshell.Run "Taskkill /f /im excel.exe"
End Sub