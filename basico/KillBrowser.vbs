'**************************************************************
'*****This Script would close Open Applicatopns*******

Main()
sub Main()
Set oShell=createobject("wscript.shell")
Oshell.Run "Taskkill /f /im iexplore.exe*"
Oshell.Run "Taskkill /f /im chrome.exe*"
End Sub