strNomeArquivo = "C:\Users\Josenildo\Documents\Stefanini\Vbs\Leitura.txt"

Dim fso
Set fso = CreateObject("Scripting.Filesystemobject")

If fso.FileExists(strNomeArquivo) Then
  Set obj = fso.OpenTextFile(strNomeArquivo,1,true)
  
   str = obj.ReadLine
   MsgBox(str)
   splitStr = Split(str)
   fiststr = splitStr(0)
   secodStr = splitStr(1)
   fiststr = LCase(fiststr)
   secodStr = LCase(secodStr)
   MsgBox(fiststr+" "+secodStr)
   
Else
 MsgBox("Arquivo não encontrado!")
End If