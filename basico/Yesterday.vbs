'Pegar o dia anterior
'Valor Informado: " & WScript.Arguments.Item(0)


dtmYesterday = Date() - 1

WScript.StdOut.WriteLine(dtmYesterday) 'Retorna dia de ontem