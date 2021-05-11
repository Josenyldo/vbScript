'This script gets the last day in a month
MsgBox "Argumento 1: " & WScript.Arguments.Item(0)
MsgBox "Argumento 2: " & WScript.Arguments.Item(1)

Dim Days

'Select Case WScript.Arguments.Item(0)
'	Case 01,03,05,07,08,10,12 Days=31
'	Case 02 If (((WScript.Arguments.Item(1) Mod 4 = 0) And (WScript.Arguments.Item(1) Mod 100 <> 0)) Or (WScript.Arguments.Item(1) Mod 4 = 0)) Then Days=29 Else Days=28
'	Case 04,06,08,11 Days=30
'End Select

'MsgBox "Dias: " & Days
WScript.StdOut.WriteLine(Days) 'Returns the last day