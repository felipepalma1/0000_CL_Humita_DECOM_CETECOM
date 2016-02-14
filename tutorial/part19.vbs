Option Explicit
Dim fso, oFile, a
Const Read=1, Write=2, Append=8

'oFile.Read(15) Leer las primeras 15 lineas
'oFile.ReadLine
'oFile.ReadAll
'oFile.Close

Set fso = CreateObject("Scripting.FileSystemObject")
Set oFile = fso.OpenTextFile ("resultzone/hi.txt")

For a = 1 to 3
 MsgBox oFile.Read(1)
Next


MsgBox oFile.ReadLine


MsgBox "Done"
