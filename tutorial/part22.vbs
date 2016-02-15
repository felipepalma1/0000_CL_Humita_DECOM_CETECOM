Option Explicit
Dim objFso, oFile
Set objFso = CreateObject("Scripting.FileSystemObject")

Set oFile = objFso.OpenTextFile("C:\Users\fpalmac\Documents\humita\tutorial\resultzone\hi.txt", 8)

oFile.Write "Hello Youtube"
oFile.WriteBlankLines(1)
