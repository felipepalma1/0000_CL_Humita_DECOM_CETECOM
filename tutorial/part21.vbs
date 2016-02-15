Option Explicit
Dim fso, oFile, route, a
Const WR = 2
Const R =1
Set fso = CreateObject("Scripting.FileSystemObject")

Set oFile = fso.OpenTextFile("C:\Users\fpalmac\Documents\humita\tutorial\resultzone\hi.txt",WR,true)

For a=1 to 5
  oFile.Write "Testing this file"
  oFile.WriteBlankLines(1)
Next
oFile.Close

Set oFile = fso.OpenTextFile("C:\Users\fpalmac\Documents\humita\tutorial\resultzone\hi.txt", R, true)
MsgBox oFile.ReadAll
