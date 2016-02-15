Option Explicit
Dim fso
Dim oFile
Dim RUTE

RUTE = InputBox("Insert route")

Const Read = 1
Const Write = 2
Const Append = 8


Set fso = CreateObject("Scripting.FileSystemObject")
fso.OpenTextFile RUTE, Read

Set oFile = fso.OpenTextFile(RUTE)
